####################################################################################################
# BSTVC-R Shiny 交互式建模系统
#
# 本文件是一个完整的 Shiny 应用程序，目标是把 BSTVC R 包中 data.check、BSTVC 和 BSVC
# 三个核心工作流封装成“点选式”界面。用户只需要上传建模数据、地图数据，选择响应变量、
# 解释变量、时间字段、空间字段和响应类型，即可完成模型运行，并以表格形式查看和下载
# BSTVC/BSVC 帮助文档中定义的所有模型输出。
#
# 设计原则：
# 1. 只展示模型输出表格，不再展示或下载建模结果图形。
# 2. 输入地图预览仍然保留，因为它属于数据检查，而不是模型结果可视化。
# 3. 代码注释尽量解释“为什么这样做”，帮助后续维护者理解 Shiny 界面和 BSTVC 包参数之间
#    的对应关系。
# 4. 侧边栏沿用用户喜欢的绿色渐变风格，其余布局统一整理成规整的卡片式工作流。
####################################################################################################

####################################################################################################
# 一、依赖包与全局配置
#
# Shiny 默认上传大小偏小，空间面板数据和 shapefile 经常会超过默认限制，因此先把上传上限
# 调整为 1GB。这里把依赖包显式列出并在启动阶段检查，避免用户进入界面后才在某个按钮处
# 遇到“找不到函数”的错误。
####################################################################################################
options(shiny.maxRequestSize = 1024 * 1024^2)

# required_packages <- c(
#   "shiny", "bs4Dash", "DT", "sf", "dplyr", "leaflet",
#   "BSTVC", "openxlsx", "readxl", "spdep", "waiter"
# )
# 
# missing_packages <- required_packages[
#   !vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)
# ]
# 
# if (length(missing_packages) > 0) {
#   stop(
#     sprintf(
#       "缺少必要 R 包：%s。请先安装这些包后再运行应用。",
#       paste(missing_packages, collapse = ", ")
#     ),
#     call. = FALSE
#   )
# }

library(shiny)
library(bs4Dash)
library(DT)
library(sf)
library(dplyr)
library(leaflet)
library(BSTVC)
# 改用 devtools 或 pkgload 动态从本地项目文件夹载入该包
# if (!requireNamespace("devtools", quietly = TRUE)) install.packages("devtools")
# devtools::load_all("INLA")
# devtools::load_all("BSTVC") # 这里的 "BSTVC" 是你刚刚复制进来的文件夹名称
library(openxlsx)
library(readxl)
library(spdep)
library(waiter)


# # 关闭严格的快照验证（防止 INLA 再次卡住）
# options(renv.config.snapshot.validate = FALSE)
# 
# # 显式允许从远程或本地打包
# options(rsconnect.packrat.github = TRUE)
# 
# 
# # 执行部署
# rsconnect::deployApp(
#   appDir = "D:/Codex/example/BSTVC-R Shiny",
#   appPrimaryDoc = "BSTVC_R Shiny 20260526.R",
#   account = "tangxt",
#   server = "shinyapps.io",
#   appName = "BSTVC-Shiny",
#   forceUpdate = TRUE, # 强制覆盖旧配置
#   lint = FALSE
# )
####################################################################################################
# 二、基础工具函数
#
# 本节放置多个通用小函数。它们不直接运行模型，而是解决 Shiny 应用中反复出现的问题：
# 空值兜底、字段名转公式、表格标准化、Excel 工作表命名、图片资源定位等。
####################################################################################################

# `%||%` 用于给 NULL、长度为 0 或全 NA 的对象提供默认值，减少服务器逻辑中的重复判断。
`%||%` <- function(a, b) {
  if (!is.null(a) && length(a) > 0 && !all(is.na(a))) a else b
}

# R 公式中遇到带空格、括号或特殊字符的列名时需要用反引号包起来。该函数统一处理字段名，
# 让用户上传非标准列名时也尽量能构造合法公式。
quote_formula_name <- function(x) {
  paste0("`", gsub("`", "``", x, fixed = TRUE), "`")
}

# BSTVC/BSVC 包返回的结果名称固定写在帮助文档中。这里把它们集中定义，后续 UI 选择器、
# 下载全部结果和测试报告都以此为准。
expected_result_names <- function(model_name) {
  if (identical(model_name, "BSTVC")) {
    return(c(
      "model.evaluation",
      "local.prediction",
      "summary.random.effects",
      "time.coefficients",
      "space.coefficients",
      "STVPI"
    ))
  }

  c(
    "model.evaluation",
    "local.prediction",
    "summary.random.effects",
    "space.coefficients",
    "STVPI"
  )
}

# Excel 的工作表名有 31 个字符长度限制，并且不能包含部分特殊字符。模型结果名目前都合法，
# 但这里仍做清洗，防止以后 BSTVC 包新增结果模块时导出失败。
safe_sheet_name <- function(x) {
  x <- gsub("[:\\\\/?*\\[\\]]", "_", x)
  substr(x, 1, 31)
}

# local.prediction、space.coefficients 等结果通常是 data.frame 或 matrix；个别对象可能是向量、
# 列表或其他结构。DT 和 openxlsx 都更适合处理 data.frame，所以先把所有结果统一转成表格。
normalize_result_df <- function(x) {
  if (is.data.frame(x)) {
    df <- x
  } else if (is.matrix(x)) {
    df <- as.data.frame(x, stringsAsFactors = FALSE)
  } else if (is.atomic(x) || is.factor(x)) {
    df <- data.frame(value = as.vector(x), stringsAsFactors = FALSE)
  } else if (is.list(x)) {
    df <- data.frame(
      content = vapply(
        x,
        function(y) paste(capture.output(str(y, max.level = 2)), collapse = " "),
        character(1)
      ),
      stringsAsFactors = FALSE
    )
  } else {
    df <- data.frame(value = as.character(x), stringsAsFactors = FALSE)
  }

  row_id <- rownames(df)
  if (!is.null(row_id) && !identical(row_id, as.character(seq_len(nrow(df))))) {
    df <- cbind(row_name = row_id, df)
  }
  rownames(df) <- NULL
  df
}

# 将模型结果列表转成“已命名 data.frame 列表”。结果顺序优先遵循帮助文档；如果未来包版本
# 额外返回了新模块，则追加在末尾，避免静默丢失信息。
normalize_result_list <- function(result, model_name) {
  req(result)
  expected <- expected_result_names(model_name)
  ordered_names <- c(intersect(expected, names(result)), setdiff(names(result), expected))
  tables <- lapply(result[ordered_names], normalize_result_df)
  names(tables) <- ordered_names
  tables
}

# 该函数把一个“结果表列表”写成一个 Excel 工作簿。每个模型输出模块对应一个 worksheet，
# 正好满足“一键下载所有表格”的需求。
write_result_workbook <- function(tables, file) {
  wb <- openxlsx::createWorkbook()

  for (nm in names(tables)) {
    sheet <- safe_sheet_name(nm)
    openxlsx::addWorksheet(wb, sheetName = sheet)
    openxlsx::writeData(wb, sheet = sheet, x = tables[[nm]])
    openxlsx::freezePane(wb, sheet = sheet, firstRow = TRUE)
    openxlsx::setColWidths(wb, sheet = sheet, cols = seq_along(tables[[nm]]), widths = "auto")
  }

  openxlsx::saveWorkbook(wb, file = file, overwrite = TRUE)
}

# 当前表下载需要支持 CSV、XLSX、XLS 三种格式。CSV 适合单张表，Excel 格式适合保留更完整的
# 表格结构；这里把写文件逻辑集中在一个函数里，避免两个 downloadHandler 中重复判断格式。
write_single_result_table <- function(df, file, format) {
  if (identical(format, "csv")) {
    utils::write.csv(df, file, row.names = FALSE, fileEncoding = "UTF-8")
    return(invisible(TRUE))
  }

  openxlsx::write.xlsx(df, file, overwrite = TRUE)
  invisible(TRUE)
}

# “下载全部表格”在选择 Excel 格式时写成一个多 worksheet 工作簿；选择 CSV 时则打包成 zip，
# 因为一个 CSV 文件无法自然保存多个结果模块。zip 内每个结果模块对应一张 CSV。
write_all_result_tables <- function(tables, file, format) {
  if (identical(format, "csv")) {
    tmp_dir <- tempfile("bstvc_csv_tables_")
    dir.create(tmp_dir, recursive = TRUE, showWarnings = FALSE)
    on.exit(unlink(tmp_dir, recursive = TRUE, force = TRUE), add = TRUE)

    csv_files <- character(0)
    for (nm in names(tables)) {
      csv_file <- file.path(tmp_dir, paste0(safe_sheet_name(nm), ".csv"))
      utils::write.csv(tables[[nm]], csv_file, row.names = FALSE, fileEncoding = "UTF-8")
      csv_files <- c(csv_files, csv_file)
    }

    old_wd <- getwd()
    on.exit(setwd(old_wd), add = TRUE)
    setwd(tmp_dir)
    utils::zip(zipfile = file, files = basename(csv_files))
    return(invisible(TRUE))
  }

  write_result_workbook(tables, file)
  invisible(TRUE)
}

# 建模页的下拉框使用这个包装函数。shinyWidgets 不再作为必要依赖，减少应用启动门槛；
# 基础 Shiny 的 selectizeInput 已经足够支持搜索和多选。
single_select <- function(input_id, label, choices = NULL, selected = NULL) {
  selectizeInput(
    input_id,
    label,
    choices = choices %||% character(0),
    selected = selected,
    multiple = FALSE,
    options = list(placeholder = "请选择...", allowEmptyOption = TRUE)
  )
}

multi_select <- function(input_id, label, choices = NULL, selected = NULL) {
  selectizeInput(
    input_id,
    label,
    choices = choices %||% character(0),
    selected = selected,
    multiple = TRUE,
    options = list(placeholder = "请选择一个或多个变量")
  )
}

# 首页图片优先从应用目录的 www 文件夹读取；如果用户保留了 D:/Codex/example/Figures，
# 也会自动映射。图片不存在时 UI 会使用图标兜底。
figure_candidates <- unique(c(
  file.path(getwd(), "www"),
  file.path(getwd(), "Figures"),
  file.path(getwd(), "..", "Figures"),
  "D:/Codex/example/BSTVC-R Shiny/www",
  "D:/Codex/example/Figures"
))

existing_figure_dirs <- figure_candidates[dir.exists(figure_candidates)]
figure_dir <- if (length(existing_figure_dirs) > 0) existing_figure_dirs[1] else NA_character_
has_figures <- !is.na(figure_dir) && nzchar(figure_dir)

if (has_figures) {
  shiny::addResourcePath("figs", normalizePath(figure_dir, winslash = "/", mustWork = TRUE))
}

figure_src <- function(filename) {
  if (!has_figures) return(NULL)
  f <- file.path(figure_dir, filename)
  if (!file.exists(f)) return(NULL)
  paste0("figs/", filename)
}

logo_src_1 <- figure_src("R_logo.png")
logo_src_2 <- figure_src("R_logo_pic.png")
card_logo_src <- figure_src("bstvc_card_logo.png")
analysis_banner_src <- figure_src("bg-BSTVC.jpg")

####################################################################################################
# 三、数据读取与检查工具
#
# BSTVC 和 BSVC 对数据顺序非常敏感：数据表中空间单元的排列必须与地图 sf 对象中的几何
# 单元排列一致。这里先可靠读取数据、地图和可选空间权重矩阵，再提供顺序检查函数。
####################################################################################################

# 读取建模数据。CSV 会自动尝试 UTF-8、GB18030/GBK 等常见中文编码；Excel 用 readxl。
# check.names = FALSE 是为了保留用户原始字段名，避免界面中看到的列名和原文件不一致。
normalize_uploaded_text <- function(x) {
  if (!is.character(x)) return(x)

  y <- tryCatch(
    enc2utf8(x),
    error = function(e) iconv(x, from = "", to = "UTF-8", sub = "byte")
  )
  invalid <- !is.na(y) & is.na(iconv(y, from = "UTF-8", to = "UTF-8"))
  if (any(invalid)) {
    y[invalid] <- iconv(x[invalid], from = "", to = "UTF-8", sub = "byte")
  }
  y
}

normalize_uploaded_data_frame <- function(df) {
  df <- as.data.frame(df, stringsAsFactors = FALSE, check.names = FALSE)
  names(df) <- normalize_uploaded_text(names(df))

  for (nm in names(df)) {
    if (is.character(df[[nm]]) || is.factor(df[[nm]])) {
      df[[nm]] <- normalize_uploaded_text(as.character(df[[nm]]))
    }
  }

  df
}

read_csv_with_encoding <- function(path, encoding_label) {
  args <- list(
    file = path,
    stringsAsFactors = FALSE,
    check.names = FALSE,
    comment.char = "",
    na.strings = c("", "NA", "NaN")
  )

  if (!identical(encoding_label, "default")) {
    args$fileEncoding <- encoding_label
  }

  warnings_seen <- character(0)
  df <- withCallingHandlers(
    do.call(utils::read.csv, args),
    warning = function(w) {
      warnings_seen <<- c(warnings_seen, conditionMessage(w))
      invokeRestart("muffleWarning")
    }
  )

  encoding_warning <- grepl(
    "invalid input|invalid multibyte|multi-byte|多字节|编码",
    warnings_seen,
    ignore.case = TRUE
  )
  if (any(encoding_warning)) {
    stop(paste(unique(warnings_seen[encoding_warning]), collapse = "；"), call. = FALSE)
  }

  normalize_uploaded_data_frame(df)
}

safe_read_csv_data <- function(upload_info) {
  encodings <- c("UTF-8", "UTF-8-BOM", "GB18030", "GBK", "CP936", "GB2312", "default")
  errors <- character(0)

  for (enc in encodings) {
    dat <- tryCatch(
      read_csv_with_encoding(upload_info$datapath, enc),
      error = function(e) {
        errors <<- c(errors, sprintf("%s: %s", enc, conditionMessage(e)))
        NULL
      }
    )

    if (!is.null(dat)) {
      if (ncol(dat) == 0) {
        errors <- c(errors, sprintf("%s: 未读取到任何列。", enc))
        next
      }
      if (nrow(dat) == 0) {
        errors <- c(errors, sprintf("%s: 只读取到表头，没有读取到任何数据行；这通常是编码不匹配或 CSV 文件实际为空导致的。", enc))
        next
      }
      if (any(is.na(names(dat))) || any(!nzchar(trimws(names(dat))))) {
        errors <- c(errors, sprintf("%s: 数据中存在空列名，请先在原始文件中补充列名。", enc))
        next
      }
      attr(dat, "bstvc_csv_encoding") <- enc
      return(dat)
    }
  }

  stop(
    paste(
      c(
        "CSV 文件读取失败。请确认文件不是损坏文件、不是只有表头没有数据行，并优先使用 UTF-8 或 GBK/GB18030 编码保存。",
        "系统已自动尝试 UTF-8、UTF-8-BOM、GB18030、GBK、CP936、GB2312 和默认编码。",
        "各编码尝试结果：",
        errors
      ),
      collapse = "\n"
    ),
    call. = FALSE
  )
}

safe_read_data <- function(upload_info) {
  req(upload_info)
  ext <- tolower(tools::file_ext(upload_info$name))

  if (identical(ext, "csv")) {
    return(safe_read_csv_data(upload_info))
  }

  if (ext %in% c("xlsx", "xls")) {
    dat <- as.data.frame(readxl::read_excel(upload_info$datapath), stringsAsFactors = FALSE)
    return(normalize_uploaded_data_frame(dat))
  }

  stop("建模表格数据仅支持 csv、xlsx 或 xls 文件！")
}
# Shiny 上传 shapefile 时，浏览器会把每个文件放到临时路径中；如果只把 .shp 临时文件交给
# sf::st_read()，它在同一临时目录下找不到 .dbf/.shx 等配套文件，地图属性表就会丢失。
# 因此界面允许用户选择 shapefile 文件组，程序内部把这些文件按原文件名复制到同一临时目录，
# 再像本地 R 代码 st_read("NAT.shp") 一样读取主 .shp 文件。

# 识别上传表格更像空间截面宽表还是时空面板长表。用户的列名规则可能不同，因此这里只做
# 温和提示，真正转换仍以用户手动选择的字段和时间序列为准。
detect_panel_data_type <- function(dat) {
  if (is.null(dat) || !is.data.frame(dat) || ncol(dat) == 0) {
    return(list(type = "idle", class = "info", lines = "请先上传并读取待转换的原始数据。"))
  }

  cols <- names(dat)
  lower_cols <- tolower(cols)
  likely_time_cols <- cols[
    lower_cols %in% c("year", "time", "date", "month") |
      grepl("年份|年度|时间|月份|日期", cols)
  ]
  likely_wide_cols <- cols[
    grepl("(^|[^0-9])(18|19|20)[0-9]{2}([^0-9]|$)", cols) |
      grepl("^[0-9]{4}$", cols) |
      grepl("^[0-9]{6}$", cols)
  ]

  if (length(likely_time_cols) == 1 && length(likely_wide_cols) <= 1) {
    return(list(
      type = "panel",
      class = "success",
      lines = c(
        sprintf("数据格式判断：时空面板格式（长表）。检测到一个时间字段：%s。", likely_time_cols[1]),
        "如果原始数据已经是时空面板长表，可直接进入“数据输入”；如需强制转换，仍可在下方设置参数。"
      )
    ))
  }

  if (length(likely_wide_cols) >= 2) {
    return(list(
      type = "wide",
      class = "warning",
      lines = c(
        sprintf("数据格式判断：空间截面格式（宽表，每个时间点是一列）。检测到 %d 个含时间信息的列名。", length(likely_wide_cols)),
        "请在下方按变量选择各自的时间序列列，并确保每个变量的列数与输入的时间序列长度完全一致。"
      )
    ))
  }

  list(
    type = "unknown",
    class = "info",
    lines = c(
      "数据格式判断：暂无法确认。",
      "若每个时间点是一列，请按空间截面宽表继续设置转换参数；若已有唯一时间字段，则可直接进入“数据输入”。"
    )
  )
}

parse_numeric_time_sequence <- function(x) {
  x <- trimws(x %||% "")
  if (!nzchar(x)) {
    stop("请输入时间序列，例如 2000:2021、2000,2001,2002 或 1:12。", call. = FALSE)
  }

  compact <- gsub("\\s+", "", x)
  if (grepl("^-?[0-9]+(\\.[0-9]+)?[:-]-?[0-9]+(\\.[0-9]+)?$", compact)) {
    parts <- strsplit(compact, "[:-]")[[1]]
    start <- as.numeric(parts[1])
    end <- as.numeric(parts[2])
    if (is.na(start) || is.na(end)) {
      stop("时间序列必须能转换为数值型。", call. = FALSE)
    }
    step <- if (start <= end) 1 else -1
    return(seq(start, end, by = step))
  }

  pieces <- strsplit(x, "[,，;；\\s\\n\\r]+")[[1]]
  pieces <- pieces[nzchar(trimws(pieces))]
  values <- suppressWarnings(as.numeric(pieces))
  if (length(values) == 0 || any(is.na(values))) {
    stop("时间序列必须为数值型，请使用英文逗号、空格或冒号范围输入。", call. = FALSE)
  }
  values
}


build_panel_conversion <- function(dat, id_cols, time_values, time_col, specs) {
  if (is.null(dat) || !is.data.frame(dat) || nrow(dat) == 0) {
    stop("请先读取待转换的原始数据。", call. = FALSE)
  }
  if (length(id_cols) == 0) {
    stop("请选择至少一个 id 列。", call. = FALSE)
  }
  if (!all(id_cols %in% names(dat))) {
    stop("部分 id 列不存在于原始数据中。", call. = FALSE)
  }
  if (!nzchar(trimws(time_col))) {
    stop("请输入新时间列的列名。", call. = FALSE)
  }
  if (time_col %in% id_cols) {
    stop("新时间列名不能与 id 列重名。", call. = FALSE)
  }
  if (length(time_values) == 0) {
    stop("时间序列不能为空。", call. = FALSE)
  }

  value_names <- vapply(specs, function(x) trimws(x$value_name %||% ""), character(1))
  if (any(!nzchar(value_names))) {
    stop("每个变量都必须填写转换后的数值列名。", call. = FALSE)
  }
  if (anyDuplicated(value_names) > 0) {
    stop("转换后的数值列名不能重复。", call. = FALSE)
  }
  if (any(value_names %in% c(id_cols, time_col))) {
    stop("转换后的数值列名不能与 id 列或新时间列重名。", call. = FALSE)
  }

  expected_len <- length(time_values)
  for (i in seq_along(specs)) {
    measure_cols <- specs[[i]]$measure_cols %||% character(0)
    value_name <- value_names[i]

    if (length(measure_cols) == 0) {
      stop(sprintf("第 %d 个变量（%s）尚未选择需要转换的列。", i, value_name), call. = FALSE)
    }
    if (!all(measure_cols %in% names(dat))) {
      stop(sprintf("第 %d 个变量（%s）包含原始数据中不存在的列。", i, value_name), call. = FALSE)
    }
    if (length(measure_cols) != expected_len) {
      stop(
        sprintf(
          "第 %d 个变量（%s）的列数为 %d，但时间序列长度为 %d。必须保证每个变量的时间序列数完全一致；请补齐缺失时间点，空值填 NA 后重新上传。",
          i,
          value_name,
          length(measure_cols),
          expected_len
        ),
        call. = FALSE
      )
    }
  }

  row_index <- rep(seq_len(nrow(dat)), times = expected_len)
  converted <- dat[row_index, id_cols, drop = FALSE]
  rownames(converted) <- NULL
  converted[[time_col]] <- rep(time_values, each = nrow(dat))

  for (i in seq_along(specs)) {
    measure_cols <- specs[[i]]$measure_cols
    converted[[value_names[i]]] <- unlist(dat[measure_cols], use.names = FALSE)
  }

  converted
}
safe_read_shp_upload <- function(upload_info) {
  req(upload_info)

  ext <- tolower(tools::file_ext(upload_info$name))
  if (!"shp" %in% ext) {
    stop("建模地图数据必须包含一个 .shp 文件！")
  }

  tmp_dir <- tempfile("uploaded_shapefile_")
  dir.create(tmp_dir, recursive = TRUE, showWarnings = FALSE)

  for (i in seq_len(nrow(upload_info))) {
    file.copy(upload_info$datapath[i], file.path(tmp_dir, upload_info$name[i]), overwrite = TRUE)
  }

  old_env <- Sys.getenv("SHAPE_RESTORE_SHX", unset = NA_character_)
  on.exit({
    if (is.na(old_env)) {
      Sys.unsetenv("SHAPE_RESTORE_SHX")
    } else {
      Sys.setenv(SHAPE_RESTORE_SHX = old_env)
    }
  }, add = TRUE)

  Sys.setenv(SHAPE_RESTORE_SHX = "YES")
  shp_name <- upload_info$name[which(ext == "shp")[1]]
  sf::st_read(file.path(tmp_dir, shp_name), quiet = TRUE)
}

# 用户可以不上传空间权重矩阵；这时 BSTVC/BSVC 会自动用 study_map 构造 QUEEN 邻接的 B 型矩阵。
# 如果用户上传 RDS，则允许 matrix、data.frame、spdep::listw 和 spdep::nb 四类常见对象。
safe_read_matrix_rds <- function(upload_info) {
  if (is.null(upload_info) || is.null(upload_info$name)) return(NULL)

  ext <- tolower(tools::file_ext(upload_info$name))
  if (!identical(ext, "rds")) {
    stop("空间权重矩阵仅支持 .rds 文件！")
  }

  obj <- readRDS(upload_info$datapath)
  mat <- NULL

  if (inherits(obj, "listw")) {
    mat <- spdep::listw2mat(obj)
  } else if (inherits(obj, "nb")) {
    mat <- spdep::nb2mat(obj, style = "B", zero.policy = TRUE)
  } else if (is.matrix(obj)) {
    mat <- obj
  } else if (is.data.frame(obj)) {
    mat <- as.matrix(obj)
  } else {
    stop("不支持的 RDS 对象，请上传 matrix、data.frame、listw 或 nb 对象！")
  }

  storage.mode(mat) <- "numeric"
  mat
}

# 地图属性字段可能因为 shapefile 驱动、编码或大小写差异表现为 FIPS/fips/Fips。
# 这个函数优先寻找完全同名字段；找不到时再做大小写不敏感匹配，并把匹配到的地图字段复制成
# 用户选择的数据字段名，后续 BSTVC::data.check 和模型函数就能拿到一致字段。
align_map_space_field <- function(shp, space_col) {
  map_df <- sf::st_drop_geometry(shp)
  if (space_col %in% names(map_df)) {
    return(shp)
  }

  matched <- names(map_df)[tolower(names(map_df)) == tolower(space_col)]
  if (length(matched) > 0) {
    shp[[space_col]] <- shp[[matched[1]]]
    return(shp)
  }

  stop(sprintf(
    "空间字段不在地图属性表中！数据字段：%s，地图属性字段包括：%s",
    space_col,
    paste(names(map_df), collapse = ", ")
  ))
}

get_space_field <- function(shp, space_col) {
  shp <- align_map_space_field(shp, space_col)
  sf::st_drop_geometry(shp)[[space_col]]
}

# 输入预览和矩阵预览都需要画地图。这里使用 leaflet::addTiles()，避免 addProviderTiles() 对
# leaflet.providers 的额外依赖；即使用户电脑没有安装 leaflet.providers，也能正常显示地图。
# 对普通地图预览，右下角图例说明绿色面代表研究区空间单元；对矩阵地图预览，图例显示数值色带。
render_basic_map <- function(shp, values = NULL, label_prefix = NULL) {
  map_df <- sf::st_drop_geometry(shp)
  label_field <- names(map_df)[1]
  labels <- as.character(map_df[[label_field]])

  if (!is.null(values) && length(values) == nrow(shp)) {
    pal <- leaflet::colorNumeric(c("#f8efe5", "#d8bdca", "#927088"), domain = values, na.color = "#d9d9d9")
    labels <- paste0(labels, if (!is.null(label_prefix)) paste0(" | ", label_prefix, ": ") else " | value: ", round(values, 4))
    return(
      leaflet(shp) %>%
        addTiles() %>%
        addPolygons(
          weight = 1,
          color = "#756b78",
          fillColor = pal(values),
          fillOpacity = 0.72,
          label = labels
        ) %>%
        addLegend("bottomright", pal = pal, values = values, title = label_prefix %||% "value")
    )
  }

  leaflet(shp) %>%
    addTiles() %>%
    addPolygons(weight = 1, color = "#756b78", fillColor = "#d8bdca", fillOpacity = 0.58, label = labels) %>%
    addLegend(
      position = "bottomright",
      colors = "#d8bdca",
      labels = "研究区空间单元",
      title = "地图图例",
      opacity = 0.72
    )
}

####################################################################################################
# 四、建模模块 UI
#
# BSTVC 和 BSVC 的界面高度相似，因此写成同一个 module。BSTVC 需要 Time 和 Space 两个字段；
# BSVC 只需要 Space 字段，并且要求用户直接提供空间截面数据，不在界面中再做时间切片。
# BSTVC 和 BSVC 包函数都提供 threads 参数用于控制并行线程数；界面默认沿用包函数的 threads = 6，
# 并根据当前机器 CPU 核心数限制输入上限，避免用户误填过大的线程数。
####################################################################################################
model_ui <- function(id, model_name = c("BSTVC", "BSVC")) {
  model_name <- match.arg(model_name)
  ns <- NS(id)
  is_bstvc <- identical(model_name, "BSTVC")
  max_threads <- max(1L, parallel::detectCores(logical = TRUE) %||% 6L)
  default_threads <- min(6L, max_threads)

  tagList(
    fluidRow(
      column(
        12,
        bs4Card(
          width = 12,
          status = "info",
          solidHeader = TRUE,
          title = tagList(icon("sliders-h"), paste0(" ", model_name, " 建模参数")),
          fluidRow(
            column(3, single_select(ns("data_source"), "数据来源", choices = c("请选择" = "", "原始数据" = "raw", "检查后数据" = "checked"), selected = "")),
            column(3, single_select(ns("response_var"), "响应变量 Y", choices = NULL)),
            column(3, multi_select(ns("covars"), "解释变量 X", choices = NULL)),
            column(3, single_select(ns("response_type"), "响应类型", choices = c("请选择" = "", "连续型 continuous" = "continuous", "二分类 binary" = "binary", "计数型 count" = "count"), selected = ""))
          ),
          fluidRow(
            if (is_bstvc) {
              column(3, single_select(ns("time_var"), "时间字段 Time", choices = NULL))
            },
            column(3, single_select(ns("space_var"), "空间字段 Space", choices = NULL)),
            column(3, radioButtons(ns("std_switch"), "解释变量标准化", choices = c("否" = "no", "是" = "yes"), selected = "no", inline = TRUE)),
            column(3, radioButtons(ns("spmat_mode"), "空间权重矩阵", choices = c("默认 QUEEN-B" = "default", "自定义 RDS" = "custom"), selected = "default", inline = TRUE))
          ),
          fluidRow(
            column(
              3,
              numericInput(
                ns("threads"),
                "线程数 threads",
                value = default_threads,
                min = 1,
                max = max_threads,
                step = 1
              )
            )
          ),
          if (!is_bstvc) {
            div(
              class = "model-note",
              icon("circle-info"),
              " BSVC 使用空间截面数据：请上传每个空间单元仅出现一次的表格，不需要时间字段！"
            )
          },
          div(
            class = "model-note",
            icon("circle-info"),
            " 数据来源说明：如果上一步数据检查显示顺序已经一致，请选择“原始数据”；如果检查提示已完成自动重排，请选择“检查后数据”！"
          ),
          div(
            class = "model-note",
            icon("circle-info"),
            " 标准化说明：我们强烈推荐对X变量执行标准化处理！原因如下：1.能够确保所有指标在同一尺度上进行比较，有助于衡量各个
变量对模型的贡献大小；2.标准化X变量能够加速模型的计算过程！"
          ),
          actionButton(ns("run_model"), tagList(icon("play"), " 运行模型"), class = "btn btn-success run-btn"),
          br(),
          br(),
          verbatimTextOutput(ns("run_status"), placeholder = TRUE)
        )
      )
    ),
    fluidRow(
      column(
        12,
        bs4Card(
          width = 12,
          status = "info",
          solidHeader = TRUE,
          title = tagList(icon("table"), paste0(" ", model_name, " 输出结果表格")),
          fluidRow(
            column(3, selectInput(ns("result_part"), "结果模块", choices = character(0))),
            column(
              3,
              selectInput(
                ns("download_format"),
                "下载文件类型",
                choices = c("CSV" = "csv", "XLSX" = "xlsx", "XLS" = "xls"),
                selected = "xlsx"
              )
            ),
            column(3, downloadButton(ns("download_table"), "下载当前表", class = "btn-outline-success download-btn result-action-btn")),
            column(3, downloadButton(ns("download_all_tables"), "下载全部表格", class = "btn-success download-btn result-action-btn"))
          ),
          DTOutput(ns("result_table"))
        )
      )
    )
  )
}

####################################################################################################
# 五、建模模块 Server
#
# 这个 module 负责把用户在界面中选择的参数转换成 BSTVC::BSTVC 或 BSTVC::BSVC 的调用。
# 模型运行前会检查数据、地图、变量选择和自定义矩阵是否完整；运行时使用 waiter 遮罩和
# withProgress 阶段条，让用户明确知道应用正在工作。
####################################################################################################
model_server <- function(id, model_name = c("BSTVC", "BSVC"), raw_data_r, checked_data_r, map_r, matrix_r) {
  model_name <- match.arg(model_name)

  moduleServer(id, function(input, output, session) {
    result_r <- reactiveVal(NULL)
    status_r <- reactiveVal("尚未运行模型。")
    tables_r <- reactiveVal(NULL)
    last_base_cols <- reactiveVal(character(0))
    last_cov_signature <- reactiveVal("")

    model_fun <- if (identical(model_name, "BSTVC")) BSTVC::BSTVC else BSTVC::BSVC

    # 用户可选择原始数据或检查后数据。这里不能在“检查后数据”为空时自动退回原始数据：
    # 如果用户主动选择了检查后数据，却还没有完成 data.check，就必须明确提示，否则会让用户
    # 误以为模型使用的是系统处理后的新数据。
    base_data <- reactive({
      if (is.null(input$data_source) || identical(input$data_source, "")) {
        return(NULL)
      }

      if (identical(input$data_source, "checked")) {
        checked_data_r()
      } else {
        raw_data_r()
      }
    })

    model_data <- reactive({
      dat <- base_data()
      if (is.null(dat)) return(NULL)
      dat
    })

    # 只有字段集合真正变化时才刷新基础字段选择器。旧版 observe 每次 input 改变都会重复
    # updateSelectizeInput，导致解释变量多选框持续闪烁。
    observe({
      dat <- base_data()
      if (is.null(dat)) return()

      cols <- names(dat)
      if (identical(cols, last_base_cols())) return()
      last_base_cols(cols)

      updateSelectizeInput(session, "response_var", choices = c("请选择" = "", cols), selected = "", server = TRUE)
      updateSelectizeInput(session, "space_var", choices = c("请选择" = "", cols), selected = "", server = TRUE)

      if (identical(model_name, "BSTVC")) {
        updateSelectizeInput(session, "time_var", choices = c("请选择" = "", cols), selected = "", server = TRUE)
      }
    })

    # 根据已选响应变量、空间字段、时间字段自动排除不能作为解释变量的列，避免用户误把 Y/Time/Space
    # 也放进 ST() 或 S() 中。
    observe({
      dat <- model_data()
      if (is.null(dat)) return()

      blacklist <- c(input$response_var, input$space_var)
      if (identical(model_name, "BSTVC")) {
        blacklist <- c(blacklist, input$time_var)
      }

      cov_choices <- setdiff(names(dat), blacklist[!is.na(blacklist) & nzchar(blacklist)])
      signature <- paste(c(cov_choices, input$response_var, input$space_var, input$time_var %||% ""), collapse = "\r")
      if (identical(signature, last_cov_signature())) return()
      last_cov_signature(signature)

      selected <- intersect(input$covars %||% character(0), cov_choices)

      updateSelectizeInput(session, "covars", choices = cov_choices, selected = selected, server = TRUE)
    })

    observeEvent(input$run_model, {
      dat <- model_data()
      shp <- map_r()

      if (is.null(input$data_source) || identical(input$data_source, "")) {
        status_r("请选择数据来源！")
        return()
      }

      if (identical(input$data_source, "checked") && is.null(checked_data_r())) {
        status_r("你选择了“检查后数据”，但当前没有可用的检查后数据。请先在“数据检查”页面运行对应模型的数据检查，或改选“原始数据”。")
        return()
      }

      if (is.null(dat) || nrow(dat) == 0) {
        status_r("没有可用于建模的数据，请先上传数据，或根据检查结果选择正确的数据来源！")
        return()
      }

      if (is.null(shp)) {
        status_r("没有可用于建模的地图数据，请先上传 shapefile！")
        return()
      }

      if (length(input$covars %||% character(0)) == 0) {
        status_r("请至少选择一个解释变量！")
        return()
      }

      if (is.null(input$response_var) || identical(input$response_var, "")) {
        status_r("请选择响应变量 Y！")
        return()
      }

      if (is.null(input$space_var) || identical(input$space_var, "")) {
        status_r("请选择空间字段 Space！")
        return()
      }

      if (identical(model_name, "BSTVC") && (is.null(input$time_var) || identical(input$time_var, ""))) {
        status_r("请选择时间字段 Time！")
        return()
      }

      if (is.null(input$response_type) || identical(input$response_type, "")) {
        status_r("请选择响应类型！")
        return()
      }

      if (identical(input$spmat_mode, "custom") && is.null(matrix_r())) {
        status_r("已选择自定义空间矩阵，但尚未上传有效 RDS 文件！")
        return()
      }

      waiter_obj <- waiter::Waiter$new(
        html = tagList(
          waiter::spin_fading_circles(),
          tags$h4("Bayesian STVC/SVC模型正在运行，请稍候...")
        ),
        color = "rgba(52, 42, 57, 0.58)"
      )

      waiter_obj$show()
      on.exit(waiter_obj$hide(), add = TRUE)

      tryCatch({
        withProgress(message = paste(model_name, "建模进度"), value = 0, {
          incProgress(0.10, detail = "检查输入数据和字段")

          work_data <- dat
          shp <- align_map_space_field(shp, input$space_var)
          map_df <- sf::st_drop_geometry(shp)

          if (!input$response_var %in% names(work_data)) stop("响应变量不在数据中！")
          if (!input$space_var %in% names(work_data)) stop("空间字段不在数据中！")
          if (!input$space_var %in% names(map_df)) stop("空间字段不在地图属性表中！")
          if (identical(model_name, "BSTVC") && !input$time_var %in% names(work_data)) stop("时间字段不在数据中！")

          incProgress(0.15, detail = "处理解释变量标准化")
          std_note <- "未执行解释变量标准化！"
          if (identical(input$std_switch, "yes")) {
            standardized_covars <- intersect(input$covars, names(work_data))
            numeric_vars <- standardized_covars[vapply(work_data[standardized_covars], is.numeric, logical(1))]
            skipped_vars <- setdiff(standardized_covars, numeric_vars)

            if (length(numeric_vars) > 0) {
              work_data[numeric_vars] <- lapply(work_data[numeric_vars], function(x) as.numeric(scale(x)))
              std_note <- paste0("已标准化：", paste(numeric_vars, collapse = ", "), "。")
            }

            if (length(skipped_vars) > 0) {
              std_note <- paste0(std_note, " 跳过非数值变量：", paste(skipped_vars, collapse = ", "), "。")
            }
          }

          incProgress(0.15, detail = "构造模型公式和参数")
          rhs <- paste(quote_formula_name(input$covars), collapse = " + ")
          response <- quote_formula_name(input$response_var)
          formula_obj <- if (identical(model_name, "BSTVC")) {
            as.formula(sprintf("%s ~ ST(%s)", response, rhs))
          } else {
            as.formula(sprintf("%s ~ S(%s)", response, rhs))
          }

          arg_list <- list(
            formula = formula_obj,
            data = work_data,
            study_map = shp,
            response_type = input$response_type,
            spatial_matrix = if (identical(input$spmat_mode, "custom")) matrix_r() else NULL
          )

          threads_value <- suppressWarnings(as.integer(input$threads))
          if (is.na(threads_value) || threads_value < 1L) {
            stop("线程数 threads 必须是大于等于 1 的整数！")
          }
          arg_list$threads <- threads_value

          if (identical(model_name, "BSTVC")) {
            arg_list$Time <- input$time_var
            arg_list$Space <- input$space_var
          } else {
            arg_list$Space <- input$space_var
          }

          status_r(paste("模型运行中。公式：", deparse(formula_obj)))
          incProgress(0.20, detail = "调用 BSTVC 包执行模型")

          model_res <- do.call(model_fun, arg_list)

          incProgress(0.25, detail = "整理模型输出表格")
          tables <- normalize_result_list(model_res, model_name)
          result_r(model_res)
          tables_r(tables)

          updateSelectInput(
            session,
            "result_part",
            choices = names(tables),
            selected = names(tables)[1]
          )

          missing_expected <- setdiff(expected_result_names(model_name), names(model_res))
          missing_note <- if (length(missing_expected) > 0) {
            paste0(" 未返回的帮助文档结果模块：", paste(missing_expected, collapse = ", "), "。")
          } else {
            ""
          }

          incProgress(0.15, detail = "完成")
          status_r(paste0(
            model_name, " 模型运行完成！", std_note,
            " 当前可查看 ", length(tables), " 个结果表。", missing_note
          ))
        })
      }, error = function(e) {
        status_r(paste("模型运行失败：", conditionMessage(e)))
      })
    }, ignoreInit = TRUE)

    output$run_status <- renderText({
      status_r()
    })

    selected_df <- reactive({
      tables <- tables_r()
      req(tables, input$result_part)
      tables[[input$result_part]]
    })

    output$result_table <- renderDT({
      req(selected_df())
      df <- selected_df()
      numeric_cols <- names(df)[vapply(df, is.numeric, logical(1))]

      dt <- DT::datatable(
        selected_df(),
        rownames = FALSE,
        filter = "top",
        options = list(
          pageLength = 15,
          scrollX = TRUE,
          autoWidth = TRUE,
          dom = "frtip",
          columnDefs = list(list(className = "dt-center", targets = "_all"))
        )
      )

      if (length(numeric_cols) > 0) {
        dt <- DT::formatSignif(dt, columns = numeric_cols)
      }

      dt
    })

    output$download_table <- downloadHandler(
      filename = function() {
        fmt <- input$download_format %||% "xlsx"
        paste0(model_name, "_", input$result_part %||% "result", ".", fmt)
      },
      content = function(file) {
        df <- selected_df()
        req(df)
        write_single_result_table(df, file, input$download_format %||% "xlsx")
      }
    )

    output$download_all_tables <- downloadHandler(
      filename = function() {
        fmt <- input$download_format %||% "xlsx"
        ext <- if (identical(fmt, "csv")) "zip" else fmt
        paste0(model_name, "_all_result_tables.", ext)
      },
      content = function(file) {
        tables <- tables_r()
        req(tables)
        write_all_result_tables(tables, file, input$download_format %||% "xlsx")
      }
    )
  })
}

####################################################################################################
# 六、UI 前端
#
# 侧边栏视觉上只展示“概览、数据输入、数据检查”三个入口；BSTVC/BSVC 建模页由数据检查页的按钮进入。
# bs4Dash 的 updateTabItems() 需要目标 tabName 已经在 bs4SidebarMenu() 中注册，所以这里会保留两个
# 隐藏菜单项作为跳转锚点。用户看不到这两个入口，但按钮仍能稳定切换到对应建模页面。
# 首页只说明建模流程和表格输出，不再提及结果图形展示。侧边栏绿色渐变保持不变，其余组件采用统一
# 间距、圆角、图标和按钮样式，提升整体规整度。
####################################################################################################
ui <- bs4DashPage(
  title = "BSTVC-R Shiny Glass",
  header = bs4DashNavbar(
    skin = "light",
    border = TRUE,
    status = "white",
    title = bs4DashBrand(
      title = tagList(
        if (!is.null(logo_src_1)) {
          tags$img(src = logo_src_1, class = "brand-logo-img", alt = "BSTVC logo")
        },
        tags$strong("BSTVC 交互系统")
      ),
      color = "info"
    )
  ),
  sidebar = bs4DashSidebar(
    skin = "light",
    status = "info",
    elevation = 3,
    bs4SidebarMenu(
      id = "tabs",
      bs4SidebarMenuItem("概览", tabName = "overview", icon = icon("house"), selected = TRUE),
      bs4SidebarMenuItem("空间截面数据转换", tabName = "spatial_convert", icon = icon("shuffle")),
      bs4SidebarMenuItem("数据输入", tabName = "input", icon = icon("database")),
      bs4SidebarMenuItem("数据检查", tabName = "check", icon = icon("list-check")),
      tagAppendAttributes(
        bs4SidebarMenuItem("BSTVC", tabName = "bstvc", icon = icon("chart-line")),
        style = "display:none;"
      ),
      tagAppendAttributes(
        bs4SidebarMenuItem("BSVC", tabName = "bsvc", icon = icon("chart-area")),
        style = "display:none;"
      )
    )
  ),
  controlbar = bs4DashControlbar(collapsed = TRUE),
  footer = bs4DashFooter(left = "BSTVC-R Shiny", right = "© 2026 West China Health and Medical Geography Research Group"),
  body = bs4DashBody(
    waiter::use_waiter(),
    tags$head(
      tags$style(HTML(
        "
        :root {
          --bstvc-bg: #f4f1f2;
          --bstvc-surface: rgba(255,255,255,0.70);
          --bstvc-surface-strong: rgba(255,255,255,0.86);
          --bstvc-border: rgba(72, 58, 69, 0.10);
          --bstvc-divider: rgba(72, 58, 69, 0.13);
          --bstvc-text: #221b2b;
          --bstvc-muted: #756b78;
          --bstvc-accent: #9b7a91;
          --bstvc-accent-2: #c8949f;
          --bstvc-accent-soft: #f0dfe6;
          --bstvc-cream: #f4e7cf;
          --bstvc-lavender: #e4e1f2;
          --bstvc-shadow: 0 18px 44px rgba(55, 45, 58, 0.12);
        }
        body, .content-wrapper, .main-sidebar {
          font-family: 'Segoe UI', 'Microsoft YaHei UI', Arial, sans-serif;
          color: var(--bstvc-text);
        }
        body {
          background:
            radial-gradient(circle at 14% 10%, rgba(244, 231, 207, 0.58), transparent 34%),
            radial-gradient(circle at 84% 8%, rgba(228, 225, 242, 0.58), transparent 30%),
            radial-gradient(circle at 62% 42%, rgba(240, 223, 230, 0.44), transparent 34%),
            linear-gradient(135deg, #fbfaf8 0%, #f3eef1 52%, #eef1f4 100%);
          background-attachment: fixed;
        }
        .content-wrapper {
          background: transparent;
        }
        .content {
          padding-top: 18px;
        }
        .main-header .navbar {
          border-bottom: 1px solid var(--bstvc-divider);
          box-shadow: 0 10px 30px rgba(55, 45, 58, 0.07);
          background: rgba(255, 255, 255, 0.76) !important;
          backdrop-filter: blur(18px);
          -webkit-backdrop-filter: blur(18px);
        }
        .brand-link, .main-sidebar {
          background: rgba(255, 255, 255, 0.64) !important;
          border-right: 1px solid var(--bstvc-divider);
          backdrop-filter: blur(22px);
          -webkit-backdrop-filter: blur(22px);
        }
        .brand-link {
          min-height: 58px;
          display: flex;
          align-items: center;
        }
        .brand-link strong, .brand-text {
          color: var(--bstvc-text) !important;
          font-weight: 700;
          letter-spacing: 0;
        }
        .brand-logo-img {
          width: 26px;
          height: 26px;
          object-fit: contain;
          margin-right: 9px;
          vertical-align: middle;
          filter: drop-shadow(0 5px 10px rgba(72, 58, 69, 0.10));
        }
        .nav-sidebar .nav-link {
          border-radius: 8px;
          color: #5d5461 !important;
          margin: 4px 10px;
          transition: background-color .16s ease, color .16s ease, box-shadow .16s ease;
        }
        .nav-sidebar .nav-link:hover {
          background: rgba(255,255,255,0.66);
          color: #3a303f !important;
        }
        .nav-sidebar .nav-link.active {
          background: rgba(255, 255, 255, 0.82) !important;
          color: #2d2432 !important;
          box-shadow: inset 3px 0 0 var(--bstvc-accent), 0 8px 22px rgba(155,122,145,0.12);
        }
        .card, .info-box, .small-box {
          border-radius: 8px;
          border: 1px solid var(--bstvc-border);
          box-shadow: var(--bstvc-shadow);
          overflow: hidden;
          background: var(--bstvc-surface) !important;
          backdrop-filter: blur(20px);
          -webkit-backdrop-filter: blur(20px);
        }
        .card-header {
          border-bottom: 1px solid var(--bstvc-divider);
          background: rgba(255,255,255,0.68) !important;
          color: var(--bstvc-text) !important;
          min-height: 50px;
          box-shadow: inset 0 -1px 0 rgba(255,255,255,0.34);
        }
        .card-title, .card-header .card-title i {
          color: var(--bstvc-text) !important;
          font-weight: 650;
        }
        .card-tools {
          display: flex;
          align-items: center;
        }
        .card-tools .btn-tool {
          width: 34px;
          height: 34px;
          padding: 0;
          margin-left: 6px;
          display: inline-flex !important;
          align-items: center;
          justify-content: center;
          color: #5e5263 !important;
          background: rgba(255,255,255,0.62) !important;
          border: 1px solid rgba(72,58,69,0.16) !important;
          border-radius: 8px !important;
          opacity: 1 !important;
          box-shadow: none !important;
        }
        .card-tools .btn-tool:hover {
          background: rgba(244,231,207,0.62) !important;
          color: #2d2432 !important;
        }
        .card-tools .btn-tool i,
        .card-tools .btn-tool svg {
          color: currentColor !important;
          font-size: 13px !important;
          opacity: 1 !important;
        }
        .bg-primary, .bg-info, .bg-success, .bg-olive {
          background: rgba(255,255,255,0.62) !important;
          color: var(--bstvc-text) !important;
        }
        .bg-warning {
          background: rgba(244, 231, 207, 0.72) !important;
          color: var(--bstvc-text) !important;
        }
        .small-box {
          min-height: 118px;
          border: 1px solid rgba(72,58,69,0.08) !important;
          background:
            linear-gradient(135deg, rgba(255,255,255,0.76), rgba(240,223,230,0.38)) !important;
          box-shadow: 0 16px 34px rgba(72,58,69,0.08) !important;
          position: relative;
        }
        .small-box::after {
          content: '';
          position: absolute;
          inset: auto 14px 0 14px;
          height: 1px;
          background: rgba(255,255,255,0.72);
        }
        .small-box > .inner {
          padding: 18px 20px !important;
          position: relative;
          z-index: 2;
        }
        .small-box h3 {
          color: #2d2432 !important;
          font-size: 24px !important;
          font-weight: 700 !important;
          margin: 0 0 6px 0 !important;
        }
        .small-box p {
          color: #5d5461 !important;
          font-size: 15px !important;
          font-weight: 600 !important;
          margin: 0 !important;
        }
        .small-box .icon {
          top: 16px !important;
          right: 18px !important;
          color: rgba(72,58,69,0.14) !important;
          z-index: 1;
        }
        .small-box .icon > i,
        .small-box .icon > svg {
          color: rgba(72,58,69,0.14) !important;
          font-size: 70px !important;
        }
        .small-box:hover .icon > i,
        .small-box:hover .icon > svg {
          color: rgba(155,122,145,0.18) !important;
          transform: none !important;
        }
        .btn-success, .btn-primary {
          background: rgba(255, 255, 255, 0.72) !important;
          border-color: rgba(155, 122, 145, 0.32) !important;
          color: #2d2432 !important;
          box-shadow: 0 10px 22px rgba(72,58,69,0.07) !important;
        }
        .btn-success:hover, .btn-primary:hover {
          background: rgba(244, 231, 207, 0.88) !important;
          border-color: rgba(201, 178, 139, 0.56) !important;
          color: #2d2432 !important;
        }
        .btn-outline-success, .btn-outline-primary {
          color: #4c3d51 !important;
          border-color: rgba(155, 122, 145, 0.28) !important;
          background: rgba(255,255,255,0.46) !important;
        }
        .btn-outline-success:hover, .btn-outline-primary:hover {
          background-color: rgba(244, 231, 207, 0.78) !important;
          border-color: rgba(201, 178, 139, 0.52) !important;
          color: #2d2432 !important;
        }
        .btn-success.run-btn, .btn-primary.run-btn {
          background: rgba(244, 231, 207, 0.84) !important;
          border-color: rgba(201, 178, 139, 0.46) !important;
        }
        .btn-success.run-btn:hover, .btn-primary.run-btn:hover {
          background: rgba(239, 219, 181, 0.90) !important;
          border-color: rgba(185, 151, 98, 0.52) !important;
        }
        .progress {
          height: 12px !important;
          border-radius: 999px !important;
          background: rgba(255,255,255,0.60) !important;
          border: 1px solid rgba(72,58,69,0.08);
          box-shadow: inset 0 1px 2px rgba(72,58,69,0.04);
          overflow: hidden;
        }
        .progress-bar {
          background: linear-gradient(90deg, #f6ead2, #d6b673) !important;
          color: #3a303f !important;
          font-size: 11px !important;
          font-weight: 700 !important;
          line-height: 12px !important;
          box-shadow: none !important;
        }
        .btn, .form-control, .selectize-input {
          border-radius: 7px !important;
        }
        .form-control, .selectize-input, .selectize-control.single .selectize-input {
          border-color: rgba(117, 107, 120, 0.24);
          background: rgba(255,255,255,0.78) !important;
          box-shadow: none;
        }
        .form-control:focus, .selectize-input.focus {
          border-color: var(--bstvc-accent);
          box-shadow: 0 0 0 3px rgba(155, 122, 145, 0.13);
        }
        .overview-band {
          background:
            linear-gradient(135deg, rgba(244,231,207,0.82), rgba(228,225,242,0.72)),
            rgba(255,255,255,0.48);
          color: var(--bstvc-text);
          padding: 30px;
          border-radius: 8px;
          min-height: 320px;
          border: 1px solid rgba(255,255,255,0.68);
          box-shadow: var(--bstvc-shadow);
          backdrop-filter: blur(22px);
          -webkit-backdrop-filter: blur(22px);
        }
        .overview-band h2 {
          font-size: 30px;
          font-weight: 700;
          letter-spacing: 0;
          margin-bottom: 14px;
        }
        .overview-band p {
          font-size: 15px;
          line-height: 1.8;
        }
        .overview-logo {
          max-width: 170px;
          border-radius: 8px;
          background: rgba(255,255,255,0.48);
          padding: 12px;
        }
        .overview-user-card {
          min-height: 520px;
        }
        .overview-user-card .card-body {
          padding-top: 64px;
        }
        .overview-user-header {
          min-height: 178px;
          background-size: cover;
          background-position: center;
          border-bottom: 1px solid var(--bstvc-divider);
          position: relative;
        }
        .overview-user-header::after {
          content: '';
          position: absolute;
          inset: 0;
          background: linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0.18));
        }
        .overview-user-card .widget-user-image {
          position: absolute;
          top: 126px;
          left: 50% !important;
          right: auto !important;
          width: 118px;
          margin-left: -59px !important;
          transform: none;
          z-index: 3;
        }
        .overview-user-card .widget-user-image > img {
          width: 118px;
          height: 118px;
          object-fit: contain;
          background: rgba(255,255,255,0.84);
          border: 7px solid rgba(255,255,255,0.78);
          box-shadow: 0 14px 34px rgba(72,58,69,0.18);
        }
        .overview-user-title {
          color: #221b2b;
          font-size: 28px;
          font-weight: 760;
          letter-spacing: 0;
          margin: 18px 0 12px 0;
          text-align: center;
        }
        .overview-user-text {
          color: #453b49;
          font-size: 15px;
          line-height: 1.8;
          margin: 0 auto 12px auto;
          max-width: 86%;
          text-align: center;
        }
        .overview-link-list {
          margin: 18px auto 0 auto;
          max-width: 92%;
          text-align: center;
          white-space: normal;
        }
        .overview-link-list a {
          color: #7a5c70;
          background: transparent;
          border: 0;
          border-radius: 0;
          display: inline-block;
          font-size: 13px;
          font-weight: 620;
          margin: 2px 10px 4px 10px;
          padding: 0;
          text-decoration: underline;
          text-decoration-color: rgba(122,92,112,0.34);
          text-underline-offset: 4px;
        }
        .overview-link-list a:hover {
          color: #2d2432;
          background: transparent;
          text-decoration-color: rgba(45,36,50,0.56);
        }
        .info-block {
          background: rgba(255,255,255,0.62);
          border: 1px solid var(--bstvc-divider);
          border-left: 4px solid var(--bstvc-accent);
          border-radius: 8px;
          padding: 14px 16px;
          margin-bottom: 12px;
        }
        .info-block strong i {
          color: var(--bstvc-accent);
          margin-right: 6px;
        }
        .overview-step .info-box {
          background: rgba(255,255,255,0.62) !important;
          border: 1px solid rgba(72,58,69,0.08) !important;
          box-shadow: 0 14px 32px rgba(72,58,69,0.08) !important;
        }
        .overview-step .info-box-icon {
          color: #2d2432 !important;
          border-radius: 8px !important;
          margin: 10px 0 10px 10px;
          width: 78px !important;
        }
        .overview-step .info-box-icon i {
          color: #2d2432 !important;
        }
        .overview-step-input .info-box-icon {
          background: linear-gradient(135deg, rgba(248,239,216,0.60), rgba(226,191,117,0.28)) !important;
        }
        .overview-step-check .info-box-icon {
          background: linear-gradient(135deg, rgba(245,226,232,0.58), rgba(215,167,184,0.26)) !important;
        }
        .overview-step-model .info-box-icon {
          background: linear-gradient(135deg, rgba(238,232,246,0.58), rgba(197,181,218,0.26)) !important;
        }
        .overview-step .info-box-content {
          padding: 12px 14px !important;
        }
        .overview-step .info-box-text {
          color: #5d5461 !important;
          font-weight: 500 !important;
        }
        .overview-step .info-box-number {
          color: #221b2b !important;
          font-size: 17px !important;
          font-weight: 760 !important;
          margin-top: 5px;
        }
        .model-note, .matrix-note {
          color: #5d5461;
          background: rgba(255,255,255,0.50);
          border: 1px solid rgba(72,58,69,0.08);
          border-radius: 8px;
          padding: 10px 12px;
          line-height: 1.6;
          margin: 8px 0 14px 0;
        }
        .model-note i, .matrix-note i {
          color: var(--bstvc-accent);
        }
        .slice-panel {
          border: 1px solid rgba(72,58,69,0.08);
          border-radius: 8px;
          padding: 12px 14px 2px 14px;
          margin: 8px 0 12px 0;
          background: rgba(255,255,255,0.50);
        }
        .convert-banner {
          background: rgba(255,255,255,0.58);
          border: 1px solid rgba(72,58,69,0.09);
          border-left: 4px solid rgba(155,122,145,0.58);
          border-radius: 8px;
          color: #243044;
          font-weight: 600;
          line-height: 1.8;
          margin-bottom: 14px;
          padding: 12px 16px;
          box-shadow: 0 10px 28px rgba(72,58,69,0.08);
          backdrop-filter: blur(18px);
          -webkit-backdrop-filter: blur(18px);
        }
        .convert-banner i {
          color: #9a6b2f;
          margin-right: 8px;
        }
        .convert-banner strong, .convert-detail summary, .download-tabs .nav-link {
          color: #6d5367;
        }
        .convert-banner-bottom {
          margin-top: 4px;
        }
        .convert-type-alert {
          line-height: 1.7;
          margin-bottom: 0;
        }
        .convert-type-alert h4 {
          margin-bottom: 10px;
        }
        .convert-detail {
          padding: 0;
        }
        .convert-detail summary {
          cursor: pointer;
          font-weight: 700;
          list-style: none;
          padding: 12px 14px;
        }
        .convert-detail summary::-webkit-details-marker {
          display: none;
        }
        .convert-detail summary i {
          color: #7e6276;
          margin-right: 6px;
        }
        .convert-detail-body {
          border-top: 1px solid var(--bstvc-divider);
          padding: 12px 14px 2px 14px;
        }
        .convert-equal-row {
          display: flex;
          flex-wrap: wrap;
          align-items: stretch;
        }
        .convert-equal-row > [class*='col-'] {
          display: flex;
        }
        .convert-equal-row > [class*='col-'] > .card {
          width: 100%;
        }
        .convert-top-card .card-body {
          min-height: 310px;
        }
        .convert-work-row {
          align-items: flex-start;
        }
        .download-tabs .nav-tabs {
          border-bottom: 1px solid var(--bstvc-divider);
          margin-bottom: 12px;
        }
        .download-tabs .nav-link {
          font-weight: 600;
        }
        .download-tabs .nav-link.active {
          color: #2d2432;
          background-color: rgba(240, 223, 230, 0.78);
          border-color: rgba(155,122,145,0.24);
        }
        .card-header .nav-tabs,
        .nav-tabs {
          border-bottom: 1px solid rgba(72,58,69,0.10) !important;
        }
        .card-header .nav-tabs .nav-item,
        .nav-tabs .nav-item {
          margin-bottom: 0;
        }
        .card-header .nav-tabs .nav-link,
        .nav-tabs .nav-link {
          color: #6d5367 !important;
          background: transparent !important;
          border: 1px solid transparent !important;
          border-radius: 7px !important;
          font-weight: 650;
          margin-right: 6px;
          transition: background-color .16s ease, color .16s ease, border-color .16s ease;
        }
        .card-header .nav-tabs .nav-link i,
        .nav-tabs .nav-link i {
          color: currentColor !important;
          opacity: 1 !important;
        }
        .card-header .nav-tabs .nav-link:hover,
        .nav-tabs .nav-link:hover {
          color: #2d2432 !important;
          background: rgba(255,255,255,0.54) !important;
          border-color: rgba(155,122,145,0.16) !important;
        }
        .card-header .nav-tabs .nav-link.active,
        .card-header .nav-tabs .nav-item.show .nav-link,
        .nav-tabs .nav-link.active,
        .nav-tabs .nav-item.show .nav-link {
          color: #2d2432 !important;
          background: rgba(244,231,207,0.74) !important;
          border-color: rgba(155,122,145,0.18) !important;
          box-shadow: 0 8px 18px rgba(72,58,69,0.06);
        }
        .card-header-tabs {
          margin-right: 44px;
        }
        .nav-pills .nav-link {
          color: #6d5367 !important;
          background: transparent !important;
          border: 1px solid transparent !important;
          border-radius: 7px !important;
          font-weight: 650;
          margin-right: 6px;
        }
        .nav-pills .nav-link i {
          color: currentColor !important;
          opacity: 1 !important;
        }
        .nav-pills .nav-link:hover {
          color: #2d2432 !important;
          background: rgba(255,255,255,0.54) !important;
          border-color: rgba(155,122,145,0.16) !important;
        }
        .nav-pills .nav-link.active,
        .nav-pills .show > .nav-link {
          color: #2d2432 !important;
          background: rgba(244,231,207,0.78) !important;
          border-color: rgba(155,122,145,0.18) !important;
          box-shadow: 0 8px 18px rgba(72,58,69,0.06);
        }
        .range-select-row {
          margin-top: 10px;
        }
        .range-select-row .form-group {
          margin-bottom: 8px;
        }
        .range-select-help {
          color: var(--bstvc-muted);
          font-size: 12px;
          line-height: 1.5;
          margin: 4px 0 10px 0;
        }
        .run-btn, .download-btn {
          width: 100%;
          font-weight: 600;
        }
        .alert {
          border-radius: 8px;
          border-width: 1px;
          backdrop-filter: blur(14px);
          -webkit-backdrop-filter: blur(14px);
        }
        .alert-info, .alert-success, .alert-warning, .alert-danger {
          color: #342a39;
          background: rgba(255,255,255,0.58);
          border-color: rgba(72,58,69,0.10);
          box-shadow: none;
        }
        .alert-info {
          border-left: 4px solid rgba(155,122,145,0.44);
        }
        .alert-success {
          border-left: 4px solid rgba(150,166,130,0.58);
        }
        .alert-warning {
          border-left: 4px solid rgba(204,159,87,0.52);
        }
        .alert-danger {
          border-left: 4px solid rgba(190,112,112,0.52);
        }
        .alert h4 {
          color: #2d2432;
          font-weight: 650;
        }
        .alert p:last-child {
          margin-bottom: 0;
        }
        .check-card-fixed .card-body {
          height: 420px;
          overflow-y: auto;
        }
        .check-record-scroll {
          height: 340px;
          overflow-y: auto;
          padding-right: 6px;
        }
        .check-message-text {
          white-space: pre-wrap;
          border: 0;
          background: rgba(255,255,255,0.58);
          color: #342a39;
          font-family: 'Segoe UI', 'Microsoft YaHei UI', Arial, sans-serif;
          font-size: 13px;
          margin-bottom: 0;
          padding: 8px 10px;
        }
        .result-action-btn {
          height: 38px;
          line-height: 24px;
          margin-top: 25px;
          border-radius: 6px;
        }
        table.dataTable {
          background: rgba(255,255,255,0.62);
        }
        table.dataTable thead th {
          background: rgba(244, 231, 207, 0.56) !important;
          color: #332839 !important;
          font-weight: 700 !important;
          border-bottom: 1px solid rgba(72, 58, 69, 0.14) !important;
          text-align: center !important;
        }
        table.dataTable tbody td {
          vertical-align: middle;
        }
        "
      ))
    ),
    bs4TabItems(
      bs4TabItem(
        tabName = "overview",
        fluidRow(
          column(
            5,
            tagAppendAttributes(
              bs4UserCard(
                width = 12,
                status = NULL,
                title = list(
                  tags$div(
                    class = "widget-user-header overview-user-header",
                    style = paste0("background-image: url('", analysis_banner_src %||% "", "');")
                  ),
                  NULL
                ),
                tags$div(
                  class = "widget-user-image",
                  if (!is.null(card_logo_src)) {
                    tags$img(src = card_logo_src, class = "img-circle", alt = "BSTVC card logo")
                  } else if (!is.null(logo_src_1)) {
                    tags$img(src = logo_src_1, class = "img-circle", alt = "BSTVC logo")
                  } else {
                    tags$div(icon("hexagon", class = "fa-4x"))
                  }
                ),
                tags$div(
                  class = "overview-user-body",
                  tags$h2(class = "overview-user-title", "BSTVC-R 交互式建模"),
                  tags$p(
                    class = "overview-user-text",
                    "BSTVC：时空异质视角下影响因素分析、关键驱动因素识别与动态预测的 “全地图” 建模工具，面向时空面板数据与空间截面数据，支持数据类型转换、数据顺序检查、时空变系数建模和空间变系数建模。"
                  ),
                  tags$div(
                    class = "overview-link-list",
                    tags$a(href = "https://github.com/songbi123/BSTVC", target = "_blank", icon("github"), " GitHub: songbi123/BSTVC"),
                    tags$a(href = "https://chaosong.blog/bayesian-stvc/", target = "_blank", icon("book-open"), " Bayesian STVC Profile")
                  )
                )
              ),
              class = "overview-user-card"
            )
          ),
          column(
            7,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("book-open"), " 包核心功能介绍"),
              tags$div(
                class = "info-block",
                tags$strong(icon("clipboard-check"), "data.check"),
                tags$p("检查面板数据或空间截面数据中的空间单元顺序是否与地图一致，并在可自动匹配时重排数据。")
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("chart-line"), "BSTVC"),
                tags$p("用于分析解释变量与响应变量关系在空间和时间维度上的非平稳变化，适合带有重复时间观测的空间面板数据。")
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("chart-area"), "BSVC"),
                tags$p("用于分析解释变量作用在不同空间单元之间的异质性，适合每个空间单元只有一条记录的空间截面数据。")
              )
            )
          )
        ),
        fluidRow(
          column(
            4,
            tags$div(
              class = "overview-step overview-step-input",
              bs4InfoBox("步骤 1", "上传数据、地图和/或空间矩阵", icon = icon("file-import"), color = "info", width = 12)
            )
          ),
          column(
            4,
            tags$div(
              class = "overview-step overview-step-check",
              bs4InfoBox("步骤 2", "运行数据顺序检查", icon = icon("list-check"), color = "warning", width = 12)
            )
          ),
          column(
            4,
            tags$div(
              class = "overview-step overview-step-model",
              bs4InfoBox("步骤 3", "建模并输出结果表格", icon = icon("download"), color = "info", width = 12)
            )
          )
        )
      ),
      bs4TabItem(
        tabName = "spatial_convert",
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner",
              icon("triangle-exclamation"),
              tags$span(
                tags$strong("注意："),
                " 若数据为空间截面类型，则须使用此工具进行时空面板类型的转换！若用户已将原始数据转换为时空面板格式/原始数据已经是时空面板数据格式，则可直接进行下一步“数据输入”！"
              )
            )
          )
        ),
        tags$div(
          class = "row convert-equal-row",
          column(
            4,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              class = "convert-top-card",
              title = tagList(icon("file-import"), " 原始数据读取"),
              fileInput("convert_file", "空间截面原始数据（CSV / Excel）", accept = c(".csv", ".xlsx", ".xls")),
              actionButton("convert_load_file", tagList(icon("folder-open"), " 读取转换数据"), class = "btn btn-success run-btn"),
              tags$div(
                class = "model-note",
                icon("circle-info"),
                "列名规则不作强制限制。请确保同一个变量的各时间列在原始数据中按时间序列顺序排列，或者在选择时按原始列顺序与时间序列一一对应。"
              )
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              class = "convert-top-card",
              title = tagList(icon("table"), " 原始数据"),
              DTOutput("convert_raw_preview")
            )
          )
        ),
        fluidRow(
          column(
            12,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("clipboard-check"), " 检查结果"),
              uiOutput("convert_structure_feedback")
            )
          )
        ),
        fluidRow(
          class = "convert-work-row",
          column(
            4,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("sliders"), " 转换参数"),
              multi_select("convert_id_cols", "id 列（每个时间点不变，可多选）"),
              textAreaInput("convert_time_values", "时间序列（数值型）", value = "2000:2021", rows = 2, placeholder = "例如：2000:2021 或 2000,2001,2002 或 1:12"),
              textInput("convert_time_col", "新时间列名", value = "Year"),
              numericInput("convert_var_count", "需要转换的变量个数", value = 1, min = 1, max = 20, step = 1),
              uiOutput("convert_variable_specs"),
              actionButton("run_panel_convert", tagList(icon("play"), " 开始转换"), class = "btn btn-success run-btn")
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("table-columns"), " 转换结果"),
              DTOutput("converted_panel_preview"),
              tags$hr(),
              uiOutput("convert_result_feedback"),
              tags$div(
                class = "download-tabs",
                tabsetPanel(
                  id = "convert_download_format",
                  type = "tabs",
                  tabPanel("CSV", value = "csv"),
                  tabPanel("Excel xlsx", value = "xlsx"),
                  tabPanel("Excel xls", value = "xls")
                )
              ),
              downloadButton("download_converted_panel", "下载转换后表格", class = "btn btn-outline-success download-btn")
            )
          )
        ),
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner convert-banner-bottom",
              icon("download"),
              tags$span("转换完成后，请下载转换后的数据，并使用该数据进入下一步“数据输入”。"),
              tags$br(),
              icon("list-check"),
              tags$span("时空面板数据格式的转换并不代表空间单元顺序与空间地图的单元顺序一致，故仍需进行后续“数据检查”。")
            )
          )
        )
      ),      bs4TabItem(
        tabName = "input",
        fluidRow(
          column(
            4,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("upload"), " 数据输入"),
              fileInput("data_file", "建模表格数据（CSV / Excel）", accept = c(".csv", ".xlsx", ".xls")),
              fileInput("shp_file", "建模地图数据（shapefile）", accept = c(".shp", ".dbf", ".shx", ".prj", ".cpg"), multiple = TRUE),
              tags$div(
                class = "matrix-note",
                icon("circle-info"),
                " 请同时上传shp文件及其配套文件：.dbf/.shx/.prj/.cpg！"
              ),
              fileInput("sp_matrix_file", "空间权重矩阵（可自定义 .rds）", accept = c(".rds")),
              tags$div(
                class = "matrix-note",
                icon("circle-info"),
                " BSTVC/BSVC 均支持自定义空间权重矩阵，例：k临近、反距离、固定距离等，若用户未输入，此系统也将按默认规则自动构造 QUEEN 邻接 B 型空间权重矩阵！"
              ),
              actionButton("load_resources", tagList(icon("folder-open"), " 读取并验证输入"), class = "btn btn-success run-btn"),
              br(),
              br(),
              uiOutput("load_feedback")
            )
          ),
          column(
            8,
            fluidRow(
              column(4, bs4ValueBox(value = textOutput("n_data_rows", inline = TRUE), subtitle = "数据行数", icon = icon("table-list"), color = "info", width = 12)),
              column(4, bs4ValueBox(value = textOutput("n_data_cols", inline = TRUE), subtitle = "数据列数", icon = icon("table-columns"), color = "info", width = 12)),
              column(4, bs4ValueBox(value = textOutput("n_map_units", inline = TRUE), subtitle = "地图单元数", icon = icon("map"), color = "info", width = 12))
            ),
            bs4TabCard(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("eye"), " 输入预览"),
              id = "input_preview_tabs",
              tabPanel("数据表", icon = icon("table"), DTOutput("data_preview")),
              tabPanel("地图", icon = icon("map-marked-alt"), leafletOutput("map_preview")),
              tabPanel("空间矩阵", icon = icon("border-all"), uiOutput("matrix_preview_ui"))
            )
          )
        )
      ),
      bs4TabItem(
        tabName = "check",
        fluidRow(
          column(
            4,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              class = "check-card-fixed",
              title = tagList(icon("list-check"), " 数据顺序检查"),
              radioButtons("check_mode", "检查模式", choices = c("BSTVC 时空面板" = "bstvc", "BSVC 空间截面" = "bsvc"), selected = "bstvc", inline = FALSE),
              conditionalPanel(
                condition = "input.check_mode == 'bstvc'",
                selectInput("check_time", "时间字段 Time", choices = NULL)
              ),
              selectInput("check_space", "空间字段 Space", choices = NULL),
              actionButton("run_check", tagList(icon("play"), " 运行检查"), class = "btn btn-success run-btn"),
              br(),
              br(),
              tags$div(
                class = "model-note",
                icon("circle-info"),
                " 若检查结果显示数据顺序已经一致，建模页请选择“原始数据”；若检查结果提示系统已完成自动重排，建模页必须选择“检查后数据”。"
              )
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              class = "check-card-fixed",
              title = tagList(icon("clipboard-check"), " 检查结果记录"),
              uiOutput("check_result_ui")
            )
          )
        ),
        fluidRow(
          column(6, actionButton("go_bstvc", tagList(icon("chart-line"), " 进入 BSTVC 建模"), class = "btn btn-success run-btn")),
          column(6, actionButton("go_bsvc", tagList(icon("chart-area"), " 进入 BSVC 建模"), class = "btn btn-outline-success run-btn"))
        )
      ),
      bs4TabItem(tabName = "bstvc", model_ui("bstvc_mod", "BSTVC")),
      bs4TabItem(tabName = "bsvc", model_ui("bsvc_mod", "BSVC"))
    )
  )
)

####################################################################################################
# 七、Server 后端
#
# Server 保存用户上传的数据、地图、空间矩阵和检查后的数据，并把这些响应式对象传给 BSTVC/BSVC
# 建模模块。这里的状态对象尽量保持简单：raw_data 是原始数据，checked_bstvc 和 checked_bsvc
# 分别保存经检查/重排后的建模数据。
####################################################################################################
server <- function(input, output, session) {
  raw_data <- reactiveVal(NULL)
  study_map <- reactiveVal(NULL)
  spatial_matrix <- reactiveVal(NULL)
  checked_bstvc <- reactiveVal(NULL)
  checked_bsvc <- reactiveVal(NULL)
  converter_raw_data <- reactiveVal(NULL)
  converted_panel_data <- reactiveVal(NULL)
  load_status <- reactiveVal(list(type = "info", lines = "请上传数据和地图后点击“读取并验证输入”。"))
  check_state <- reactiveVal(list(code = "idle", message = "等待检查。", mode = "bstvc"))
  check_history <- reactiveVal(list())

  output$n_data_rows <- renderText({ if (is.null(raw_data())) "-" else as.character(nrow(raw_data())) })
  output$n_data_cols <- renderText({ if (is.null(raw_data())) "-" else as.character(ncol(raw_data())) })
  output$n_map_units <- renderText({ if (is.null(study_map())) "-" else as.character(nrow(study_map())) })

  # 点击读取按钮后才更新响应式数据，避免用户选择文件过程中重复触发大量读取操作。

  convert_status <- reactiveVal(list(
    type = "info",
    lines = "请上传空间截面原始数据后点击“读取转换数据”。"
  ))
  convert_result_status <- reactiveVal(list(
    type = "info",
    lines = "尚未执行转换。"
  ))

  observeEvent(input$convert_load_file, {
    if (is.null(input$convert_file)) {
      convert_status(list(type = "warning", lines = "请先选择 CSV、xlsx 或 xls 原始数据文件。"))
      return()
    }

    tryCatch({
      dat <- safe_read_data(input$convert_file)
      converter_raw_data(dat)
      converted_panel_data(NULL)
      convert_result_status(list(type = "info", lines = "尚未执行转换。"))
      detected <- detect_panel_data_type(dat)
      convert_status(list(
        type = detected$class,
        lines = c(
          sprintf("原始数据读取成功：%d 行，%d 列。", nrow(dat), ncol(dat)),
          detected$lines
        )
      ))
    }, error = function(e) {
      converter_raw_data(NULL)
      converted_panel_data(NULL)
      convert_result_status(list(type = "info", lines = "尚未执行转换。"))
      convert_status(list(type = "danger", lines = paste("原始数据读取失败：", conditionMessage(e))))
    })
  }, ignoreInit = TRUE)

  observe({
    dat <- converter_raw_data()
    cols <- if (is.null(dat)) character(0) else names(dat)
    updateSelectizeInput(session, "convert_id_cols", choices = cols, selected = input$convert_id_cols %||% character(0))

    n_vars <- input$convert_var_count %||% 1
    range_choices <- c("请选择" = "", cols)
    for (i in seq_len(n_vars)) {
      updateSelectizeInput(
        session,
        paste0("convert_measure_", i),
        choices = cols,
        selected = input[[paste0("convert_measure_", i)]] %||% character(0)
      )
      updateSelectInput(
        session,
        paste0("convert_measure_start_", i),
        choices = range_choices,
        selected = input[[paste0("convert_measure_start_", i)]] %||% ""
      )
      updateSelectInput(
        session,
        paste0("convert_measure_end_", i),
        choices = range_choices,
        selected = input[[paste0("convert_measure_end_", i)]] %||% ""
      )
    }
  })


  for (range_i in seq_len(20)) {
    local({
      i <- range_i
      observeEvent(input[[paste0("convert_apply_measure_range_", i)]], {
        dat <- converter_raw_data()
        cols <- if (is.null(dat)) character(0) else names(dat)
        start_col <- input[[paste0("convert_measure_start_", i)]] %||% ""
        end_col <- input[[paste0("convert_measure_end_", i)]] %||% ""

        if (length(cols) == 0) {
          showNotification("请先读取原始数据，再按范围选择列。", type = "warning")
          return()
        }
        if (!nzchar(start_col) || !nzchar(end_col)) {
          showNotification("请选择起始列和终止列。", type = "warning")
          return()
        }
        if (!start_col %in% cols || !end_col %in% cols) {
          showNotification("起始列或终止列不在原始数据中。", type = "error")
          return()
        }

        start_pos <- match(start_col, cols)
        end_pos <- match(end_col, cols)
        selected_cols <- cols[seq(min(start_pos, end_pos), max(start_pos, end_pos))]

        updateSelectizeInput(session, paste0("convert_measure_", i), selected = selected_cols)
        showNotification(
          paste0("已按原始数据列顺序选中 ", length(selected_cols), " 列。"),
          type = "message"
        )
      }, ignoreInit = TRUE)
    })
  }
  output$convert_variable_specs <- renderUI({
    dat <- converter_raw_data()
    cols <- if (is.null(dat)) character(0) else names(dat)
    n_vars <- input$convert_var_count %||% 1

    tagList(lapply(seq_len(n_vars), function(i) {
      do.call(
        tags$details,
        c(
          list(class = "slice-panel convert-detail"),
          if (i == 1) list(open = "open") else list(),
          list(
            tags$summary(icon("chevron-right"), paste0("变量 ", i)),
            tags$div(
              class = "convert-detail-body",
              textInput(paste0("convert_value_name_", i), "转换后的数值列名", value = if (i == 1) "Value" else paste0("Value", i)),
              multi_select(paste0("convert_measure_", i), "该变量对应的时间序列列", choices = cols),
              tags$div(
                class = "range-select-row",
                fluidRow(
                  column(6, selectInput(paste0("convert_measure_start_", i), "起始列", choices = c("请选择" = "", cols))),
                  column(6, selectInput(paste0("convert_measure_end_", i), "终止列", choices = c("请选择" = "", cols)))
                ),
                actionButton(paste0("convert_apply_measure_range_", i), tagList(icon("wand-magic-sparkles"), " 按范围选择"), class = "btn btn-outline-success run-btn"),
                tags$div(class = "range-select-help", "选择起始列和终止列后，系统会按原始数据中的列顺序一次性选中这一区间内的全部列。")
              )
            )
          )
        )
      )
    }))
  })

  output$convert_raw_preview <- renderDT({
    dat <- converter_raw_data()
    req(dat)
    DT::datatable(dat, rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })

  output$convert_structure_feedback <- renderUI({
    dat <- converter_raw_data()
    if (is.null(dat)) {
      st <- convert_status()
      return(div(
        class = paste("alert convert-type-alert", paste0("alert-", st$type)),
        tags$h4(icon("clipboard-check"), " 数据类型检查"),
        lapply(st$lines, tags$p)
      ))
    }

    detected <- detect_panel_data_type(dat)
    div(
      class = paste("alert convert-type-alert", paste0("alert-", detected$class)),
      tags$h4(icon("clipboard-check"), " 数据类型检查"),
      lapply(detected$lines, tags$p)
    )
  })

  observeEvent(input$run_panel_convert, {
    dat <- converter_raw_data()

    tryCatch({
      withProgress(message = "空间截面数据转换进度", value = 0, {
        incProgress(0.15, detail = "检查转换参数")
        time_values <- parse_numeric_time_sequence(input$convert_time_values)
        n_vars <- input$convert_var_count %||% 1
        specs <- lapply(seq_len(n_vars), function(i) {
          list(
            value_name = input[[paste0("convert_value_name_", i)]],
            measure_cols = input[[paste0("convert_measure_", i)]] %||% character(0)
          )
        })

        incProgress(0.35, detail = "检查各变量时间序列长度")
        converted <- build_panel_conversion(
          dat = dat,
          id_cols = input$convert_id_cols %||% character(0),
          time_values = time_values,
          time_col = input$convert_time_col,
          specs = specs
        )

        incProgress(0.30, detail = "生成时空面板表格")
        converted_panel_data(converted)
        incProgress(0.20, detail = "完成")
      })

      convert_result_status(list(
        type = "success",
        lines = c(
          sprintf("转换完成：已生成 %d 行，%d 列的时空面板数据。", nrow(converted_panel_data()), ncol(converted_panel_data())),
          "请下载转换后的数据，并使用该数据进行下一步“数据输入”。",
          "注意：格式转换不代表空间单元顺序与空间地图单元顺序一致，仍需进行后续“数据检查”。"
        )
      ))
    }, error = function(e) {
      converted_panel_data(NULL)
      convert_result_status(list(type = "danger", lines = conditionMessage(e)))
    })
  }, ignoreInit = TRUE)

  output$converted_panel_preview <- renderDT({
    dat <- converted_panel_data()
    req(dat)
    DT::datatable(dat, rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })


  output$convert_result_feedback <- renderUI({
    st <- convert_result_status()
    div(
      class = paste("alert", paste0("alert-", st$type)),
      lapply(st$lines, tags$p)
    )
  })

  output$download_converted_panel <- downloadHandler(
    filename = function() {
      fmt <- input$convert_download_format %||% "csv"
      paste0("BSTVC_panel_converted_", format(Sys.Date(), "%Y%m%d"), ".", fmt)
    },
    content = function(file) {
      dat <- converted_panel_data()
      req(dat)
      write_single_result_table(dat, file, input$convert_download_format %||% "csv")
    }
  )
  observeEvent(input$load_resources, {
    msgs <- character(0)
    has_error <- FALSE

    if (!is.null(input$data_file)) {
      tryCatch({
        dat <- safe_read_data(input$data_file)
        raw_data(dat)
        checked_bstvc(NULL)
        checked_bsvc(NULL)
        msgs <- c(msgs, sprintf("[OK] 数据读取成功！"))
      }, error = function(e) {
        has_error <<- TRUE
        msgs <<- c(msgs, paste("[FAIL] 数据读取失败：", conditionMessage(e)))
      })
    } else {
      msgs <- c(msgs, "[INFO] 未更新建模数据！")
    }

    if (!is.null(input$shp_file)) {
      tryCatch({
        shp <- safe_read_shp_upload(input$shp_file)
        study_map(shp)
        checked_bstvc(NULL)
        checked_bsvc(NULL)
        msgs <- c(msgs, sprintf("[OK] 地图读取成功！"))
      }, error = function(e) {
        has_error <<- TRUE
        msgs <<- c(msgs, paste("[FAIL] 地图读取失败：", conditionMessage(e)))
      })
    } else {
      msgs <- c(msgs, "[INFO] 未更新地图数据！")
    }

    if (!is.null(input$sp_matrix_file)) {
      tryCatch({
        mat <- safe_read_matrix_rds(input$sp_matrix_file)
        spatial_matrix(mat)
        msgs <- c(msgs, sprintf("[OK] 空间矩阵读取成功：%d x %d！", nrow(mat), ncol(mat)))
      }, error = function(e) {
        has_error <<- TRUE
        spatial_matrix(NULL)
        msgs <<- c(msgs, paste("[FAIL] 空间矩阵读取失败：", conditionMessage(e)))
      })
    } else {
      spatial_matrix(NULL)
      msgs <- c(msgs, "[INFO] 未上传自定义矩阵，建模时将使用包默认矩阵！")
    }

    load_status(list(type = if (has_error) "danger" else "success", lines = msgs))
  }, ignoreInit = TRUE)

  output$load_feedback <- renderUI({
    st <- load_status()
    div(
      class = paste("alert", paste0("alert-", st$type)),
      lapply(st$lines, tags$p)
    )
  })

  output$data_preview <- renderDT({
    dat <- raw_data()
    req(dat)
    DT::datatable(dat, rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })

  output$map_preview <- renderLeaflet({
    shp <- study_map()
    req(shp)
    render_basic_map(shp)
  })

  output$matrix_preview_ui <- renderUI({
    mat <- spatial_matrix()
    if (is.null(mat)) {
      return(tags$div(class = "alert alert-info", "未上传自定义空间矩阵；建模时将使用 BSTVC/BSVC 默认矩阵！"))
    }
    tagList(
      tags$h5("空间矩阵预览（前 20 x 20）"),
      DTOutput("matrix_preview"),
      tags$hr(),
      tags$h5("空间矩阵地图预览（按每个空间单元的邻居数量着色）"),
      leafletOutput("matrix_map_preview")
    )
  })

  output$matrix_preview <- renderDT({
    mat <- spatial_matrix()
    req(mat)
    preview <- as.data.frame(mat)
    preview <- preview[seq_len(min(20, nrow(preview))), seq_len(min(20, ncol(preview))), drop = FALSE]
    DT::datatable(preview, rownames = FALSE, options = list(dom = "t", scrollX = TRUE))
  })

  output$matrix_map_preview <- renderLeaflet({
    mat <- spatial_matrix()
    shp <- study_map()
    req(mat, shp)

    if (nrow(mat) != nrow(shp)) {
      return(
        leaflet() %>%
          addTiles() %>%
          addControl(
            html = "空间矩阵行数与地图空间单元数不一致，无法生成矩阵地图预览！",
            position = "topright"
          )
      )
    }

    neighbor_count <- rowSums(mat != 0, na.rm = TRUE)
    render_basic_map(shp, values = neighbor_count, label_prefix = "邻居数量")
  })

  # 数据列名更新后，同步刷新数据检查页中的字段选择器。BSTVC 默认优先选择 Year/FIPS，
  # 这是 Florida 示例数据和帮助文档中的字段名。
  observe({
    dat <- raw_data()
    if (is.null(dat)) return()

    cols <- names(dat)
    updateSelectInput(session, "check_time", choices = c("请选择" = "", cols), selected = input$check_time %||% "")
    updateSelectInput(session, "check_space", choices = c("请选择" = "", cols), selected = input$check_space %||% "")
  })

  observeEvent(input$run_check, {
    dat <- raw_data()
    shp <- study_map()

    if (is.null(dat) || is.null(shp)) {
      st <- list(code = "error", message = "请先读取建模表格数据和建模地图数据！", mode = input$check_mode, time = format(Sys.time(), "%H:%M:%S"))
      check_state(st)
      check_history(c(list(st), check_history()))
      return()
    }

    map_df <- sf::st_drop_geometry(shp)

    tryCatch({
      if (is.null(input$check_space) || identical(input$check_space, "")) stop("请选择空间字段 Space！")
      if (!input$check_space %in% names(dat)) stop("空间字段不在数据中！")
      shp <- align_map_space_field(shp, input$check_space)
      map_df <- sf::st_drop_geometry(shp)

      if (identical(input$check_mode, "bstvc")) {
        if (is.null(input$check_time) || identical(input$check_time, "")) stop("请选择时间字段 Time！")
        if (!input$check_time %in% names(dat)) stop("时间字段不在数据中！")

        data_check_output <- capture.output({
          corrected <- BSTVC::data.check(data = dat, study_map = shp, Time = input$check_time, Space = input$check_space)
        })

        if (is.null(corrected) || !is.data.frame(corrected)) {
          checked_bstvc(NULL)
          stop(paste(c("data.check 未返回可用于建模的数据！", data_check_output), collapse = "\n"))
        }

        checked_bstvc(corrected)
        checked_bsvc(NULL)

        check_message <- paste(c("BSTVC::data.check 已完成检查！", data_check_output), collapse = "\n")
        st <- list(
          code = if (identical(corrected, dat)) "ok_aligned" else "ok_rematched",
          mode = "bstvc",
          message = check_message,
          time = format(Sys.time(), "%H:%M:%S")
        )
        check_state(st)
        check_history(c(list(st), check_history()))
      } else {
        if (anyDuplicated(dat[[input$check_space]]) > 0) {
          stop("BSVC 需要单期空间截面数据，即每个空间单元只能出现一次！")
        }

        temp_time <- ".bstvc_shiny_bsvc_time"
        while (temp_time %in% names(dat)) {
          temp_time <- paste0(".", temp_time)
        }

        check_dat <- dat
        check_dat[[temp_time]] <- 1L

        data_check_output <- capture.output({
          corrected <- BSTVC::data.check(data = check_dat, study_map = shp, Time = temp_time, Space = input$check_space)
        })

        if (is.null(corrected) || !is.data.frame(corrected)) {
          checked_bsvc(NULL)
          stop(paste(c("data.check 未返回可用于建模的数据！", data_check_output), collapse = "\n"))
        }

        corrected[[temp_time]] <- NULL
        checked_bsvc(corrected)
        checked_bstvc(NULL)

        check_message <- paste(c("BSVC 空间截面数据已通过临时时间字段调用 BSTVC::data.check 完成检查！", data_check_output), collapse = "\n")
        st <- list(
          code = if (identical(corrected, dat)) "ok_aligned" else "ok_rematched",
          mode = "bsvc",
          message = check_message,
          time = format(Sys.time(), "%H:%M:%S")
        )
        check_state(st)
        check_history(c(list(st), check_history()))
      }
    }, error = function(e) {
      st <- list(code = "error", mode = input$check_mode, message = conditionMessage(e), time = format(Sys.time(), "%H:%M:%S"))
      check_state(st)
      check_history(c(list(st), check_history()))
    })
  }, ignoreInit = TRUE)

  output$check_result_ui <- renderUI({
    history <- check_history()

    if (length(history) == 0) {
      return(div(class = "alert alert-info", tags$h4(icon("info-circle"), " 等待检查"), tags$p(check_state()$message)))
    }

    div(class = "check-record-scroll", tagList(lapply(history, function(st) {
      cls <- switch(
        st$code,
        error = "alert alert-danger",
        ok_rematched = "alert alert-warning",
        ok_aligned = "alert alert-success",
        "alert alert-info"
      )

      title <- switch(
        st$code,
        error = "检查失败",
        ok_rematched = "已完成自动重排",
        ok_aligned = "检查通过",
        "检查记录"
      )

      div(
        class = cls,
        style = "margin-bottom:10px;",
        tags$h4(icon(if (identical(st$code, "error")) "exclamation-triangle" else "check-circle"), paste0(" ", title)),
        tags$p(tags$strong(paste0("[", toupper(st$mode), "] ", st$time %||% ""))),
        tags$pre(class = "check-message-text", st$message)
      )
    })))
  })

  observeEvent(input$go_bstvc, {
    updateTabItems(session, inputId = "tabs", selected = "bstvc")
  })

  observeEvent(input$go_bsvc, {
    updateTabItems(session, inputId = "tabs", selected = "bsvc")
  })

  model_server(
    id = "bstvc_mod",
    model_name = "BSTVC",
    raw_data_r = raw_data,
    checked_data_r = checked_bstvc,
    map_r = study_map,
    matrix_r = spatial_matrix
  )

  model_server(
    id = "bsvc_mod",
    model_name = "BSVC",
    raw_data_r = raw_data,
    checked_data_r = checked_bsvc,
    map_r = study_map,
    matrix_r = spatial_matrix
  )
}

####################################################################################################
# 八、应用入口
#
# shinyApp 会把上方定义的 ui 和 server 组合成可运行应用。用户可以在 RStudio 中直接运行本文件，
# 也可以通过 shiny::runApp("文件所在目录") 启动。
####################################################################################################
shinyApp(ui, server)
