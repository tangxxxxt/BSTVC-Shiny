####################################################################################################
# BSTVC-R Shiny 交互式建模系统
#
# 本文件是一个完整的 Shiny 应用程序，目标是把 BSTVC R 包中 data.check、BSTVC 和 BSVC
# 三个核心工作流封装成“点选式”界面。用户只需要上传建模数据、Map数据，选择响应变量、
# 解释变量、时间字段、空间字段和Response Type，即可Done模型运行，并以表格形式查看和下载
# BSTVC/BSVC 帮助文档中定义的所有模型输出。
#
# 设计原则：
# 1. 只展示模型输出表格，不再展示或下载建模结果图形。
# 2. 输入Map预览仍然保留，因为它属于Data Check，而不是模型结果可视化。
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
#       "Missing required R packages: %s. Please install these packages before running the app.",
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
#   appDir = "<local app directory>",
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
# 空值兜底、Field Name转公式、表格标准化、Excel 工作表命名、图片资源定位等。
####################################################################################################

# `%||%` 用于给 NULL、长度为 0 或全 NA 的对象提供默认值，减少服务器逻辑中的重复判断。
`%||%` <- function(a, b) {
  if (!is.null(a) && length(a) > 0 && !all(is.na(a))) a else b
}

# R 公式中遇到带空格、括号或特殊字符的列名时需要用反引号包起来。该函数统一处理Field Name，
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
# 但这里仍做清洗，防止以后 BSTVC 包新增Result Module时导出失败。
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
  names(tables)[names(tables) == "STVPI"] <- "contribution.STVPI"
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

write_field_converted_table <- function(df, file, format) {
  if (!identical(format, "csv")) {
    return(write_single_result_table(df, file, format))
  }

  classes <- vapply(df, function(x) class(x)[1], character(1))
  schema <- paste(
    paste(utils::URLencode(names(classes), reserved = TRUE), classes, sep = "="),
    collapse = ";"
  )
  con <- file(file, open = "w", encoding = "UTF-8")
  on.exit(close(con), add = TRUE)
  writeLines(paste0("# BSTVC_FIELD_TYPES:", schema), con, useBytes = TRUE)
  utils::write.table(
    df,
    con,
    sep = ",",
    row.names = FALSE,
    col.names = TRUE,
    quote = TRUE,
    na = "",
    qmethod = "double",
    fileEncoding = "UTF-8"
  )
  invisible(TRUE)
}

# “Download All Tables”在选择 Excel 格式时写成一个多 worksheet 工作簿；选择 CSV 时则打包成 zip，
# 因为一个 CSV 文件无法自然保存多个Result Module。zip 内每个Result Module对应一张 CSV。
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
    options = list(placeholder = "Please select...", allowEmptyOption = TRUE)
  )
}

multi_select <- function(input_id, label, choices = NULL, selected = NULL) {
  selectizeInput(
    input_id,
    label,
    choices = choices %||% character(0),
    selected = selected,
    multiple = TRUE,
    options = list(placeholder = "Select one or more variables")
  )
}

app_dir <- Sys.getenv("BSTVC_APP_DIR", unset = "")
if (!nzchar(app_dir) || !dir.exists(app_dir)) {
  app_dir <- getwd()
}
app_dir <- normalizePath(app_dir, winslash = "/", mustWork = FALSE)

# Static files must live beside app.R so the packaged app works on other machines.
figure_candidates <- unique(c(
  file.path(app_dir, "www"),
  file.path(app_dir, "Figures"),
  file.path(dirname(app_dir), "Figures"),
  file.path(getwd(), "www"),
  file.path(getwd(), "Figures")
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
card_logo_src <- figure_src("bstvc_card_logo_202607_circle.png")
analysis_banner_src <- figure_src("bg-BSTVC-202607.png")
wechat_qr_src <- figure_src("HealthGeography_qr.png") %||% figure_src("HealthGeography_qr.jpg") %||% figure_src("HealthGeography_qr.jpeg")

####################################################################################################
# 三、数据读取与检查工具
#
# BSTVC 和 BSVC 对数据顺序非常敏感：Data Table中空间单元的排列必须与Map sf 对象中的几何
# 单元排列一致。这里先可靠读取数据、Map和可选Spatial weight matrix，再提供顺序检查函数。
####################################################################################################

# 读取建模数据。CSV 会自动尝试 UTF-8、GB18030/GBK 等常见中文编码；Excel 用 readxl。
# check.names = FALSE 是为了保留用户原始Field Name，避免界面中看到的列名和原文件不一致。
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

read_field_type_schema <- function(path) {
  first_line <- tryCatch(readLines(path, n = 1, warn = FALSE, encoding = "UTF-8"), error = function(e) "")
  first_line <- sub("^\ufeff", "", first_line)
  prefix <- "# BSTVC_FIELD_TYPES:"
  if (!length(first_line) || !startsWith(first_line, prefix)) {
    return(list(skip = 0, colClasses = NULL))
  }

  raw_schema <- sub(prefix, "", first_line, fixed = TRUE)
  pairs <- strsplit(raw_schema, ";", fixed = TRUE)[[1]]
  pairs <- pairs[nzchar(pairs)]
  col_classes <- character(0)
  for (pair in pairs) {
    parts <- strsplit(pair, "=", fixed = TRUE)[[1]]
    if (length(parts) < 2) next
    nm <- utils::URLdecode(parts[1])
    cls <- parts[2]
    if (cls %in% c("integer", "numeric", "character", "logical")) {
      col_classes[nm] <- cls
    }
  }
  list(skip = 1, colClasses = if (length(col_classes) > 0) col_classes else NULL)
}

read_csv_with_encoding <- function(path, encoding_label) {
  type_schema <- read_field_type_schema(path)
  args <- list(
    file = path,
    stringsAsFactors = FALSE,
    check.names = FALSE,
    comment.char = "",
    na.strings = c("", "NA", "NaN"),
    skip = type_schema$skip
  )
  if (!is.null(type_schema$colClasses)) {
    args$colClasses <- type_schema$colClasses
  }

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
    "invalid input|invalid multibyte|multi-byte|multi-byte|encoding",
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
        errors <- c(errors, sprintf("%s: no columns were read.", enc))
        next
      }
      if (nrow(dat) == 0) {
        errors <- c(errors, sprintf("%s: only the header was read and no data rows were found. This is usually caused by an encoding mismatch or an empty CSV file.", enc))
        next
      }
      if (any(is.na(names(dat))) || any(!nzchar(trimws(names(dat))))) {
        errors <- c(errors, sprintf("%s: the data contains empty column names. Please fill in column names in the source file first.", enc))
        next
      }
      attr(dat, "bstvc_csv_encoding") <- enc
      return(dat)
    }
  }

  stop(
    paste(
      c(
        "Failed to read the CSV file. Please make sure the file is not corrupted, not header-only, and preferably saved as UTF-8 or GBK/GB18030.",
        "The system has automatically tried UTF-8, UTF-8-BOM, GB18030, GBK, CP936, GB2312, and the default encoding.",
        "Encoding trial results:",
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

  stop("Modeling table data only supports csv, xlsx, or xls files.")
}

summarize_field_types <- function(df) {
  if (is.null(df) || !is.data.frame(df) || ncol(df) == 0) {
    return(data.frame('Field Name' = character(0), 'Field Type' = character(0), 'Missing Rate' = character(0), 'Example Value' = character(0), check.names = FALSE))
  }
  n <- max(nrow(df), 1)
  data.frame(
    'Field Name' = names(df),
    'Field Type' = vapply(df, function(x) paste(class(x), collapse = " / "), character(1)),
    'Missing Rate' = vapply(df, function(x) sprintf("%.2f%%", 100 * sum(is.na(x)) / n), character(1)),
    'Example Value' = vapply(df, function(x) {
      value <- x[which(!is.na(x))[1]]
      if (length(value) == 0 || is.na(value)) "" else substr(as.character(value), 1, 80)
    }, character(1)),
    check.names = FALSE,
    row.names = NULL
  )
}

file_input_with_clear <- function(input_id, label, ..., button_label = "Clear file") {
  tags$div(
    class = "file-input-with-clear",
    fileInput(input_id, label, ...),
    tags$button(
      type = "button",
      class = "btn btn-outline-success clear-upload-btn",
      `data-target` = input_id,
      title = button_label,
      icon("trash")
    )
  )
}

convert_field_vector <- function(x, target_type) {
  x_chr <- as.character(x)
  switch(
    target_type,
    numeric = suppressWarnings(as.numeric(x_chr)),
    integer = suppressWarnings(as.integer(as.numeric(x_chr))),
    character = x_chr,
    factor = as.factor(x_chr),
    logical = {
      y <- tolower(trimws(x_chr))
      out <- rep(NA, length(y))
      out[y %in% c("true", "t", "1", "yes", "y", "yes")] <- TRUE
      out[y %in% c("false", "f", "0", "no", "n", "no")] <- FALSE
      out
    },
    x
  )
}
# Shiny 上传 shapefile 时，浏览器会把每个文件放到临时路径中；如果只把 .shp 临时文件交给
# sf::st_read()，它在同一临时目录下找不到 .dbf/.shx 等配套文件，Map属性表就会丢失。
# 因此界面允许用户选择 shapefile 文件组，程序内部把这些文件按原文件名复制到同一临时目录，
# 再像本地 R 代码 st_read("NAT.shp") 一样读取主 .shp 文件。

# 识别上传表格更像空间截面宽表还是时空面板长表。用户的列名规则可能不同，因此这里只做
# 温和提示，真正转换仍以用户手动选择的字段和时间序列为准。
detect_panel_data_type <- function(dat) {
  if (is.null(dat) || !is.data.frame(dat) || ncol(dat) == 0) {
    return(list(type = "idle", class = "info", lines = "Please upload and load the source data to be converted first."))
  }

  cols <- names(dat)
  lower_cols <- tolower(cols)
  likely_time_cols <- cols[
    lower_cols %in% c("year", "time", "date", "month") |
      grepl("year|time|date|month", cols)
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
        sprintf("Data format check: spatiotemporal panel format (long table). One time field was detected: %s.", likely_time_cols[1]),
        "If the source data is already a spatiotemporal panel long table, you can go directly to Data Input. If conversion is still needed, set the parameters below."
      )
    ))
  }

  if (length(likely_wide_cols) >= 2) {
    return(list(
      type = "wide",
      class = "warning",
      lines = c(
        sprintf("Data format check: spatial cross-sectional format (wide table, one column per time point). Detected %d column names containing time information.", length(likely_wide_cols)),
        "Please select the time-series columns for each variable below, and make sure each variable has exactly the same number of columns as the time sequence length."
      )
    ))
  }

  list(
    type = "unknown",
    class = "info",
    lines = c(
      "Data format check: unable to determine the format.",
      "If each time point is stored as a separate column, continue setting conversion parameters as a spatial wide table. If a unique time field already exists, go directly to Data Input."
    )
  )
}

parse_numeric_time_sequence <- function(x) {
  x <- trimws(x %||% "")
  if (!nzchar(x)) {
    stop("Enter a time sequence, for example 2000:2021, 2000,2001,2002, or 1:12.", call. = FALSE)
  }

  compact <- gsub("\\s+", "", x)
  if (grepl("^-?[0-9]+(\\.[0-9]+)?[:-]-?[0-9]+(\\.[0-9]+)?$", compact)) {
    parts <- strsplit(compact, "[:-]")[[1]]
    start <- as.numeric(parts[1])
    end <- as.numeric(parts[2])
    if (is.na(start) || is.na(end)) {
      stop("The time sequence must be convertible to numeric values.", call. = FALSE)
    }
    step <- if (start <= end) 1 else -1
    return(seq(start, end, by = step))
  }

  pieces <- strsplit(x, "[,，;；\\s\\n\\r]+")[[1]]
  pieces <- pieces[nzchar(trimws(pieces))]
  values <- suppressWarnings(as.numeric(pieces))
  if (length(values) == 0 || any(is.na(values))) {
    stop("The time sequence must be numeric. Use commas, spaces, or a colon range.", call. = FALSE)
  }
  values
}


build_panel_conversion <- function(dat, id_cols, time_values, time_col, specs) {
  if (is.null(dat) || !is.data.frame(dat) || nrow(dat) == 0) {
    stop("Please load the source data to be converted first.", call. = FALSE)
  }
  if (length(id_cols) == 0) {
    stop("Please select at least one ID column.", call. = FALSE)
  }
  if (!all(id_cols %in% names(dat))) {
    stop("Some ID columns do not exist in the source data.", call. = FALSE)
  }
  if (!nzchar(trimws(time_col))) {
    stop("Enter the new time column name.", call. = FALSE)
  }
  if (time_col %in% id_cols) {
    stop("The new time column name cannot duplicate an ID column.", call. = FALSE)
  }
  if (length(time_values) == 0) {
    stop("The time sequence cannot be empty.", call. = FALSE)
  }

  value_names <- vapply(specs, function(x) trimws(x$value_name %||% ""), character(1))
  if (any(!nzchar(value_names))) {
    stop("Each variable must have a converted value column name.", call. = FALSE)
  }
  if (anyDuplicated(value_names) > 0) {
    stop("Converted value column names cannot be duplicated.", call. = FALSE)
  }
  if (any(value_names %in% c(id_cols, time_col))) {
    stop("Converted value column names cannot duplicate ID columns or the new time column.", call. = FALSE)
  }

  expected_len <- length(time_values)
  for (i in seq_along(specs)) {
    measure_cols <- specs[[i]]$measure_cols %||% character(0)
    value_name <- value_names[i]

    if (length(measure_cols) == 0) {
      stop(sprintf("Variable %d (%s) has no selected columns to convert.", i, value_name), call. = FALSE)
    }
    if (!all(measure_cols %in% names(dat))) {
      stop(sprintf("Variable %d (%s) contains columns that do not exist in the source data.", i, value_name), call. = FALSE)
    }
    if (length(measure_cols) != expected_len) {
      stop(
        sprintf(
          "Variable %d (%s) has %d columns, but the time sequence length is %d. Each variable must have exactly the same number of time points; please add missing time points, fill empty values with NA, and upload again.",
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
    stop("The modeling map data must include one .shp file.")
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

# 用户可以不上传Spatial weight matrix；这时 BSTVC/BSVC 会自动用 study_map 构造 QUEEN 邻接的 B 型矩阵。
# 如果用户上传 RDS，则允许 matrix、data.frame、spdep::listw 和 spdep::nb 四类常见对象。
safe_read_matrix_rds <- function(upload_info) {
  if (is.null(upload_info) || is.null(upload_info$name)) return(NULL)

  ext <- tolower(tools::file_ext(upload_info$name))
  if (!identical(ext, "rds")) {
    stop("The spatial weight matrix only supports .rds files.")
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
    stop("Unsupported RDS object. Please upload a matrix, data.frame, listw, or nb object.")
  }

  storage.mode(mat) <- "numeric"
  validate_spatial_matrix(mat, "Spatial weight matrix")
}

validate_spatial_matrix <- function(mat, label = "Spatial weight matrix") {
  if (is.null(mat)) stop(paste0(label, " is empty."), call. = FALSE)
  mat <- as.matrix(mat)
  if (length(dim(mat)) != 2 || nrow(mat) == 0 || ncol(mat) == 0) {
    stop(paste0(label, " must be a non-empty two-dimensional matrix."), call. = FALSE)
  }
  if (nrow(mat) != ncol(mat)) {
    stop(sprintf("%s must be a square matrix; current dimensions are %d x %d.", label, nrow(mat), ncol(mat)), call. = FALSE)
  }
  storage.mode(mat) <- "numeric"
  if (any(is.na(mat))) {
    stop(paste0(label, " contains NA values. Please rebuild it or check spatial-unit distances."), call. = FALSE)
  }
  if (any(!is.finite(mat))) {
    stop(paste0(label, " contains Inf or NaN values. Please check for duplicated coordinates, zero distances, or unreasonable distance-decay parameters."), call. = FALSE)
  }
  diag(mat) <- 0
  mat
}

# Map Attribute Fields可能因为 shapefile 驱动、编码或大小写差异表现为 FIPS/fips/Fips。
# 这个函数优先寻找完全同名字段；找不到时再做大小写不敏感匹配，并把匹配到的Map字段复制成
# 用户选择的数据Field Name，后续 BSTVC::data.check 和模型函数就能拿到一致字段。
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
    "The spatial field is not in the map attribute table. Data field: %s; map fields include: %s",
    space_col,
    paste(names(map_df), collapse = ", ")
  ))
}

get_space_field <- function(shp, space_col) {
  shp <- align_map_space_field(shp, space_col)
  sf::st_drop_geometry(shp)[[space_col]]
}

guess_utm_epsg <- function(shp) {
  shp_ll <- sf::st_transform(shp, 4326)
  bbox <- sf::st_bbox(shp_ll)
  lon <- mean(c(bbox[["xmin"]], bbox[["xmax"]]))
  lat <- mean(c(bbox[["ymin"]], bbox[["ymax"]]))
  zone <- floor((lon + 180) / 6) + 1
  if (is.na(zone) || zone < 1 || zone > 60) stop("Unable to infer the UTM zone from the map extent. Please use a manual EPSG code.", call. = FALSE)
  if (lat >= 0) 32600 + zone else 32700 + zone
}

prepare_weight_map <- function(shp, crs_mode, epsg = NULL) {
  if (identical(crs_mode, "keep")) return(shp)
  if (identical(crs_mode, "auto_utm")) {
    return(sf::st_transform(shp, guess_utm_epsg(shp)))
  }
  epsg <- suppressWarnings(as.integer(epsg))
  if (is.na(epsg) || epsg <= 0) stop("Enter a valid EPSG code.", call. = FALSE)
  sf::st_transform(shp, epsg)
}

nb_to_distance_listw <- function(nb, coords, centroid_sf, type, style = "raw", alpha = 1, dmax = NULL) {
  if ("nb2listwdist" %in% getNamespaceExports("spdep")) {
    return(spdep::nb2listwdist(nb, centroid_sf, type = type, style = style, alpha = alpha, dmax = dmax, zero.policy = TRUE))
  }
  if (!identical(type, "idw")) stop("The current spdep version does not support nb2listwdist(); only the built-in IDW fallback can be used.", call. = FALSE)
  glist <- lapply(seq_along(nb), function(i) {
    js <- nb[[i]]
    if (length(js) == 1 && js[1] == 0) return(numeric(0))
    dist <- sqrt((coords[i, 1] - coords[js, 1])^2 + (coords[i, 2] - coords[js, 2])^2)
    1 / (pmax(dist, .Machine$double.eps) ^ alpha)
  })
  spdep::nb2listw(nb, glist = glist, style = if (identical(style, "raw")) "B" else style, zero.policy = TRUE)
}

build_custom_weight_object <- function(shp, matrix_type, crs_mode, epsg, distance_lower, distance_upper, k, style, alpha, dmax) {
  if (is.null(shp) || nrow(shp) == 0) stop("Please load valid map data first.", call. = FALSE)
  style <- style %||% "B"
  distance_types <- c("distance_binary", "knn_binary", "distance_idw", "distance_exp", "distance_dpd", "knn_idw", "knn_exp")
  shp_work <- if (matrix_type %in% distance_types) {
    prepare_weight_map(shp, crs_mode, epsg)
  } else {
    shp
  }
  centroid <- suppressWarnings(sf::st_centroid(sf::st_geometry(shp_work), of_largest_polygon = TRUE))
  coords <- sf::st_coordinates(centroid)

  if (identical(matrix_type, "queen_binary")) {
    nb <- spdep::poly2nb(shp_work, queen = TRUE)
    listw <- spdep::nb2listw(nb, style = style, zero.policy = TRUE)
    label <- "QUEEN binary adjacency weight matrix"
  } else if (identical(matrix_type, "rook_binary")) {
    nb <- spdep::poly2nb(shp_work, queen = FALSE)
    listw <- spdep::nb2listw(nb, style = style, zero.policy = TRUE)
    label <- "ROOK binary adjacency weight matrix"
  } else if (identical(matrix_type, "distance_binary")) {
    if (is.na(distance_upper) || distance_upper <= 0) stop("The distance threshold must be greater than 0.", call. = FALSE)
    nb <- spdep::dnearneigh(coords, d1 = distance_lower %||% 0, d2 = distance_upper)
    listw <- spdep::nb2listw(nb, style = style, zero.policy = TRUE)
    label <- sprintf("Distance-threshold adjacency matrix (%.0f - %.0f m)", distance_lower %||% 0, distance_upper)
  } else if (identical(matrix_type, "knn_binary")) {
    k <- suppressWarnings(as.integer(k))
    if (is.na(k) || k < 1) stop("K for K-nearest neighbors must be an integer greater than or equal to 1.", call. = FALSE)
    nb <- spdep::knn2nb(spdep::knearneigh(coords, k = k), sym = TRUE)
    listw <- spdep::nb2listw(nb, style = style, zero.policy = TRUE)
    label <- sprintf("K-nearest-neighbor binary weight matrix (K = %d)", k)
  } else if (matrix_type %in% c("distance_idw", "distance_exp", "distance_dpd")) {
    if (is.na(distance_upper) || distance_upper <= 0) stop("A distance weight matrix requires a distance threshold greater than 0.", call. = FALSE)
    alpha <- suppressWarnings(as.numeric(alpha %||% 1))
    if (is.na(alpha) || alpha <= 0) stop("The distance weight parameter alpha must be greater than 0.", call. = FALSE)
    nb <- spdep::dnearneigh(coords, d1 = distance_lower %||% 0, d2 = distance_upper)
    dist_type <- switch(matrix_type, distance_idw = "idw", distance_exp = "exp", distance_dpd = "dpd")
    dmax_value <- if (identical(dist_type, "dpd")) {
      dmax <- suppressWarnings(as.numeric(dmax %||% distance_upper))
      if (is.na(dmax) || dmax <= 0) stop("Double-power distance weights (DPD) require dmax greater than 0.", call. = FALSE)
      dmax
    } else NULL
    listw <- nb_to_distance_listw(nb, coords, sf::st_sf(geometry = centroid), type = dist_type, style = "raw", alpha = alpha, dmax = dmax_value)
    label <- switch(
      matrix_type,
      distance_idw = sprintf("Distance-threshold inverse-distance weight matrix (IDW, alpha = %s)", alpha),
      distance_exp = sprintf("Distance-threshold exponential-decay weight matrix (EXP, alpha = %s)", alpha),
      distance_dpd = sprintf("Distance-threshold double-power distance weight matrix (DPD, alpha = %s, dmax = %s)", alpha, dmax_value)
    )
  } else if (matrix_type %in% c("knn_idw", "knn_exp")) {
    k <- suppressWarnings(as.integer(k))
    if (is.na(k) || k < 1) stop("K for K-nearest neighbors must be an integer greater than or equal to 1.", call. = FALSE)
    alpha <- suppressWarnings(as.numeric(alpha %||% 1))
    if (is.na(alpha) || alpha <= 0) stop("The distance weight parameter alpha must be greater than 0.", call. = FALSE)
    nb <- spdep::knn2nb(spdep::knearneigh(coords, k = k), sym = TRUE)
    dist_type <- if (identical(matrix_type, "knn_idw")) "idw" else "exp"
    listw <- nb_to_distance_listw(nb, coords, sf::st_sf(geometry = centroid), type = dist_type, style = "raw", alpha = alpha, dmax = NULL)
    label <- if (identical(matrix_type, "knn_idw")) {
      sprintf("K-nearest-neighbor inverse-distance weight matrix (K = %d, alpha = %s)", k, alpha)
    } else {
      sprintf("K-nearest-neighbor exponential-decay weight matrix (K = %d, alpha = %s)", k, alpha)
    }
  } else {
    stop("The selected matrix type is not currently supported.", call. = FALSE)
  }

  mat <- validate_spatial_matrix(spdep::listw2mat(listw), label)
  list(label = label, listw = listw, nb = nb, matrix = mat, map = shp, map_work = shp_work, coords = coords)
}

# Input Preview和矩阵预览都需要画Map。这里使用 leaflet::addTiles()，避免 addProviderTiles() 对
# leaflet.providers 的额外依赖；即使用户电脑没有安装 leaflet.providers，也能正常显示Map。
# 对普通Map预览，右下角图例说明绿色面代表Study-area spatial units；对矩阵Map预览，图例显示数值色带。
render_basic_map <- function(shp, values = NULL, label_prefix = NULL, range_legend = FALSE) {
  map_df <- sf::st_drop_geometry(shp)
  label_field <- names(map_df)[1]
  labels <- as.character(map_df[[label_field]])

  if (!is.null(values) && length(values) == nrow(shp)) {
    pal <- leaflet::colorNumeric(c("#f8efe5", "#d8bdca", "#927088"), domain = values, na.color = "#d9d9d9")
    labels <- paste0(labels, if (!is.null(label_prefix)) paste0(" | ", label_prefix, ": ") else " | value: ", round(values, 4))
    map <- leaflet(shp) %>%
        addTiles() %>%
        addPolygons(
          weight = 1,
          color = "#756b78",
          fillColor = pal(values),
          fillOpacity = 0.72,
          label = labels
        )
    if (isTRUE(range_legend)) {
      return(
        map %>%
          addControl(
            html = paste0(
              "<div style='background:rgba(255,255,255,0.86);padding:10px 12px;border-radius:8px;",
              "box-shadow:0 8px 24px rgba(72,58,69,0.13);color:#221b2b;font-size:12px;'>",
              "<div style='font-weight:700;margin-bottom:6px;'>", label_prefix %||% "value", "</div>",
              "<div style='text-align:center;color:#756b78;'>Low</div>",
              "<div style='width:18px;height:78px;margin:4px auto;border-radius:4px;",
              "background:linear-gradient(to bottom,#f8efe5,#d8bdca,#927088);'></div>",
              "<div style='text-align:center;color:#756b78;'>High</div>",
              "</div>"
            ),
            position = "bottomright"
          )
      )
    }
    return(
      map %>% addLegend("bottomright", pal = pal, values = values, title = label_prefix %||% "value")
    )
  }

  leaflet(shp) %>%
    addTiles() %>%
    addPolygons(weight = 1, color = "#756b78", fillColor = "#d8bdca", fillOpacity = 0.58, label = labels) %>%
    addLegend(
      position = "bottomright",
      colors = "#d8bdca",
      labels = "Study-area spatial units",
      title = "Map legend",
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
          title = tagList(icon("sliders-h"), paste0(" ", model_name, " Modeling Parameters")),
          if (!is_bstvc) {
            div(
              class = "model-note",
              icon("circle-info"),
              " BSVC uses spatial cross-sectional data: upload a table where each spatial unit appears only once. No time field is required."
            )
          },
          div(
            class = "model-note",
            icon("circle-info"),
            " Data source: if the previous data check shows the order is already aligned, choose Raw Data; if it reports automatic rearrangement, choose Checked Data."
          ),
          div(
            class = "model-note",
            icon("circle-info"),
            " Standardization: we strongly recommend standardizing X variables. This keeps all indicators on a comparable scale and can speed up model computation."
          ),
          fluidRow(
            column(3, single_select(ns("data_source"), "Data Source", choices = c("Please select" = "", "Raw Data" = "raw", "Checked Data" = "checked"), selected = "")),
            column(3, single_select(ns("response_var"), "Response Variable Y", choices = NULL)),
            column(3, multi_select(ns("covars"), "Explanatory Variables X", choices = NULL)),
            column(3, single_select(ns("response_type"), "Response Type", choices = c("Please select" = "", "Continuous" = "continuous", "Binary" = "binary", "Count" = "count"), selected = ""))
          ),
          fluidRow(
            if (is_bstvc) {
              column(3, single_select(ns("time_var"), "Time Field", choices = NULL))
            },
            column(3, single_select(ns("space_var"), "Space Field", choices = NULL)),
            column(3, radioButtons(ns("std_switch"), "Standardize X Variables", choices = c("No" = "no", "Yes" = "yes"), selected = "yes", inline = TRUE)),
            column(3, radioButtons(ns("spmat_mode"), "Spatial weight matrix", choices = c("Default QUEEN-B" = "default", "Custom RDS" = "custom"), selected = "default", inline = TRUE))
          ),
          fluidRow(
            column(
              3,
              numericInput(
                ns("threads"),
                "Threads",
                value = default_threads,
                min = 1,
                max = max_threads,
                step = 1
              )
            )
          ),
          actionButton(ns("run_model"), tagList(icon("play"), " Run Model"), class = "btn btn-success run-btn"),
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
          title = tagList(icon("table"), paste0(" ", model_name, " Result Tables")),
          fluidRow(
            column(3, selectInput(ns("result_part"), "Result Module", choices = character(0))),
            column(
              3,
              selectInput(
                ns("download_format"),
                "Download File Type",
                choices = c("CSV" = "csv", "XLSX" = "xlsx", "XLS" = "xls"),
                selected = "xlsx"
              )
            ),
            column(3, downloadButton(ns("download_table"), "Download Current Table", class = "btn-outline-success download-btn result-action-btn")),
            column(3, downloadButton(ns("download_all_tables"), "Download All Tables", class = "btn-success download-btn result-action-btn"))
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
# 模型运行前会检查数据、Map、变量选择和自定义矩阵是否完整；运行时使用 waiter 遮罩和
# withProgress 阶段条，让用户明确知道应用正在工作。
####################################################################################################
model_server <- function(id, model_name = c("BSTVC", "BSVC"), raw_data_r, checked_data_r, map_r, matrix_r) {
  model_name <- match.arg(model_name)

  moduleServer(id, function(input, output, session) {
    result_r <- reactiveVal(NULL)
    status_r <- reactiveVal("The model has not been run yet.")
    tables_r <- reactiveVal(NULL)
    last_base_cols <- reactiveVal(character(0))
    last_cov_signature <- reactiveVal("")

    model_fun <- if (identical(model_name, "BSTVC")) BSTVC::BSTVC else BSTVC::BSVC

    # 用户可选择Raw Data或Checked Data。这里不能在“Checked Data”为空时自动退回Raw Data：
    # 如果用户主动选择了Checked Data，却还没有Done data.check，就必须明确提示，否则会让用户
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

      updateSelectizeInput(session, "response_var", choices = c("Please select" = "", cols), selected = "", server = TRUE)
      updateSelectizeInput(session, "space_var", choices = c("Please select" = "", cols), selected = "", server = TRUE)

      if (identical(model_name, "BSTVC")) {
        updateSelectizeInput(session, "time_var", choices = c("Please select" = "", cols), selected = "", server = TRUE)
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
        status_r("Please select a data source.")
        return()
      }

      if (identical(input$data_source, "checked") && is.null(checked_data_r())) {
        status_r("Checked Data was selected, but no checked data is available. Run the corresponding data check on the Data Check page first, or switch to Raw Data.")
        return()
      }

      if (is.null(dat) || nrow(dat) == 0) {
        status_r("No modeling data is available. Upload data first, or choose the correct data source based on the check result.")
        return()
      }

      if (is.null(shp)) {
        status_r("No modeling map data is available. Upload a shapefile first.")
        return()
      }

      if (length(input$covars %||% character(0)) == 0) {
        status_r("Please select at least one explanatory variable.")
        return()
      }

      if (is.null(input$response_var) || identical(input$response_var, "")) {
        status_r("Please select response variable Y.")
        return()
      }

      if (is.null(input$space_var) || identical(input$space_var, "")) {
        status_r("Please select the Space field.")
        return()
      }

      if (identical(model_name, "BSTVC") && (is.null(input$time_var) || identical(input$time_var, ""))) {
        status_r("Please select the Time field.")
        return()
      }

      if (is.null(input$response_type) || identical(input$response_type, "")) {
        status_r("Please select a response type.")
        return()
      }

      if (identical(input$spmat_mode, "custom") && is.null(matrix_r())) {
        status_r("A custom spatial matrix was selected, but no valid RDS file has been uploaded.")
        return()
      }

      waiter_obj <- waiter::Waiter$new(
        html = tagList(
          waiter::spin_fading_circles(),
          tags$h4("The Bayesian STVC/SVC model is running. Please wait...")
        ),
        color = "rgba(52, 42, 57, 0.58)"
      )

      waiter_obj$show()
      on.exit(waiter_obj$hide(), add = TRUE)

      tryCatch({
        withProgress(message = paste(model_name, "Modeling Progress"), value = 0, {
          incProgress(0.10, detail = "Checking input data and fields")

          work_data <- dat
          shp <- align_map_space_field(shp, input$space_var)
          map_df <- sf::st_drop_geometry(shp)

          if (!input$response_var %in% names(work_data)) stop("The response variable is not in the data.")
          if (!input$space_var %in% names(work_data)) stop("The spatial field is not in the data.")
          if (!input$space_var %in% names(map_df)) stop("The spatial field is not in the map attribute table.")
          if (identical(model_name, "BSTVC") && !input$time_var %in% names(work_data)) stop("The time field is not in the data.")

          incProgress(0.15, detail = "Processing X-variable standardization")
          std_note <- "X-variable standardization was not performed."
          if (identical(input$std_switch, "yes")) {
            standardized_covars <- intersect(input$covars, names(work_data))
            numeric_vars <- standardized_covars[vapply(work_data[standardized_covars], is.numeric, logical(1))]
            skipped_vars <- setdiff(standardized_covars, numeric_vars)

            if (length(numeric_vars) > 0) {
              work_data[numeric_vars] <- lapply(work_data[numeric_vars], function(x) as.numeric(scale(x)))
              std_note <- paste0("Standardized:", paste(numeric_vars, collapse = ", "), "。")
            }

            if (length(skipped_vars) > 0) {
              std_note <- paste0(std_note, " Skipped non-numeric variables:", paste(skipped_vars, collapse = ", "), "。")
            }
          }

          incProgress(0.15, detail = "Constructing the model formula and parameters")
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
            stop("Threads must be an integer greater than or equal to 1.")
          }
          arg_list$threads <- threads_value

          if (identical(model_name, "BSTVC")) {
            arg_list$Time <- input$time_var
            arg_list$Space <- input$space_var
          } else {
            arg_list$Space <- input$space_var
          }

          status_r(paste("Model running. Formula: ", deparse(formula_obj)))
          incProgress(0.20, detail = "Calling the BSTVC package to run the model")

          model_res <- do.call(model_fun, arg_list)

          incProgress(0.25, detail = "Preparing model output tables")
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
            paste0(" Help-document result modules not returned: ", paste(missing_expected, collapse = ", "), "。")
          } else {
            ""
          }

          incProgress(0.15, detail = "Done")
          status_r(paste0(
            model_name, " model completed.", std_note,
            " Currently available: ", length(tables), " result tables.", missing_note
          ))
        })
      }, error = function(e) {
        status_r(paste("Model run failed: ", conditionMessage(e)))
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
# 侧边栏展示完整工作流入口，并通过分割线区分主流程、建模、辅助转换和About信息。
# Data Input、Data Check、建模和Data Conversion页面之间同时保留按钮式跳转，便于按流程操作。
# 首页只说明建模流程和表格输出，不再提及结果图形展示。侧边栏绿色渐变保持不变，其余组件采用统一
# 间距、圆角、图标和按钮样式，提升整体规整度。
####################################################################################################
ui <- bs4DashPage(
  title = "BSTVC",
  header = bs4DashNavbar(
    skin = "light",
    border = TRUE,
    status = "white",
    title = bs4DashBrand(
      title = tagList(
        tags$strong("Spatiotemporal Interpretability")
      ),
      color = "info"
    )
  ),
  sidebar = bs4DashSidebar(
    skin = "light",
    status = "info",
    elevation = 3,
    collapsed = FALSE,
    bs4SidebarMenu(
      id = "tabs",
      bs4SidebarMenuItem("Overview", tabName = "overview", icon = icon("house"), selected = TRUE),
      tags$li(class = "nav-group-divider"),
      bs4SidebarMenuItem("Data Input", tabName = "input", icon = icon("database")),
      bs4SidebarMenuItem("Data Check", tabName = "check", icon = icon("list-check")),
      tags$li(class = "nav-group-divider"),
      bs4SidebarMenuItem("BSTVC modeling", tabName = "bstvc", icon = icon("chart-line")),
      bs4SidebarMenuItem("BSVC modeling", tabName = "bsvc", icon = icon("chart-area")),
      tags$li(class = "nav-group-divider"),
      bs4SidebarMenuItem("Data Conversion", tabName = "spatial_convert", icon = icon("shuffle")),
      bs4SidebarMenuItem("Field Conversion", tabName = "field_convert", icon = icon("wand-magic-sparkles")),
      bs4SidebarMenuItem("Custom Spatial Matrix", tabName = "matrix_builder", icon = icon("project-diagram")),
      tags$li(class = "nav-group-divider"),
      bs4SidebarMenuItem("About", tabName = "about", icon = icon("circle-info"))
    )
  ),
  controlbar = bs4DashControlbar(collapsed = TRUE),
  footer = bs4DashFooter(left = "BSTVC", right = "© 2026 West China Health and Medical Geography Research Group"),
  body = bs4DashBody(
    waiter::use_waiter(),
    tags$head(
      tags$style(HTML(
        "
        :root {
          --bstvc-bg: #fffdf9;
          --bstvc-surface: rgba(255,255,255,0.88);
          --bstvc-surface-strong: rgba(255,255,255,0.96);
          --bstvc-border: rgba(72, 58, 69, 0.09);
          --bstvc-divider: rgba(72, 58, 69, 0.11);
          --bstvc-text: #221b2b;
          --bstvc-muted: #756b78;
          --bstvc-accent: #9b7a91;
          --bstvc-accent-2: #c8949f;
          --bstvc-accent-soft: #f4e8ef;
          --bstvc-cream: #fff1d8;
          --bstvc-lavender: #fae8f1;
          --bstvc-gold: #e2a655;
          --bstvc-red: #c8949f;
          --bstvc-cyan: #d6b673;
          --bstvc-shadow: 0 14px 32px rgba(72, 58, 69, 0.06);
        }
        body, .content-wrapper, .main-sidebar {
          font-family: 'Segoe UI', 'Microsoft YaHei UI', Arial, sans-serif;
          color: var(--bstvc-text);
        }
        body {
          background:
            radial-gradient(circle at 14% 10%, rgba(255, 241, 216, 0.26), transparent 38%),
            radial-gradient(circle at 84% 8%, rgba(244, 232, 239, 0.22), transparent 34%),
            radial-gradient(circle at 62% 42%, rgba(255, 229, 207, 0.14), transparent 38%),
            linear-gradient(135deg, #fffefb 0%, #fffdf9 54%, #fcfaf8 100%);
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
          box-shadow: 0 10px 26px rgba(72, 58, 69, 0.055);
          background: rgba(255, 255, 255, 0.86) !important;
          backdrop-filter: blur(18px);
          -webkit-backdrop-filter: blur(18px);
        }
        .brand-link, .main-sidebar {
          background: rgba(255, 255, 255, 0.82) !important;
          border-right: 1px solid var(--bstvc-divider);
          backdrop-filter: blur(22px);
          -webkit-backdrop-filter: blur(22px);
        }
        body .main-sidebar .brand-link,
        body .main-sidebar .brand-link:hover,
        body .main-sidebar .brand-link:focus,
        body .main-sidebar .brand-link:active,
        body .main-sidebar .brand-link[class*='bg-'],
        body .main-sidebar .brand-link[class*='bg-']:hover,
        body .main-sidebar .brand-link[class*='navbar-'],
        body .main-sidebar .brand-link[class*='navbar-']:hover {
          background: rgba(255, 255, 255, 0.82) !important;
          background-color: rgba(255, 255, 255, 0.82) !important;
          background-image: none !important;
          color: var(--bstvc-text) !important;
        }
        .brand-link {
          min-height: 70px;
          display: flex;
          align-items: center;
          justify-content: center;
          text-align: center;
          white-space: normal !important;
          padding-right: 48px;
        }
        .brand-link:hover,
        .brand-link:focus,
        .brand-link:active {
          background: rgba(255, 255, 255, 0.82) !important;
          color: var(--bstvc-text) !important;
        }
        .brand-link strong, .brand-text {
          color: var(--bstvc-text) !important;
          font-family: Georgia, 'Times New Roman', 'Noto Serif SC', 'Source Han Serif SC', serif;
          font-weight: 700;
          font-style: italic;
          letter-spacing: 0;
          white-space: normal !important;
          overflow: visible !important;
          text-overflow: clip !important;
          display: block;
          text-align: center;
          line-height: 1.05;
          font-size: 18px;
          max-width: 170px;
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
          color: #756b78 !important;
          margin: 4px 10px;
          transition: background-color .16s ease, color .16s ease, box-shadow .16s ease;
        }
        .nav-sidebar .nav-link:hover {
          background: rgba(255,255,255,0.66);
          color: #221b2b !important;
        }
        .nav-sidebar .nav-link.active {
          background: rgba(255, 255, 255, 0.82) !important;
          color: #221b2b !important;
          box-shadow: inset 3px 0 0 var(--bstvc-accent), 0 8px 22px rgba(155,122,145,0.11);
        }
        .nav-group-divider {
          height: 1px;
          margin: 12px 20px;
          background: linear-gradient(90deg, transparent, rgba(117,107,120,0.20), transparent);
          list-style: none;
        }
        .sidebar-mini.sidebar-collapse .nav-sidebar .nav-group-divider {
          margin: 12px 18px;
        }
        .card, .info-box, .small-box {
          border-radius: 8px;
          border: 1px solid var(--bstvc-border);
          box-shadow: var(--bstvc-shadow);
          overflow: visible;
          background: var(--bstvc-surface) !important;
          backdrop-filter: blur(20px);
          -webkit-backdrop-filter: blur(20px);
        }
        .card-body, .tab-content, .tab-pane {
          overflow: visible !important;
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
          border: 1px solid rgba(72,58,69,0.14) !important;
          border-radius: 8px !important;
          opacity: 1 !important;
          box-shadow: none !important;
        }
        .card-tools .btn-tool:hover {
          background: rgba(244,231,207,0.62) !important;
          color: #221b2b !important;
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
          border: 1px solid rgba(72,58,69,0.07) !important;
          background:
            linear-gradient(135deg, rgba(255,255,255,0.76), rgba(240,223,230,0.38)) !important;
          box-shadow: 0 14px 30px rgba(72,58,69,0.055) !important;
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
          color: #221b2b !important;
          font-size: 24px !important;
          font-weight: 700 !important;
          margin: 0 0 6px 0 !important;
        }
        .small-box p {
          color: #756b78 !important;
          font-size: 15px !important;
          font-weight: 600 !important;
          margin: 0 !important;
        }
        .small-box .icon {
          top: 16px !important;
          right: 18px !important;
          color: rgba(72,58,69,0.12) !important;
          z-index: 1;
        }
        .small-box .icon > i,
        .small-box .icon > svg {
          color: rgba(72,58,69,0.12) !important;
          font-size: 70px !important;
        }
        .small-box:hover .icon > i,
        .small-box:hover .icon > svg {
          color: rgba(155,122,145,0.16) !important;
          transform: none !important;
        }
        .btn-success, .btn-primary {
          background: rgba(255, 255, 255, 0.72) !important;
          border-color: rgba(155, 122, 145, 0.24) !important;
          color: #221b2b !important;
          box-shadow: 0 10px 22px rgba(72,58,69,0.055) !important;
        }
        .btn-success:hover, .btn-primary:hover {
          background: rgba(244, 231, 207, 0.88) !important;
          border-color: rgba(201, 178, 139, 0.56) !important;
          color: #221b2b !important;
        }
        .btn-outline-success, .btn-outline-primary {
          color: #4c3d51 !important;
          border-color: rgba(155,122,145,0.24) !important;
          background: rgba(255,255,255,0.46) !important;
        }
        .btn-outline-success:hover, .btn-outline-primary:hover {
          background-color: rgba(244, 231, 207, 0.78) !important;
          border-color: rgba(201, 178, 139, 0.52) !important;
          color: #221b2b !important;
        }
        .btn-success.run-btn, .btn-primary.run-btn {
          background: rgba(244, 231, 207, 0.84) !important;
          border-color: rgba(201, 178, 139, 0.46) !important;
        }
        .btn-success.run-btn:hover, .btn-primary.run-btn:hover {
          background: rgba(239, 219, 181, 0.90) !important;
          border-color: rgba(185, 151, 98, 0.52) !important;
        }
        .file-input-with-clear {
          position: relative;
        }
        .file-input-with-clear .form-group {
          margin-bottom: 8px;
        }
        .clear-upload-btn {
          width: 38px;
          min-width: 38px;
          height: 38px;
          margin: 0 0 0 8px;
          padding: 0;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          border-radius: 7px !important;
        }
        .clear-upload-btn i,
        .clear-upload-btn svg {
          margin: 0 !important;
          font-size: 15px;
        }
        .progress {
          height: 12px !important;
          border-radius: 999px !important;
          background: rgba(255,255,255,0.60) !important;
          border: 1px solid rgba(72,58,69,0.07);
          box-shadow: inset 0 1px 2px rgba(72,58,69,0.035);
          overflow: hidden;
        }
        .progress-bar {
          background: linear-gradient(90deg, #f6ead2, #d6b673) !important;
          color: #221b2b !important;
          font-size: 11px !important;
          font-weight: 700 !important;
          line-height: 12px !important;
          box-shadow: none !important;
        }
        .btn, .form-control, .selectize-input {
          border-radius: 7px !important;
        }
        .form-control, .selectize-input, .selectize-control.single .selectize-input {
          border-color: rgba(117, 107, 120, 0.22);
          background: rgba(255,255,255,0.78) !important;
          box-shadow: none;
        }
        .form-control:focus, .selectize-input.focus {
          border-color: var(--bstvc-accent);
          box-shadow: 0 0 0 3px rgba(155, 122, 145, 0.12);
        }
        .selectize-dropdown {
          z-index: 20000 !important;
          max-height: 260px !important;
          overflow-y: auto !important;
          border-color: rgba(117, 107, 120, 0.22) !important;
          box-shadow: 0 18px 34px rgba(72,58,69,0.13) !important;
        }
        .selectize-dropdown-content {
          max-height: 240px !important;
          overflow-y: auto !important;
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
          min-height: 540px;
        }
        .overview-user-card .card-body {
          padding-top: 62px;
        }
        .overview-user-header {
          min-height: 210px;
          background-size: cover;
          background-position: center 58%;
          background-repeat: no-repeat;
          background-color: #071020;
          border-bottom: 1px solid var(--bstvc-divider);
          position: relative;
        }
        .overview-user-header::after {
          content: '';
          position: absolute;
          inset: 0;
          background: linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.10));
        }
        .overview-user-card .widget-user-image {
          position: absolute;
          top: 155px;
          left: 50% !important;
          right: auto !important;
          width: 122px;
          height: 122px;
          margin-left: -61px !important;
          transform: none;
          z-index: 3;
          border-radius: 50%;
          background: rgba(255,255,255,0.96);
          padding: 3px;
          overflow: hidden;
          box-shadow: 0 14px 30px rgba(72,58,69,0.16);
        }
        .overview-user-card .widget-user-image > img {
          position: absolute;
          left: 50%;
          top: 50%;
          width: 116px;
          height: 116px;
          object-fit: contain;
          object-position: center center;
          background: transparent;
          border: 0;
          box-shadow: none;
          transform: translate(-50%, -50%);
          transform-origin: center center;
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
          text-align: justify;
          text-justify: inter-word;
          hyphens: auto;
        }
        .overview-link-list {
          margin: 10px auto 0 auto;
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
          line-height: 1.25;
          margin: 0 10px 2px 10px;
          padding: 0;
          text-decoration: underline;
          text-decoration-color: rgba(122,92,112,0.32);
          text-underline-offset: 4px;
        }
        .overview-link-list a:hover {
          color: #221b2b;
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
        .overview-core-card .info-block {
          margin-bottom: 16px;
          padding: 16px 18px;
        }
        .overview-core-card .info-block:last-child {
          margin-bottom: 0;
        }
        .overview-step .info-box {
          background: rgba(255,255,255,0.62) !important;
          border: 1px solid rgba(72,58,69,0.07) !important;
          box-shadow: 0 12px 26px rgba(72,58,69,0.05) !important;
        }
        .overview-step .info-box-icon {
          color: #221b2b !important;
          border-radius: 8px !important;
          margin: 10px 0 10px 10px;
          width: 78px !important;
        }
        .overview-step .info-box-icon i {
          color: #221b2b !important;
        }
        .overview-step-input .info-box-icon {
          background: linear-gradient(135deg, rgba(255,244,222,0.72), rgba(235,205,145,0.22)) !important;
        }
        .overview-step-check .info-box-icon {
          background: linear-gradient(135deg, rgba(249,232,239,0.66), rgba(211,174,195,0.22)) !important;
        }
        .overview-step-model .info-box-icon {
          background: linear-gradient(135deg, rgba(242,232,246,0.68), rgba(196,179,214,0.22)) !important;
        }
        .overview-step .info-box-content {
          padding: 12px 14px !important;
        }
        .overview-step .info-box-text {
          color: #756b78 !important;
          font-weight: 500 !important;
        }
        .overview-step .info-box-number {
          color: #221b2b !important;
          font-size: 17px !important;
          font-weight: 760 !important;
          margin-top: 5px;
        }
        .workflow-next-card .card-body {
          padding: 18px 20px;
        }
        .workflow-note-list {
          color: #504754;
          line-height: 1.8;
          margin: 0 0 14px 0;
          padding-left: 20px;
        }
        .workflow-action-row {
          display: flex;
          flex-wrap: wrap;
          gap: 12px;
        }
        .workflow-split-row {
          justify-content: space-between;
        }
        .workflow-split-row .btn {
          flex: 1 1 240px;
        }
        .workflow-action-row .btn {
          min-width: 180px;
        }
        .model-note, .matrix-note {
          color: #756b78;
          background: rgba(255,255,255,0.50);
          border: 1px solid rgba(72,58,69,0.07);
          border-radius: 8px;
          padding: 10px 12px;
          line-height: 1.6;
          margin: 8px 0 14px 0;
        }
        .model-note i, .matrix-note i {
          color: var(--bstvc-accent);
        }
        .slice-panel {
          border: 1px solid rgba(72,58,69,0.07);
          border-radius: 8px;
          padding: 12px 14px 2px 14px;
          margin: 8px 0 12px 0;
          background: rgba(255,255,255,0.50);
        }
        .convert-banner {
          background: rgba(255,255,255,0.62);
          border: 1px solid rgba(72,58,69,0.08);
          border-left: 4px solid rgba(155,122,145,0.42);
          border-radius: 8px;
          color: #243044;
          font-weight: 600;
          line-height: 1.8;
          margin-bottom: 14px;
          padding: 12px 16px;
          box-shadow: 0 10px 24px rgba(72,58,69,0.055);
          backdrop-filter: blur(18px);
          -webkit-backdrop-filter: blur(18px);
        }
        .convert-banner i {
          color: #9a6b2f;
          margin-right: 8px;
        }
        .convert-banner strong, .convert-detail summary, .download-tabs .nav-link {
          color: #7a5c70;
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
          color: #7a5c70;
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
          color: #221b2b;
          background-color: rgba(240, 223, 230, 0.78);
          border-color: rgba(155,122,145,0.16);
        }
        .card-header .nav-tabs,
        .nav-tabs {
          border-bottom: 1px solid rgba(72,58,69,0.09) !important;
        }
        .card-header .nav-tabs .nav-item,
        .nav-tabs .nav-item {
          margin-bottom: 0;
        }
        .card-header .nav-tabs .nav-link,
        .nav-tabs .nav-link {
          color: #756b78 !important;
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
          color: #221b2b !important;
          background: rgba(255,255,255,0.54) !important;
          border-color: rgba(155,122,145,0.14) !important;
        }
        .card-header .nav-tabs .nav-link.active,
        .card-header .nav-tabs .nav-item.show .nav-link,
        .nav-tabs .nav-link.active,
        .nav-tabs .nav-item.show .nav-link {
          color: #221b2b !important;
          background: rgba(244,231,207,0.74) !important;
          border-color: rgba(155,122,145,0.16) !important;
          box-shadow: 0 8px 16px rgba(72,58,69,0.05);
        }
        .card-header-tabs {
          margin-right: 44px;
        }
        .nav-pills .nav-link {
          color: #756b78 !important;
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
          color: #221b2b !important;
          background: rgba(255,255,255,0.54) !important;
          border-color: rgba(155,122,145,0.14) !important;
        }
        .nav-pills .nav-link.active,
        .nav-pills .show > .nav-link {
          color: #221b2b !important;
          background: rgba(244,231,207,0.78) !important;
          border-color: rgba(155,122,145,0.16) !important;
          box-shadow: 0 8px 16px rgba(72,58,69,0.05);
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
          border-color: rgba(72,58,69,0.09);
          box-shadow: none;
        }
        .alert-info {
          border-left: 4px solid rgba(155,122,145,0.36);
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
          color: #221b2b;
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
        .preview-subsection {
          background: rgba(255,255,255,0.42);
          border: 1px solid rgba(72,58,69,0.07);
          border-radius: 8px;
          padding: 10px 12px 12px 12px;
          margin-bottom: 14px;
        }
        .preview-subsection.preview-main {
          background: rgba(255,255,255,0.20);
          border-color: rgba(72,58,69,0.05);
        }
        .preview-section-title {
          color: #342a39;
          font-size: 14px;
          font-weight: 650;
          margin: 0 0 4px 0;
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .preview-section-note {
          color: #756b78;
          font-size: 12px;
          line-height: 1.55;
          margin: 0 0 8px 0;
        }
        .conversion-divider {
          border: 0;
          border-top: 1px solid rgba(72,58,69,0.09);
          margin: 18px 0 18px 0;
        }
        .result-action-btn {
          height: 38px;
          line-height: 24px;
          margin-top: 25px;
          border-radius: 6px;
        }
        .about-card .info-block {
          min-height: 116px;
        }
        .about-card .info-block p {
          margin-bottom: 8px;
        }
        .about-card a {
          color: #7a5c70;
          font-weight: 650;
          text-decoration: underline;
          text-decoration-color: rgba(122,92,112,0.32);
          text-underline-offset: 4px;
        }
        .about-muted {
          color: #6a5f6e;
          line-height: 1.75;
        }
        .wechat-qr-wrap {
          display: flex;
          align-items: center;
          gap: 18px;
          flex-wrap: wrap;
        }
        .wechat-qr-img {
          width: 132px;
          height: 132px;
          object-fit: contain;
          border-radius: 8px;
          background: rgba(255,255,255,0.72);
          border: 1px solid rgba(72,58,69,0.09);
          padding: 8px;
        }
        .wechat-qr-placeholder {
          width: 132px;
          height: 132px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 8px;
          color: #7a5c70;
          background: rgba(244,231,207,0.48);
          border: 1px dashed rgba(122,92,112,0.28);
          border-radius: 8px;
          text-align: center;
          font-size: 12px;
          line-height: 1.45;
          padding: 10px;
        }
        .qa-list {
          margin: 10px 0 0 0;
          padding-left: 18px;
          color: #504754;
          line-height: 1.75;
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
      ,
      tags$script(HTML(
        "
        function placeClearUploadButtons() {
          $('.file-input-with-clear').each(function() {
            var wrapper = $(this);
            var btn = wrapper.children('.clear-upload-btn');
            var inputGroup = wrapper.find('.input-group').first();
            if (btn.length && inputGroup.length && !$.contains(inputGroup[0], btn[0])) {
              inputGroup.append(btn);
            }
          });
        }
        $(document).on('shiny:connected', placeClearUploadButtons);
        $(document).ready(function() {
          placeClearUploadButtons();
          setTimeout(placeClearUploadButtons, 300);
        });
        $(document).on('change', '.file-input-with-clear input[type=\"file\"]', function() {
          var wrapper = $(this).closest('.file-input-with-clear');
          var target = $(this).attr('id');
          if (this.files && this.files.length > 0) {
            wrapper.find('.progress').show();
            $('#' + target + '_progress').show();
          }
        });
        $(document).on('click', '.clear-upload-btn', function(e) {
          e.preventDefault();
          var target = $(this).data('target');
          var input = $('#' + target);
          var wrapper = $(this).closest('.file-input-with-clear');
          if (input.length) {
            input.val('');
            input.trigger('change');
          }
          wrapper.find('input.form-control, .form-control').val('').attr('placeholder', 'No file selected');
          wrapper.find('.progress').hide();
          wrapper.find('.progress-bar').css('width', '0%').text('');
          $('#' + target + '_progress').hide().find('.progress-bar').css('width', '0%').text('');
          if (window.Shiny) {
            Shiny.setInputValue(target + '_clear', new Date().getTime(), {priority: 'event'});
          }
        });
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
                    tags$div(icon("cube", class = "fa-4x"))
                  }
                ),
                tags$div(
                  class = "overview-user-body",
                  tags$h2(class = "overview-user-title", "BSTVC Spatiotemporal Interpretability Analysis Tool"),
                  tags$p(
                    class = "overview-user-text",
                    "Within a unified Bayesian framework, BSTVC supports local spatiotemporal interpretability, global contribution assessment, key-driver identification, and dynamic spatiotemporal prediction for geographic research and decision-making."
                  ),
                  tags$div(
                    class = "overview-link-list",
                    tags$a(href = "https://github.com/bayesianstvc/BSTVC-R", target = "_blank", icon("github"), " GitHub: bayesianstvc/BSTVC-R"),
                    tags$a(href = "https://chaosong.blog/bayesian-stvc/", target = "_blank", icon("book-open"), " BSTVC Profile")
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
              class = "overview-core-card",
              title = tagList(icon("book-open"), "Core Features"),
              tags$div(
                class = "info-block",
                tags$strong(icon("clipboard-check"), "Data Preprocessing"),
                tags$p("Convert spatial cross-sectional data into a spatiotemporal panel format, check whether table records match map-unit order, and automatically align records for reliable modeling input.")
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("chart-line"), "Spatiotemporal Interpretability Analysis - BSTVC Module"),
                tags$p("For spatiotemporal panel data, detect spatiotemporal nonstationarity in relationships between explanatory variables and the response, and analyze dynamic heterogeneity mechanisms.")
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("chart-area"), "Spatial Interpretability Analysis - BSVC Module"),
                tags$p("For single-time-point spatial cross-sectional data, detect spatial heterogeneity in explanatory-variable effects and characterize local driving patterns.")
              )
            )
          )
        ),
        fluidRow(
          column(
            4,
            tags$div(
              class = "overview-step overview-step-input",
              bs4InfoBox("Step 1", "Upload Data and Map", icon = icon("file-import"), color = "info", width = 12)
            )
          ),
          column(
            4,
            tags$div(
              class = "overview-step overview-step-check",
              bs4InfoBox("Step 2", "Run Data Order Check", icon = icon("list-check"), color = "warning", width = 12)
            )
          ),
          column(
            4,
            tags$div(
              class = "overview-step overview-step-model",
              bs4InfoBox("Step 3", "Run Models and Export Result Tables", icon = icon("download"), color = "info", width = 12)
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
              class = "workflow-next-card",
              title = tagList(icon("route"), " Before Use"),
              tags$ul(
                class = "workflow-note-list",
                tags$li("If the original modeling table is a spatial wide table, first use Data Conversion to reshape it into a spatiotemporal panel format and download the converted data."),
                tags$li("If the original data is already a model-ready spatiotemporal panel, or the task is BSVC spatial cross-sectional modeling, go directly to Data Input.")
              ),
              tags$div(
                class = "workflow-action-row workflow-split-row",
                actionButton("go_convert_from_overview", tagList(icon("shuffle"), " Go to Data Conversion"), class = "btn btn-outline-success run-btn"),
                actionButton("go_input_from_overview", tagList(icon("database"), " Go to Data Input"), class = "btn btn-success run-btn")
              )
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
                tags$strong("Note:"),
                " If the data is spatial cross-sectional, use this tool to convert it into a spatiotemporal panel. If it has already been converted or is already in panel format, proceed to Data Input."
              ),
              tags$br(),
              icon("language"),
              tags$span(" When converting data, use non-Chinese column names where possible to avoid CSV encoding issues. Excel files are usually less affected.")
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
              title = tagList(icon("file-import"), " Source Data Loading"),
              file_input_with_clear("convert_file", "Spatial Cross-sectional Source Data (CSV / Excel)", accept = c(".csv", ".xlsx", ".xls")),
              actionButton("convert_load_file", tagList(icon("folder-open"), " Load Data for Conversion"), class = "btn btn-success run-btn"),
              tags$div(
                class = "model-note",
                icon("circle-info"),
                "Column naming is not strictly constrained. Make sure the time columns for each variable are ordered by time, or select them in the same order as the time sequence."
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
              title = tagList(icon("table"), " Source Data"),
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
              title = tagList(icon("clipboard-check"), " Check Result"),
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
              title = tagList(icon("sliders"), " Conversion Parameters"),
              multi_select("convert_id_cols", "ID Columns (constant across time; multiple allowed)"),
              textAreaInput("convert_time_values", "Time Sequence (numeric)", value = "2000:2021", rows = 2, placeholder = "Example: 2000:2021, 2000,2001,2002, or 1:12"),
              textInput("convert_time_col", "New Time Column Name", value = "Year"),
              numericInput("convert_var_count", "Number of Variables to Convert", value = 1, min = 1, max = 20, step = 1),
              uiOutput("convert_variable_specs"),
              actionButton("run_panel_convert", tagList(icon("play"), " Start Conversion"), class = "btn btn-success run-btn")
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("table-columns"), " Conversion Result"),
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
              downloadButton("download_converted_panel", "Download Converted Table", class = "btn btn-outline-success download-btn")
            )
          )
        ),
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner convert-banner-bottom",
              icon("download"),
              tags$span("After conversion, download the converted data and use it in the next step, Data Input."),
              tags$br(),
              icon("list-check"),
              tags$span("Format conversion does not guarantee that table order matches map-unit order; you still need to run Data Check."),
              tags$div(
                class = "workflow-action-row",
                actionButton("go_input_from_convert", tagList(icon("database"), " Go to Data Input"), class = "btn btn-success run-btn")
              )
            )
          )
        )
      ),
      bs4TabItem(
        tabName = "field_convert",
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner",
              icon("wand-magic-sparkles"),
              tags$span(
                tags$strong("Field Conversion:"),
                "Use this module to batch-adjust table field types, such as converting character numbers to numeric values or categorical fields to character/factor types. Download the converted table before going to Data Input."
              )
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
              title = tagList(icon("sliders"), " Field Conversion Parameters"),
              file_input_with_clear("field_file", "Data Table for Field Type Conversion (CSV / Excel)", accept = c(".csv", ".xlsx", ".xls")),
              actionButton("field_load_file", tagList(icon("folder-open"), " Load Field Conversion Data"), class = "btn btn-success run-btn"),
              uiOutput("field_status_ui"),
              tags$div(
                class = "model-note",
                icon("circle-info"),
                "After loading data, select fields and the target type. This tool only changes field types, not field names."
              ),
              multi_select("field_columns", "Fields to Convert"),
              selectInput(
                "field_target_type",
                "Target Field Type",
                choices = c("Numeric" = "numeric", "Integer" = "integer", "Character" = "character", "Factor" = "factor", "Logical" = "logical"),
                selected = "numeric"
              ),
              actionButton("run_field_convert", tagList(icon("play"), " Start Field Conversion"), class = "btn btn-success run-btn")
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("table-columns"), " Field Conversion Result"),
              tags$div(
                class = "preview-subsection",
                tags$div(class = "preview-section-title", icon("list"), "Original Field Types"),
                tags$p(class = "preview-section-note", "After loading data, review field names, field types, missing-value percentages, and example values."),
                DTOutput("field_summary_preview")
              ),
              tags$div(
                class = "preview-subsection preview-main",
                tags$div(class = "preview-section-title", icon("wand-magic-sparkles"), "Converted Field Types"),
                tags$p(class = "preview-section-note", "After conversion, this area shows the re-detected field types rather than the full table again."),
                DTOutput("field_converted_summary_preview")
              ),
              uiOutput("field_result_ui"),
              tags$div(
                class = "download-tabs",
                tabsetPanel(
                  id = "field_download_format",
                  type = "tabs",
                  tabPanel("CSV", value = "csv"),
                  tabPanel("Excel xlsx", value = "xlsx"),
                  tabPanel("Excel xls", value = "xls")
                )
              ),
              downloadButton("download_field_converted", "Download Field-Converted Table", class = "btn btn-outline-success download-btn")
            )
          )
        ),
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner convert-banner-bottom",
              icon("download"),
              tags$span("After field conversion, download the converted data and use it in Data Input."),
              tags$div(
                class = "workflow-action-row",
                actionButton("go_input_from_field_convert", tagList(icon("database"), " Go to Data Input"), class = "btn btn-success run-btn")
              )
            )
          )
        )
      ),
      bs4TabItem(
        tabName = "matrix_builder",
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner",
              icon("project-diagram"),
              tags$span(
                tags$strong("Custom Spatial Matrix:"),
                "Build a spatial weight matrix from uploaded map data and export it as an .rds file. The exported file can be uploaded as a custom spatial weight matrix on the Data Input page."
              ),
              tags$br(),
              icon("circle-info"),
              tags$span("Distance-threshold, inverse-distance, and K-nearest-neighbor matrices depend on coordinate distances. If the map uses longitude/latitude, choose automatic UTM projection; if it already uses local projected coordinates, keep the original CRS.")
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
              title = tagList(icon("folder-open"), " Map Data Import"),
              file_input_with_clear("custom_matrix_shp", "Map Data (shapefile)", accept = c(".shp", ".dbf", ".shx", ".prj", ".cpg"), multiple = TRUE),
              actionButton("custom_matrix_load_map", tagList(icon("map"), " Load Map"), class = "btn btn-success run-btn"),
              uiOutput("custom_matrix_load_ui")
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("map"), " Map Identification"),
              tags$div(
                class = "preview-subsection",
                tags$p(class = "preview-section-note", "After loading the map, this area shows detected spatial units, geometry type, coordinate system type, extent, and related information."),
                uiOutput("custom_matrix_map_info")
              )
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
              title = tagList(icon("sliders"), " Spatial Matrix Parameters"),
              selectInput(
                "custom_matrix_type",
                "Spatial Matrix Type",
                choices = c(
                  "Polygon Adjacency: QUEEN" = "queen_binary",
                  "Polygon Adjacency: ROOK" = "rook_binary",
                  "Distance Adjacency: Threshold Binary" = "distance_binary",
                  "Neighbor Relation: KNN Binary" = "knn_binary",
                  "Distance Decay: Inverse Distance IDW" = "distance_idw",
                  "Distance Decay: Exponential EXP" = "distance_exp",
                  "Distance Decay: Double Power DPD" = "distance_dpd",
                  "KNN Distance Decay: IDW" = "knn_idw",
                  "KNN Distance Decay: EXP" = "knn_exp"
                ),
                selected = "queen_binary"
              ),
              conditionalPanel(
                condition = "['distance_binary','knn_binary','distance_idw','distance_exp','distance_dpd','knn_idw','knn_exp'].indexOf(input.custom_matrix_type) >= 0",
                selectInput(
                  "custom_matrix_crs_mode",
                  "Coordinate Distance Handling",
                  choices = c(
                    "Keep Original CRS (for projected coordinates)" = "keep",
                    "Automatically Convert to UTM (for longitude/latitude)" = "auto_utm",
                    "Enter EPSG Manually and Transform" = "manual_epsg"
                  ),
                  selected = "auto_utm"
                ),
                conditionalPanel(
                  condition = "input.custom_matrix_crs_mode == 'manual_epsg'",
                  numericInput("custom_matrix_epsg", "Target Projection EPSG", value = 3857, min = 1, step = 1)
                )
              ),
              conditionalPanel(
                condition = "['queen_binary','rook_binary','distance_binary','knn_binary'].indexOf(input.custom_matrix_type) >= 0",
                selectInput(
                  "custom_matrix_style",
                  "Weight Standardization",
                  choices = c("B: Binary Weights" = "B", "W: Row-standardized" = "W", "C: Globally Standardized" = "C", "U: Uniform" = "U", "S: Variance-stabilizing" = "S", "minmax：Kelejian-Prucha" = "minmax"),
                  selected = "B"
                )
              ),
              conditionalPanel(
                condition = "['distance_binary','distance_idw','distance_exp','distance_dpd'].indexOf(input.custom_matrix_type) >= 0",
                numericInput("custom_matrix_d1", "Minimum Distance Threshold (m)", value = 0, min = 0, step = 1000),
                numericInput("custom_matrix_d2", "Maximum Distance Threshold (m)", value = 500000, min = 1, step = 1000)
              ),
              conditionalPanel(
                condition = "['knn_binary','knn_idw','knn_exp'].indexOf(input.custom_matrix_type) >= 0",
                numericInput("custom_matrix_k", "Number of K Neighbors", value = 4, min = 1, step = 1)
              ),
              conditionalPanel(
                condition = "['distance_idw','distance_exp','distance_dpd','knn_idw','knn_exp'].indexOf(input.custom_matrix_type) >= 0",
                numericInput("custom_matrix_alpha", "Distance Decay Parameter alpha", value = 1, min = 0.0001, step = 0.1),
                tags$div(class = "model-note", icon("circle-info"), "For IDW, alpha = 1 means inverse distance and alpha = 2 means squared inverse distance. EXP and DPD also use alpha to control decay strength.")
              ),
              conditionalPanel(
                condition = "input.custom_matrix_type == 'distance_dpd'",
                numericInput("custom_matrix_dmax", "DPD Maximum Distance dmax (m)", value = 500000, min = 1, step = 1000)
              ),
              actionButton("build_custom_matrix", tagList(icon("play"), " Build Spatial Matrix"), class = "btn btn-success run-btn"),
              uiOutput("custom_matrix_result_ui"),
              tags$div(
                class = "workflow-action-row",
                downloadButton("download_custom_matrix", "Download Custom Spatial Matrix .rds", class = "btn btn-outline-success download-btn")
              )
            )
          ),
          column(
            8,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("project-diagram"), " Spatial Relationship Visualization"),
              tags$div(
                class = "preview-subsection preview-main",
                tags$p(class = "preview-section-note", "After construction, adjacency matrices show links between spatial units; continuous weight matrices are colored by weight strength."),
                leafletOutput("custom_matrix_map_preview", height = 430)
              )
            )
          )
        ),
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner convert-banner-bottom",
              icon("download"),
              tags$span("After the matrix is built, download the .rds file and upload it as a custom spatial weight matrix on the Data Input page. Make sure the map-unit order matches the modeling map."),
              tags$div(
                class = "workflow-action-row",
                actionButton("go_input_from_matrix_builder", tagList(icon("database"), " Go to Data Input"), class = "btn btn-success run-btn")
              )
            )
          )
        )
      ),
      bs4TabItem(
        tabName = "input",
        fluidRow(
          column(
            4,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("upload"), " Data Input"),
              file_input_with_clear("data_file", "Modeling Table Data (CSV / Excel)", accept = c(".csv", ".xlsx", ".xls")),
              tags$div(
                class = "matrix-note",
                icon("circle-info"),
                " Upload the .shp file together with its companion files: .shp/.dbf/.shx/.prj/.cpg."
              ),
              file_input_with_clear("shp_file", "Modeling Map Data (shapefile)", accept = c(".shp", ".dbf", ".shx", ".prj", ".cpg"), multiple = TRUE),
              tags$div(
                class = "matrix-note",
                icon("circle-info"),
                " BSTVC/BSVC support custom spatial weight matrices, such as KNN, inverse distance, and fixed distance. If none is uploaded, the system automatically constructs a default QUEEN adjacency B-type spatial weight matrix."
              ),
              file_input_with_clear("sp_matrix_file", "Spatial Weight Matrix (optional custom .rds)", accept = c(".rds")),
              actionButton("load_resources", tagList(icon("folder-open"), " Load and Validate Input"), class = "btn btn-success run-btn"),
              br(),
              br(),
              uiOutput("load_feedback")
            )
          ),
          column(
            8,
            fluidRow(
              column(4, bs4ValueBox(value = textOutput("n_data_rows", inline = TRUE), subtitle = "Rows", icon = icon("table-list"), color = "info", width = 12)),
              column(4, bs4ValueBox(value = textOutput("n_data_cols", inline = TRUE), subtitle = "Columns", icon = icon("table-columns"), color = "info", width = 12)),
              column(4, bs4ValueBox(value = textOutput("n_map_units", inline = TRUE), subtitle = "Map Units", icon = icon("map"), color = "info", width = 12))
            ),
            bs4TabCard(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              title = tagList(icon("eye"), " Input Preview"),
              id = "input_preview_tabs",
              tabPanel(
                "Data Table",
                icon = icon("table"),
                tags$div(
                  class = "preview-subsection",
                  tags$div(class = "preview-section-title", icon("list"), "Data Table Field Properties"),
                  tags$p(class = "preview-section-note", "Review the field list first: field name, field type, missing-value percentage, and example value."),
                  DTOutput("data_fields_preview")
                ),
                tags$div(
                  class = "preview-subsection preview-main",
                  tags$div(class = "preview-section-title", icon("table"), "Raw Data Preview"),
                  tags$p(class = "preview-section-note", "Then review table contents to check the uploaded raw records."),
                  DTOutput("data_preview")
                )
              ),
              tabPanel(
                "Map",
                icon = icon("map-marked-alt"),
                tags$div(
                  class = "preview-subsection",
                  tags$div(class = "preview-section-title", icon("list"), "Map Attribute Fields"),
                  tags$p(class = "preview-section-note", "Review the shapefile attribute list first: field names, field types, missing-value percentages, and example values."),
                  DTOutput("map_fields_preview")
                ),
                tags$div(
                  class = "preview-subsection preview-main",
                  tags$div(class = "preview-section-title", icon("map"), "Spatial Boundary Map Preview"),
                  tags$p(class = "preview-section-note", "Then review the map geometry to check spatial-unit boundaries and map loading results."),
                  leafletOutput("map_preview")
                )
              ),
              tabPanel("Spatial Matrix", icon = icon("border-all"), uiOutput("matrix_preview_ui"))
            )
          )
        ),
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner convert-banner-bottom",
              icon("database"),
              tags$span("After the modeling table, map data, and optional spatial weight matrix are loaded successfully, proceed to Data Check."),
              tags$div(
                class = "workflow-action-row",
                actionButton("go_check_from_input", tagList(icon("list-check"), " Go to Data Check"), class = "btn btn-success run-btn")
              )
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
              title = tagList(icon("list-check"), " Data Order Check"),
              radioButtons("check_mode", "Check Mode", choices = c("BSTVC Spatiotemporal Panel" = "bstvc", "BSVC Spatial Cross-section" = "bsvc"), selected = "bstvc", inline = FALSE),
              conditionalPanel(
                condition = "input.check_mode == 'bstvc'",
                selectInput("check_time", "Time Field", choices = NULL)
              ),
              selectInput("check_space", "Space Field", choices = NULL),
              actionButton("run_check", tagList(icon("play"), " Run Check"), class = "btn btn-success run-btn"),
              br(),
              br(),
              tags$div(
                class = "model-note",
                icon("circle-info"),
                " If the check result shows the order is already aligned, choose Raw Data on the modeling page. If the system has automatically rearranged the data, choose Checked Data."
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
              title = tagList(icon("clipboard-check"), " Check Result Log"),
              uiOutput("check_result_ui"),
              uiOutput("check_download_ui")
            )
          )
        ),
        fluidRow(
          column(
            12,
            tags$div(
              class = "convert-banner convert-banner-bottom",
              icon("clipboard-check"),
              tags$span("After data check, confirm the data source choice on the modeling page based on the result above; then enter the corresponding modeling page."),
              tags$div(
                class = "workflow-action-row workflow-split-row",
                actionButton("go_bstvc", tagList(icon("chart-line"), " Go to BSTVC Modeling"), class = "btn btn-success run-btn"),
                actionButton("go_bsvc", tagList(icon("chart-area"), " Go to BSVC Modeling"), class = "btn btn-outline-success run-btn")
              )
            )
          )
        )
      ),
      bs4TabItem(tabName = "bstvc", model_ui("bstvc_mod", "BSTVC")),
      bs4TabItem(tabName = "bsvc", model_ui("bsvc_mod", "BSVC")),
      bs4TabItem(
        tabName = "about",
        fluidRow(
          column(
            12,
            bs4Card(
              width = 12,
              status = "info",
              solidHeader = TRUE,
              class = "about-card",
              title = tagList(icon("circle-info"), " About BSTVC"),
              tags$p(
                class = "about-muted",
                "The Bayesian Spatiotemporally Varying Coefficients (BSTVC) model is an interpretability method based on Bayesian hierarchical modeling. It integrates spatial location, temporal dimension, and influencing factors into a unified framework to support local effect estimation, global contribution quantification, and dynamic trend prediction. It is suitable for public health, environmental health, socioeconomic, and other geographic-data research contexts."
              ),
              fluidRow(
                column(
                  6,
                  tags$div(
                    class = "info-block",
                    tags$strong(icon("users"), " Developers"),
                    tags$p(tags$strong("Chao Song"), tags$a(href = "https://chaosong.blog/", target = "_blank", " [Homepage]"), ": project lead, statistical theory advisor, and copyright holder."),
                    tags$p(tags$strong("Xianteng Tang"), tags$a(href = "https://tangxxxxt.github.io/", target = "_blank", " [Homepage]"), ": system developer, responsible for software design and user support.")
                  )
                ),
                column(
                  6,
                  tags$div(
                    class = "info-block",
                    tags$strong(icon("envelope"), " Contact Us"),
                    tags$p(tags$a(href = "mailto:chaosong.gis@gmail.com", "chaosong.gis@gmail.com"), "(Chao Song, statistical theory)"),
                    tags$p(tags$a(href = "mailto:tangxt.me@gmail.com", "tangxt.me@gmail.com"), "(Xianteng Tang, software use)")
                  )
                )
              ),
              fluidRow(
                column(
                  6,
                  tags$div(
                    class = "info-block",
                    tags$strong(icon("link"), " Official Links"),
                    tags$p(tags$strong("BSTVC Software Page"), "：", tags$a(href = "https://github.com/bayesianstvc/BSTVC-R", target = "_blank", "https://github.com/bayesianstvc/BSTVC-R")),
                    tags$p(tags$strong("Model Theory Page"), "：", tags$a(href = "https://chaosong.blog/bayesian-stvc/", target = "_blank", "https://chaosong.blog/bayesian-stvc/"))
                  )
                ),
                column(
                  6,
                  tags$div(
                    class = "info-block",
                    tags$strong(icon("heart"), " Acknowledgements"),
                    tags$p("We thank ", tags$strong("Jin Xue"), "、", tags$strong("Zizhu Yang"), "、", tags$strong("Mengmeng Lei"), "、", tags$strong("Linbo Jiang"), " from Sichuan University and ", tags$strong("Zhangying Tang"), " from Southwest Petroleum University for technical and practical support for this desktop modeling software; we also thank ", tags$strong("Mingyu Xie"), " for support with the desktop software logo design.")
                  )
                )
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("quote-left"), " References"),
                tags$ul(
                  class = "qa-list",
                  tags$li(
                    tags$strong("[Bayesian STVC series models] "),
                    "Song, Chao, Yin, Hao, Shi, Xun, Xie, Mingyu, Yang, Shujuan, Zhou, Junmin, Wang, Xiuli, Tang, Zhangying, Yang, Yili, & Pan, Jay. (2022). Spatiotemporal disparities in regional public risk perception of COVID-19 using Bayesian Spatiotemporally Varying Coefficients (STVC) series models across Chinese cities. International Journal of Disaster Risk Reduction, 77, 103078."
                  ),
                  tags$li(
                    tags$strong("[STVPI] "),
                    "Wan, Qin, Tang, Zhangying, Pan, Jay, Xie, Mingyu, Wang, Shaobin, Yin, Hao, Li, Junmin, Liu, Xin, Yang, Yang, & Song, Chao. (2022). Spatiotemporal heterogeneity in associations of national population ageing with socioeconomic and environmental factors at the global scale. Journal of Cleaner Production, 373, 133781."
                  )
                )
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("qrcode"), " Medical Geographic Information and Spatial Health Statistics / HealthGeography"),
                tags$div(
                  class = "wechat-qr-wrap",
                  if (!is.null(wechat_qr_src)) {
                    tags$img(src = wechat_qr_src, class = "wechat-qr-img", alt = "HealthGeography QR code")
                  } else {
                    tags$div(
                      class = "wechat-qr-placeholder",
                      icon("qrcode"),
                      tags$span("Place the QR code image named HealthGeography_qr.png in the www folder.")
                    )
                  },
                  tags$p(class = "about-muted", "Follow the WeChat official account for medical geographic information, spatial health statistics, and BSTVC model updates.")
                )
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("question-circle"), " Q & A"),
                tags$ul(
                  class = "qa-list",
                  tags$li("Global configuration: the desktop upload component has a small default limit, while spatial panel data and shapefiles often exceed it. This desktop app therefore raises the upload limit to 1 GB."),
                  tags$li("Computational complexity: BSTVC/BSVC are Bayesian local modeling methods. Computation changes with the number of spatial units, time points, explanatory variables, response type, and thread settings. A computer with sufficient memory is recommended."),
                  tags$li("Spatial weight matrix: the KNN method may report errors or fail to produce output; this may relate to the implementation of the underlying Bayesian inference algorithm (INLA)."),
                  tags$li(
                    tags$span("Thread Monitor: "),
                    tags$a(
                      href = "https://github.com/bayesianstvc/inla-monitor",
                      target = "_blank",
                      "Real-time INLA process monitor for Windows"
                    ),
                    tags$span(" - find your optimal thread count and solve Bayesian latent Gaussian models faster.")
                  ),
                  tags$li("Residual term: Poisson and logistic (binary) types currently do not include residual terms, so STVPI cannot compute the overall model interpretability percentage. This feature is planned for a later version."),
                  tags$li("Nonstationarity: the current BSTVC model assumes spatiotemporal nonstationarity for all explanatory variables. Later versions will support more flexible assumptions for individual variables.")
                )
              ),
              tags$div(
                class = "info-block",
                tags$strong(icon("history"), " Version Updates and Technical Support"),
                tags$ul(
                  class = "qa-list",
                  tags$li("This program will be updated periodically with new features and improvements. User contributions are welcome, including bug reports, feature requests, and changes. If you encounter problems or need further help, please contact us.")
                )
              )
            )
          )
        )
      )
    )
  )
)

####################################################################################################
# 七、Server 后端
#
# Server 保存用户上传的数据、Map、Spatial Matrix和检查后的数据，并把这些响应式对象传给 BSTVC/BSVC
# 建模模块。这里的状态对象尽量保持简单：raw_data 是Raw Data，checked_bstvc 和 checked_bsvc
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
  field_raw_data <- reactiveVal(NULL)
  field_converted_data <- reactiveVal(NULL)
  custom_matrix_map <- reactiveVal(NULL)
  custom_matrix_result <- reactiveVal(NULL)
  custom_matrix_load_status <- reactiveVal(list(type = "info", lines = c("Please import the complete shapefile set, not only the .shp file.", "Please import companion files such as .dbf, .shx, .prj, and .cpg together.")))
  custom_matrix_build_status <- reactiveVal(list(type = "info", lines = "No spatial matrix has been built yet."))
  load_status <- reactiveVal(list(type = "info", lines = "Upload data and a map, then click Load and Validate Input."))
  check_state <- reactiveVal(list(code = "idle", message = "Waiting for check.", mode = "bstvc"))
  check_history <- reactiveVal(list())
  field_status <- reactiveVal(list(type = "info", lines = "Please load the data table for field type conversion."))
  field_result_status <- reactiveVal(list(type = "info", lines = "Field type conversion has not been performed."))

  observeEvent(input$convert_file_clear, {
    converter_raw_data(NULL)
    converted_panel_data(NULL)
    convert_status(list(type = "info", lines = "Upload spatial cross-sectional source data, then click Load Data for Conversion."))
    convert_result_status(list(type = "info", lines = "Conversion has not been performed."))
  }, ignoreInit = TRUE)

  observeEvent(input$field_file_clear, {
    field_raw_data(NULL)
    field_converted_data(NULL)
    updateSelectizeInput(session, "field_columns", choices = character(0), selected = character(0))
    field_status(list(type = "info", lines = "Please load the data table for field type conversion."))
    field_result_status(list(type = "info", lines = "Field type conversion has not been performed."))
  }, ignoreInit = TRUE)

  observeEvent(input$custom_matrix_shp_clear, {
    custom_matrix_map(NULL)
    custom_matrix_result(NULL)
    custom_matrix_load_status(list(type = "info", lines = c("Please import the complete shapefile set, not only the .shp file.", "Please import companion files such as .dbf, .shx, .prj, and .cpg together.")))
    custom_matrix_build_status(list(type = "info", lines = "No spatial matrix has been built yet."))
  }, ignoreInit = TRUE)

  observeEvent(input$data_file_clear, {
    raw_data(NULL)
    checked_bstvc(NULL)
    checked_bsvc(NULL)
    load_status(list(type = "info", lines = "Modeling table data has been cleared. Upload data again and click Load and Validate Input."))
  }, ignoreInit = TRUE)

  observeEvent(input$shp_file_clear, {
    study_map(NULL)
    checked_bstvc(NULL)
    checked_bsvc(NULL)
    load_status(list(type = "info", lines = "Modeling map data has been cleared. Upload the shapefile and companion files again, then click Load and Validate Input."))
  }, ignoreInit = TRUE)

  observeEvent(input$sp_matrix_file_clear, {
    spatial_matrix(NULL)
    load_status(list(type = "info", lines = "The custom spatial matrix has been cleared; modeling will use the package default matrix unless you upload an .rds file again."))
  }, ignoreInit = TRUE)

  output$n_data_rows <- renderText({ if (is.null(raw_data())) "-" else as.character(nrow(raw_data())) })
  output$n_data_cols <- renderText({ if (is.null(raw_data())) "-" else as.character(ncol(raw_data())) })
  output$n_map_units <- renderText({ if (is.null(study_map())) "-" else as.character(nrow(study_map())) })

  # 点击读取按钮后才更新响应式数据，避免用户选择文件过程中重复触发大量读取操作。

  convert_status <- reactiveVal(list(
    type = "info",
    lines = "Upload spatial cross-sectional source data, then click Load Data for Conversion."
  ))
  convert_result_status <- reactiveVal(list(
    type = "info",
    lines = "Conversion has not been performed."
  ))

  observeEvent(input$convert_load_file, {
    if (is.null(input$convert_file)) {
      convert_status(list(type = "warning", lines = "Please select a CSV, xlsx, or xls source data file first."))
      return()
    }

    tryCatch({
      dat <- safe_read_data(input$convert_file)
      converter_raw_data(dat)
      converted_panel_data(NULL)
      convert_result_status(list(type = "info", lines = "Conversion has not been performed."))
      detected <- detect_panel_data_type(dat)
      convert_status(list(
        type = detected$class,
        lines = c(
          sprintf("Source data loaded successfully: %d rows, %d columns.", nrow(dat), ncol(dat)),
          detected$lines
        )
      ))
    }, error = function(e) {
      converter_raw_data(NULL)
      converted_panel_data(NULL)
      convert_result_status(list(type = "info", lines = "Conversion has not been performed."))
      convert_status(list(type = "danger", lines = paste("Failed to load source data: ", conditionMessage(e))))
    })
  }, ignoreInit = TRUE)

  observe({
    dat <- converter_raw_data()
    cols <- if (is.null(dat)) character(0) else names(dat)
    updateSelectizeInput(session, "convert_id_cols", choices = cols, selected = input$convert_id_cols %||% character(0))

    n_vars <- input$convert_var_count %||% 1
    range_choices <- c("Please select" = "", cols)
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
          showNotification("Please load source data before selecting columns by range.", type = "warning")
          return()
        }
        if (!nzchar(start_col) || !nzchar(end_col)) {
          showNotification("Please select a start column and an end column.", type = "warning")
          return()
        }
        if (!start_col %in% cols || !end_col %in% cols) {
          showNotification("The start or end column is not in the source data.", type = "error")
          return()
        }

        start_pos <- match(start_col, cols)
        end_pos <- match(end_col, cols)
        selected_cols <- cols[seq(min(start_pos, end_pos), max(start_pos, end_pos))]

        updateSelectizeInput(session, paste0("convert_measure_", i), selected = selected_cols)
        showNotification(
          paste0("Selected ", length(selected_cols), " columns in the source-data column order."),
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
            tags$summary(icon("chevron-right"), paste0("Variable ", i)),
            tags$div(
              class = "convert-detail-body",
              textInput(paste0("convert_value_name_", i), "Converted Value Column Name", value = if (i == 1) "Value" else paste0("Value", i)),
              multi_select(paste0("convert_measure_", i), "Time-series Columns for This Variable", choices = cols),
              tags$div(
                class = "range-select-row",
                fluidRow(
                  column(6, selectInput(paste0("convert_measure_start_", i), "Start Column", choices = c("Please select" = "", cols))),
                  column(6, selectInput(paste0("convert_measure_end_", i), "End Column", choices = c("Please select" = "", cols)))
                ),
                actionButton(paste0("convert_apply_measure_range_", i), tagList(icon("wand-magic-sparkles"), " Select by Range"), class = "btn btn-outline-success run-btn"),
                tags$div(class = "range-select-help", "After selecting a start and end column, the system selects all columns in that interval according to their source-data order.")
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
        tags$h4(icon("clipboard-check"), " Data Type Check"),
        lapply(st$lines, tags$p)
      ))
    }

    detected <- detect_panel_data_type(dat)
    div(
      class = paste("alert convert-type-alert", paste0("alert-", detected$class)),
      tags$h4(icon("clipboard-check"), " Data Type Check"),
      lapply(detected$lines, tags$p)
    )
  })

  observeEvent(input$run_panel_convert, {
    dat <- converter_raw_data()

    tryCatch({
      withProgress(message = "Data Conversion Progress", value = 0, {
        incProgress(0.15, detail = "Checking conversion parameters")
        time_values <- parse_numeric_time_sequence(input$convert_time_values)
        n_vars <- input$convert_var_count %||% 1
        specs <- lapply(seq_len(n_vars), function(i) {
          list(
            value_name = input[[paste0("convert_value_name_", i)]],
            measure_cols = input[[paste0("convert_measure_", i)]] %||% character(0)
          )
        })

        incProgress(0.35, detail = "Checking each variable's time-series length")
        converted <- build_panel_conversion(
          dat = dat,
          id_cols = input$convert_id_cols %||% character(0),
          time_values = time_values,
          time_col = input$convert_time_col,
          specs = specs
        )

        incProgress(0.30, detail = "Generating the spatiotemporal panel table")
        converted_panel_data(converted)
        incProgress(0.20, detail = "Done")
      })

      convert_result_status(list(
        type = "success",
        lines = c(
          sprintf("Conversion complete: generated a spatiotemporal panel with %d rows and %d columns.", nrow(converted_panel_data()), ncol(converted_panel_data())),
          "Download the converted data and use it in the next step, Data Input.",
          "Note: format conversion does not guarantee that table order matches map-unit order; run Data Check next."
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

  observeEvent(input$field_load_file, {
    if (is.null(input$field_file)) {
      field_status(list(type = "warning", lines = "Please select a CSV, xlsx, or xls data file first."))
      return()
    }

    tryCatch({
      dat <- safe_read_data(input$field_file)
      field_raw_data(dat)
      field_converted_data(NULL)
      updateSelectizeInput(session, "field_columns", choices = names(dat), selected = character(0))
      field_status(list(type = "success", lines = sprintf("Field conversion data loaded successfully: %d rows, %d columns.", nrow(dat), ncol(dat))))
      field_result_status(list(type = "info", lines = "Field type conversion has not been performed."))
    }, error = function(e) {
      field_raw_data(NULL)
      field_converted_data(NULL)
      updateSelectizeInput(session, "field_columns", choices = character(0), selected = character(0))
      field_status(list(type = "danger", lines = paste("Failed to load field conversion data: ", conditionMessage(e))))
      field_result_status(list(type = "info", lines = "Field type conversion has not been performed."))
    })
  }, ignoreInit = TRUE)

  output$field_status_ui <- renderUI({
    st <- field_status()
    div(class = paste("alert", paste0("alert-", st$type)), lapply(st$lines, tags$p))
  })

  output$field_summary_preview <- renderDT({
    dat <- field_raw_data()
    req(dat)
    DT::datatable(summarize_field_types(dat), rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })

  observeEvent(input$run_field_convert, {
    dat <- field_raw_data()
    cols <- input$field_columns %||% character(0)
    target_type <- input$field_target_type %||% "numeric"

    if (is.null(dat)) {
      field_result_status(list(type = "warning", lines = "Please load the data table for field type conversion first."))
      return()
    }
    if (length(cols) == 0) {
      field_result_status(list(type = "warning", lines = "Please select at least one field to convert."))
      return()
    }
    if (!all(cols %in% names(dat))) {
      field_result_status(list(type = "danger", lines = "The selected field is not in the data table. Reload the data or select fields again."))
      return()
    }

    converted <- dat
    for (nm in cols) {
      converted[[nm]] <- convert_field_vector(converted[[nm]], target_type)
    }
    field_converted_data(converted)
    field_result_status(list(
      type = "success",
      lines = c(
        sprintf("Field type conversion complete: %d fields converted to %s.", length(cols), target_type),
        "Review the re-detected field types, then download the field-converted table if correct."
      )
    ))
  }, ignoreInit = TRUE)

  output$field_converted_summary_preview <- renderDT({
    dat <- field_converted_data()
    req(dat)
    DT::datatable(summarize_field_types(dat), rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })

  output$field_result_ui <- renderUI({
    st <- field_result_status()
    div(class = paste("alert", paste0("alert-", st$type)), lapply(st$lines, tags$p))
  })

  output$download_field_converted <- downloadHandler(
    filename = function() {
      fmt <- input$field_download_format %||% "csv"
      paste0("BSTVC_field_type_converted_", format(Sys.Date(), "%Y%m%d"), ".", fmt)
    },
    content = function(file) {
      dat <- field_converted_data()
      req(dat)
      write_field_converted_table(dat, file, input$field_download_format %||% "csv")
    }
  )

  output$custom_matrix_load_ui <- renderUI({
    st <- custom_matrix_load_status()
    div(class = paste("alert", paste0("alert-", st$type)), lapply(st$lines, tags$p))
  })

  observeEvent(input$custom_matrix_load_map, {
    if (is.null(input$custom_matrix_shp)) {
      custom_matrix_load_status(list(type = "warning", lines = c("Please select the shapefile and its companion files first.", "Do not import only the .shp file. Import .dbf, .shx, .prj, .cpg, and other companion files together.")))
      return()
    }
    tryCatch({
      shp <- safe_read_shp_upload(input$custom_matrix_shp)
      custom_matrix_map(shp)
      custom_matrix_result(NULL)
      custom_matrix_build_status(list(type = "info", lines = "Map loaded. Set the matrix type and parameters, then click Build Spatial Matrix."))
      custom_matrix_load_status(list(
        type = "info",
        lines = c(
          "Map loaded. See Map Identification on the right for detected information.",
          "Note: a shapefile must be imported together with .dbf, .shx, .prj, .cpg, and other companion files. Importing only .shp is not recommended."
        )
      ))
    }, error = function(e) {
      custom_matrix_map(NULL)
      custom_matrix_result(NULL)
      custom_matrix_load_status(list(type = "danger", lines = paste("Failed to load map: ", conditionMessage(e))))
      custom_matrix_build_status(list(type = "info", lines = "No spatial matrix has been built yet."))
    })
  }, ignoreInit = TRUE)

  output$custom_matrix_map_info <- renderUI({
    shp <- custom_matrix_map()
    if (is.null(shp)) return(tags$p(class = "about-muted", "No map has been loaded yet."))
    map_df <- sf::st_drop_geometry(shp)
    crs <- sf::st_crs(shp)
    crs_text <- crs$input %||% "Unrecognized coordinate system"
    crs_type <- if (is.na(crs)) {
      "Unrecognized"
    } else if (isTRUE(sf::st_is_longlat(shp))) {
      "Geographic CRS (longitude/latitude)"
    } else {
      "Projected or planar CRS"
    }
    geom_types <- paste(unique(as.character(sf::st_geometry_type(shp))), collapse = "、")
    bbox <- sf::st_bbox(shp)
    empty_n <- sum(sf::st_is_empty(shp))
    invalid_n <- sum(!sf::st_is_valid(shp, NA_on_exception = FALSE))
    area_text <- if (!isTRUE(sf::st_is_longlat(shp))) {
      area <- suppressWarnings(as.numeric(sum(sf::st_area(shp), na.rm = TRUE)))
      if (is.finite(area)) sprintf("%.2f", area) else "Unable to calculate"
    } else {
      "Not calculated for longitude/latitude coordinates"
    }
    tagList(
      tags$p(tags$strong("Spatial Units: "), nrow(shp)),
      tags$p(tags$strong("Attribute Fields: "), ncol(map_df)),
      tags$p(tags$strong("Geometry Type: "), geom_types),
      tags$p(tags$strong("CRS: "), crs_text),
      tags$p(tags$strong("CRS Type: "), crs_type),
      tags$p(tags$strong("Map Extent: "), sprintf("xmin %.4f, ymin %.4f, xmax %.4f, ymax %.4f", bbox[["xmin"]], bbox[["ymin"]], bbox[["xmax"]], bbox[["ymax"]])),
      tags$p(tags$strong("Empty Geometries: "), empty_n),
      tags$p(tags$strong("Invalid Geometries: "), invalid_n),
      tags$p(tags$strong("Total Area: "), area_text)
    )
  })

  observeEvent(input$build_custom_matrix, {
    shp <- custom_matrix_map()
    if (is.null(shp)) {
      custom_matrix_build_status(list(type = "warning", lines = "Please load the map data used to build the spatial matrix first."))
      return()
    }
    tryCatch({
      result <- build_custom_weight_object(
        shp = shp,
        matrix_type = input$custom_matrix_type,
        crs_mode = input$custom_matrix_crs_mode,
        epsg = input$custom_matrix_epsg,
        distance_lower = input$custom_matrix_d1 %||% 0,
        distance_upper = input$custom_matrix_d2,
        k = input$custom_matrix_k,
        style = input$custom_matrix_style %||% "B",
        alpha = input$custom_matrix_alpha %||% 1,
        dmax = input$custom_matrix_dmax
      )
      custom_matrix_result(result)
      nonzero <- sum(result$matrix != 0, na.rm = TRUE)
      custom_matrix_build_status(list(
        type = "success",
        lines = c(
          paste("Spatial matrix built successfully: ", result$label),
          sprintf("Matrix dimensions: %d x %d; non-zero weights: %d.", nrow(result$matrix), ncol(result$matrix), nonzero)
        )
      ))
    }, error = function(e) {
      custom_matrix_result(NULL)
      custom_matrix_build_status(list(type = "danger", lines = paste("Failed to build spatial matrix: ", conditionMessage(e))))
    })
  }, ignoreInit = TRUE)

  output$custom_matrix_result_ui <- renderUI({
    st <- custom_matrix_build_status()
    div(class = paste("alert", paste0("alert-", st$type)), lapply(st$lines, tags$p))
  })

  output$custom_matrix_map_preview <- renderLeaflet({
    result <- custom_matrix_result()
    shp <- custom_matrix_map()
    if (is.null(shp)) {
      return(leaflet() %>% addTiles() %>% addControl(html = "Please load the map first.", position = "topright"))
    }
    if (is.null(result)) {
      return(render_basic_map(shp))
    }

    mat <- as.matrix(result$matrix)
    shp <- result$map
    if (nrow(mat) != nrow(shp)) return(render_basic_map(shp))
    mat[is.na(mat)] <- 0
    if (nrow(mat) == ncol(mat)) diag(mat) <- 0
    nonzero_vals <- unique(as.numeric(mat[mat != 0]))
    is_binary <- length(nonzero_vals) > 0 && all(nonzero_vals %in% c(1))

    if (is_binary) {
      coords <- suppressWarnings(sf::st_coordinates(sf::st_centroid(sf::st_geometry(shp), of_largest_polygon = TRUE)))
      links <- which(mat != 0, arr.ind = TRUE)
      links <- links[links[, 1] != links[, 2], , drop = FALSE]
      link_count <- nrow(links)
      if (link_count == 0) return(render_basic_map(shp))
      if (link_count > 5000) links <- links[seq_len(5000), , drop = FALSE]
      lng <- as.vector(rbind(coords[links[, 1], 1], coords[links[, 2], 1], NA))
      lat <- as.vector(rbind(coords[links[, 1], 2], coords[links[, 2], 2], NA))
      return(
        leaflet(shp) %>%
          addTiles() %>%
          addPolygons(weight = 1, color = "#756b78", fillColor = "#f8efe5", fillOpacity = 0.62, label = as.character(sf::st_drop_geometry(shp)[[1]])) %>%
          addPolylines(lng = lng, lat = lat, color = "#8b6f82", weight = 1, opacity = 0.45) %>%
          addControl(
            html = if (link_count > 5000) "Showing a sample of the first 5000 spatial links." else result$label,
            position = "bottomright"
          )
      )
    }

    weight_strength <- rowSums(abs(mat), na.rm = TRUE)
    render_basic_map(shp, values = weight_strength, label_prefix = "Weight Strength", range_legend = TRUE)
  })

  output$download_custom_matrix <- downloadHandler(
    filename = function() {
      paste0("BSTVC_custom_spatial_weight_", input$custom_matrix_type %||% "matrix", "_", format(Sys.Date(), "%Y%m%d"), ".rds")
    },
    content = function(file) {
      result <- custom_matrix_result()
      req(result)
      saveRDS(validate_spatial_matrix(result$matrix, result$label), file = file)
    }
  )

  observeEvent(input$load_resources, {
    showNotification("Reading and validating input files. Please wait...", type = "message", duration = 3)
    load_status(list(type = "info", lines = "Reading and validating input files. Please wait..."))

    msgs <- character(0)
    has_error <- FALSE

    if (!is.null(input$data_file)) {
      tryCatch({
        dat <- safe_read_data(input$data_file)
        raw_data(dat)
        checked_bstvc(NULL)
        checked_bsvc(NULL)
        msgs <- c(msgs, sprintf("[OK] Data loaded successfully."))
      }, error = function(e) {
        has_error <<- TRUE
        msgs <<- c(msgs, paste("[FAIL] Failed to load data: ", conditionMessage(e)))
      })
    } else {
      msgs <- c(msgs, "[INFO] Modeling data was not updated.")
    }

    if (!is.null(input$shp_file)) {
      tryCatch({
        shp <- safe_read_shp_upload(input$shp_file)
        study_map(shp)
        checked_bstvc(NULL)
        checked_bsvc(NULL)
        msgs <- c(msgs, sprintf("[OK] Map loaded successfully."))
      }, error = function(e) {
        has_error <<- TRUE
        msgs <<- c(msgs, paste("[FAIL] Failed to load map: ", conditionMessage(e)))
      })
    } else {
      msgs <- c(msgs, "[INFO] Map data was not updated.")
    }

    if (!is.null(input$sp_matrix_file)) {
      tryCatch({
        mat <- safe_read_matrix_rds(input$sp_matrix_file)
        spatial_matrix(mat)
        msgs <- c(msgs, sprintf("[OK] Spatial matrix loaded successfully: %d x %d.", nrow(mat), ncol(mat)))
      }, error = function(e) {
        has_error <<- TRUE
        spatial_matrix(NULL)
        msgs <<- c(msgs, paste("[FAIL] Failed to load spatial matrix: ", conditionMessage(e)))
      })
    } else {
      spatial_matrix(NULL)
      msgs <- c(msgs, "[INFO] No custom matrix uploaded; modeling will use the package default matrix.")
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

  output$data_fields_preview <- renderDT({
    dat <- raw_data()
    req(dat)
    DT::datatable(summarize_field_types(dat), rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })

  output$map_preview <- renderLeaflet({
    shp <- study_map()
    req(shp)
    render_basic_map(shp)
  })

  output$map_fields_preview <- renderDT({
    shp <- study_map()
    req(shp)
    map_df <- sf::st_drop_geometry(shp)
    DT::datatable(summarize_field_types(map_df), rownames = FALSE, options = list(pageLength = 8, scrollX = TRUE))
  })

  output$matrix_preview_ui <- renderUI({
    mat <- spatial_matrix()
    if (is.null(mat)) {
      return(tags$div(class = "alert alert-info", "No custom spatial matrix was uploaded; modeling will use the BSTVC/BSVC default matrix. After uploading a custom matrix, this area will show the spatial weight structure map."))
    }
    tags$div(
      class = "preview-subsection preview-main",
      tags$div(class = "preview-section-title", icon("project-diagram"), "Spatial Weight Structure Preview"),
      tags$p(class = "preview-section-note", "A 0/1 adjacency matrix will show links between spatial units; a continuous weight matrix will be colored by each unit's weight strength."),
      leafletOutput("matrix_map_preview")
    )
  })

  output$matrix_map_preview <- renderLeaflet({
    mat <- spatial_matrix()
    shp <- study_map()
    req(mat, shp)

    if (nrow(mat) != nrow(shp) || ncol(mat) != nrow(shp)) {
      return(
        leaflet() %>%
          addTiles() %>%
          addControl(
            html = "The matrix dimensions do not match the number of map units, so a matrix map preview cannot be generated.",
            position = "topright"
          )
      )
    }

    mat <- as.matrix(mat)
    mat[is.na(mat)] <- 0
    if (nrow(mat) == ncol(mat)) diag(mat) <- 0
    nonzero_vals <- unique(as.numeric(mat[mat != 0]))
    is_binary_adjacency <- length(nonzero_vals) > 0 && all(nonzero_vals %in% c(1))

    if (is_binary_adjacency) {
      coords <- suppressWarnings(sf::st_coordinates(sf::st_centroid(sf::st_geometry(shp))))
      links <- which(mat != 0, arr.ind = TRUE)
      links <- links[links[, 1] != links[, 2], , drop = FALSE]
      link_count <- nrow(links)
      if (link_count == 0) return(render_basic_map(shp))
      if (link_count > 5000) links <- links[seq_len(5000), , drop = FALSE]
      lng <- as.vector(rbind(coords[links[, 1], 1], coords[links[, 2], 1], NA))
      lat <- as.vector(rbind(coords[links[, 1], 2], coords[links[, 2], 2], NA))

      return(
        leaflet(shp) %>%
          addTiles() %>%
          addPolygons(weight = 1, color = "#756b78", fillColor = "#f8efe5", fillOpacity = 0.62, label = as.character(sf::st_drop_geometry(shp)[[1]])) %>%
          addPolylines(lng = lng, lat = lat, color = "#8b6f82", weight = 1, opacity = 0.45) %>%
          addControl(
            html = if (link_count > 5000) "0/1 adjacency matrix: showing a sample of the first 5000 links." else "0/1 adjacency matrix: showing spatial-unit links.",
            position = "bottomright"
          )
      )
    }

    weight_strength <- rowSums(abs(mat), na.rm = TRUE)
    render_basic_map(shp, values = weight_strength, label_prefix = "Weight Strength", range_legend = TRUE)
  })

  # 数据列名更新后，同步刷新Data Check页中的字段选择器。BSTVC 默认优先选择 Year/FIPS，
  # 这是 Florida 示例数据和帮助文档中的Field Name。
  observe({
    dat <- raw_data()
    if (is.null(dat)) return()

    cols <- names(dat)
    updateSelectInput(session, "check_time", choices = c("Please select" = "", cols), selected = input$check_time %||% "")
    updateSelectInput(session, "check_space", choices = c("Please select" = "", cols), selected = input$check_space %||% "")
  })

  observeEvent(input$run_check, {
    dat <- raw_data()
    shp <- study_map()

    if (is.null(dat) || is.null(shp)) {
      st <- list(code = "error", message = "Please load the modeling table data and modeling map data first.", mode = input$check_mode, time = format(Sys.time(), "%H:%M:%S"))
      check_state(st)
      check_history(c(list(st), check_history()))
      return()
    }

    map_df <- sf::st_drop_geometry(shp)

    tryCatch({
      if (is.null(input$check_space) || identical(input$check_space, "")) stop("Please select the Space field.")
      if (!input$check_space %in% names(dat)) stop("The spatial field is not in the data.")
      shp <- align_map_space_field(shp, input$check_space)
      map_df <- sf::st_drop_geometry(shp)

      if (identical(input$check_mode, "bstvc")) {
        if (is.null(input$check_time) || identical(input$check_time, "")) stop("Please select the Time field.")
        if (!input$check_time %in% names(dat)) stop("The time field is not in the data.")

        data_check_output <- capture.output({
          corrected <- BSTVC::data.check(data = dat, study_map = shp, Time = input$check_time, Space = input$check_space)
        })

        if (is.null(corrected) || !is.data.frame(corrected)) {
          checked_bstvc(NULL)
          stop(paste(c("data.check did not return data usable for modeling.", data_check_output), collapse = "\n"))
        }

        checked_bstvc(corrected)
        checked_bsvc(NULL)

        check_message <- paste(c("BSTVC::data.check completed.", data_check_output), collapse = "\n")
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
          stop("BSVC requires single-period spatial cross-sectional data, meaning each spatial unit can appear only once.")
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
          stop(paste(c("data.check did not return data usable for modeling.", data_check_output), collapse = "\n"))
        }

        corrected[[temp_time]] <- NULL
        checked_bsvc(corrected)
        checked_bstvc(NULL)

        check_message <- paste(c("BSVC spatial cross-sectional data has been checked through BSTVC::data.check using a temporary time field.", data_check_output), collapse = "\n")
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
      return(div(class = "alert alert-info", tags$h4(icon("info-circle"), " Waiting for Check"), tags$p(check_state()$message)))
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
        error = "Check Failed",
        ok_rematched = "Automatically Rearranged",
        ok_aligned = "Check Passed",
        "Check Log"
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

  check_download_data <- reactive({
    st <- check_state()
    if (!identical(st$code, "ok_rematched")) return(NULL)
    if (identical(st$mode, "bstvc")) checked_bstvc() else checked_bsvc()
  })

  output$check_download_ui <- renderUI({
    dat <- check_download_data()
    if (is.null(dat)) return(NULL)

    tags$div(
      class = "convert-banner convert-banner-bottom",
      icon("download"),
      tags$span("The system has automatically rearranged the data. Download the checked data and choose Checked Data on the modeling page."),
      tags$div(
        class = "download-tabs",
        tabsetPanel(
          id = "check_download_format",
          type = "tabs",
          tabPanel("CSV", value = "csv"),
          tabPanel("Excel xlsx", value = "xlsx"),
          tabPanel("Excel xls", value = "xls")
        )
      ),
      downloadButton("download_checked_data", "Download Checked Data", class = "btn btn-outline-success download-btn")
    )
  })

  output$download_checked_data <- downloadHandler(
    filename = function() {
      st <- check_state()
      fmt <- input$check_download_format %||% "csv"
      paste0("BSTVC_checked_", st$mode %||% "data", "_", format(Sys.Date(), "%Y%m%d"), ".", fmt)
    },
    content = function(file) {
      dat <- check_download_data()
      req(dat)
      write_single_result_table(dat, file, input$check_download_format %||% "csv")
    }
  )

  observeEvent(input$go_bstvc, {
    updateTabItems(session, inputId = "tabs", selected = "bstvc")
  })

  observeEvent(input$go_bsvc, {
    updateTabItems(session, inputId = "tabs", selected = "bsvc")
  })

  observeEvent(input$go_convert_from_overview, {
    updateTabItems(session, inputId = "tabs", selected = "spatial_convert")
  })

  observeEvent(input$go_input_from_overview, {
    updateTabItems(session, inputId = "tabs", selected = "input")
  })

  observeEvent(input$go_input_from_convert, {
    updateTabItems(session, inputId = "tabs", selected = "input")
  })

  observeEvent(input$go_input_from_field_convert, {
    updateTabItems(session, inputId = "tabs", selected = "input")
  })

  observeEvent(input$go_input_from_matrix_builder, {
    updateTabItems(session, inputId = "tabs", selected = "input")
  })

  observeEvent(input$go_check_from_input, {
    updateTabItems(session, inputId = "tabs", selected = "check")
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
# 也可以通过 shiny::runApp("File Directory") 启动。
####################################################################################################
shinyApp(ui, server)
