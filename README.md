# BSTVC-Shiny

BSTVC-Shiny is an interactive R Shiny application for Bayesian spatio-temporal varying coefficient modeling.

Users can launch the app directly from GitHub in R or RStudio. No manual download of this repository is required.

## Quick Start

Run the following code in R or RStudio:

```r
if (!require("shiny")) install.packages("shiny")
shiny::runGitHub("BSTVC-Shiny", "tangxxxxt")
```

The Shiny interface will open automatically after the repository is downloaded and loaded by R.

## Repository Structure

```text
BSTVC-Shiny/
  app.R                  # Main Shiny app file
  www/                   # Images, CSS, and static resources used by the app
  local_packages/BSTVC/  # Local BSTVC R package used by the app
  README.md
  LICENSE
```

## Notes

- Make sure your internet connection is available when running `shiny::runGitHub()`.
- The app is designed to run in a local R session through GitHub.
- If R reports missing packages, install the required packages first and run the app again.
- For large datasets or long model-fitting tasks, running the app on a computer with sufficient memory is recommended.

## License

This project is licensed under the Apache License 2.0.

