here <- rprojroot::find_package_root_file

TEST_DIR <- here("tests", "testthat")
TEST_REF_DIR <- file.path(TEST_DIR, "reference")
if (!dir.exists(TEST_REF_DIR)) dir.create(TEST_REF_DIR)

## Get readxl source in order to get tests and therefore test sheets.
get_readxl <- function(path = TEST_REF_DIR) {
  readxl_path <- file.path(path, "readxl")
  if (dir.exists(readxl_path)) {
    return(readxl_path)
  }
  if (!has_internet()) {
    return(tempfile())
  }
  url <- "https://cran.rstudio.com/src/contrib/readxl_0.1.1.tar.gz"
  dest <- tempfile()
  tryCatch(download.file(url, dest),
           error = function(e) skip(e$message))
  on.exit(file.remove(dest))
  untar(dest, exdir = path)
  readxl_path
}

## This is not terrifically portable (windows does not support it, but
## it gives a more graceful behaviour on Linux / OSX when there is no
## internet)
has_internet <- function() {
  !is.null(suppressWarnings(nsl("www.google.com")))
}
