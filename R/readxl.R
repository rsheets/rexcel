##' Read an Excel spreadsheet the same way as readxl, but slower.
##' Assumes a well behaved table of data.
##' @title Read an Excel spreadsheet like readxl
##' @param path Path to the xlsx file
##' @param sheet Sheet name or an integer
##' @param col_names TRUE (the default) indicating we should use the
##'   first row as column names, FALSE, indicating we should generate
##'   names (X1, X2, ..., Xn) or a character vector of names to apply.
##' @param col_types Either NULL (the default) indicating we should
##'   guess the column types or a vector of column types (values must
##'   be "blank", "numeric", "date" or "text").
##' @param na Values indicating missing values (if different from
##'   blank).  Not yet used.
##' @param skip Number of rows to skip.
##' @export
rexcel_readxl <- function(path, sheet=1L, col_names=TRUE,
                          col_types=NULL, na="", skip=0) {
  dat <- rexcel_read(path, sheet)
  linen::worksheet_to_table(dat, col_names, col_types, na, skip)
}
