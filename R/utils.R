vlapply <- function(X, FUN, ...) {
  vapply(X, FUN, logical(1), ...)
}
viapply <- function(X, FUN, ...) {
  vapply(X, FUN, integer(1), ...)
}
vnapply <- function(X, FUN, ...) {
  vapply(X, FUN, numeric(1), ...)
}
vcapply <- function(X, FUN, ...) {
  vapply(X, FUN, character(1), ...)
}

attr_bool <- function(x, missing=NA) {
  if (is.null(x)) missing else as.logical(as.integer(x))
}

attr_integer <- function(x, missing=NA_integer_) {
  if (is.null(x)) missing else as.integer(x)
}

attr_numeric <- function(x, missing=NA_real_) {
  if (is.null(x)) missing else as.numeric(x)
}

attr_character <- function(x, missing=NA_character_) {
  if (is.null(x)) missing else x
}

`%||%` <- function(a, b) {
  if (is.null(a)) b else a
}

process_container <- function(xml, xpath, ns, fun, ..., classes=NULL) {
  els <- xml2::xml_children(xml2::xml_find_one(xml, xpath, ns))
  rbind_df(lapply(els, fun, ns, ...), classes)
}

## The function below is a faster version of
##
##   tibble::as_data_frame(do.call("rbind", x, quote=TRUE))
##
## But it avoids constructing a very hard to validate, slow to run
## function (on the order of a second), but it's not terrible nice to
## look at or understand.
rbind_df <- function(x, classes=NULL) {
  if (length(x) == 0L) {
    return(tibble_empty_data_frame(classes))
  }
  nms <- names(x[[1L]])
  xx <- unlist(x, FALSE)
  dim(xx) <- c(length(nms), length(x))
  if (is.null(classes)) {
    preserve <- logical(length(nms))
  } else {
    preserve <- classes == "list"
  }
  ul <- function(i, x) {
    if (preserve[[i]]) x else unlist(x)
  }
  tmp <- setNames(lapply(seq_along(nms), function(i) ul(i, xx[i, ])), nms)
  tibble::as_data_frame(tmp)
}

tibble_empty_data_frame <- function(classes) {
  if (is.null(classes)) {
    ## NOTE: Once things settle down, this can be dropped.
    stop("deal with me")
    tibble::data_frame()
  } else {
    tibble::as_data_frame(lapply(classes, vector))
  }
}
