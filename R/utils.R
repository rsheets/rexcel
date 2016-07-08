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
  els <- xml2::xml_children(xml2::xml_find_first(xml, xpath, ns))
  if (isTRUE(classes)) {
    if (length(els) == 0L) {
      classes <- vcapply(fun(NULL, ns, ...), storage.mode)
    } else {
      classes <- NULL
    }
  }
  rbind_df(lapply(els, fun, ns, ...), classes)
}

## The function below is a faster version of
##
##   tibble::as_tibble(do.call("rbind", x, quote=TRUE))
##
## But it avoids constructing a very hard to validate, slow to run
## function (on the order of a second), but it's not terrible nice to
## look at or understand.
rbind_df <- function(x, classes=NULL) {
  if (length(x) == 0L) {
    return(empty_tibble(classes))
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
  tmp <- stats::setNames(lapply(seq_along(nms), function(i) ul(i, xx[i, ])), nms)
  tibble::as_tibble(tmp)
}

empty_tibble <- function(classes = NULL) {
  if (is.null(classes)) {
    ## NOTE: Once things settle down, this can be dropped.
    stop("deal with me")
    tibble::tibble()
  } else {
    tibble::as_tibble(lapply(classes, vector))
  }
}

progress <- function(fmt, total, ..., show=TRUE) {
  if (show && total > 0L) {
    pb <- progress::progress_bar$new(fmt, total=total)
    function(len=1) {
      invisible(pb$tick(len))
    }
  } else {
    function(...) {}
  }
}

path_join <- function(a, b) {
  na <- length(a)
  nb <- length(b)
  if (na == 1L && nb != 1L) {
    a <- rep_len(a, nb)
  } else if (nb == 1L && na != 1L) {
    b <- rep_len(b, na)
  } else if (na != nb && na != 1L && nb != 1L) {
    stop("Can't recycle vectors together")
  }

  i <- regexpr("(\\.\\./)+", b)
  len <- attr(i, "match.length", exact=TRUE)
  j <- len > 0L
  if (any(j)) {
    b[j] <- substr(b[j], len[j] + 1L, nchar(b[j]))
    len[j] <- len[j] / 3

    tmp <- strsplit(a[j], "/", fixed=TRUE)
    for (k in seq_along(tmp)) {
      ii <- length(tmp[[k]]) - len[j][[k]]
      if (ii < 0L) {
        ii <- 0L
        ## TODO: this turns up once in the Enron corpus
        ##   warning("Cannot resolve internal reference; above workbook?")
      }
      tmp[[k]] <- paste(tmp[[k]][seq_len(ii)], collapse="/")
    }
    a[j] <- unlist(tmp)
  }
  paste(a, b, sep="/")
}

## TODO: replace as.list(xml2::xml_attrs(...)) with this where NULL
## values are OK.
xml_attrs_list <- function(x) {
  if (is.null(x)) {
    structure(list(), names=character())
  } else {
    as.list(xml2::xml_attrs(x))
  }
}

as_na <- function(x) {
  ret <- NA
  storage.mode(ret) <- storage.mode(x)
  ret
}

is_xlsx <- function(path) {
  if (!file.exists(path)) {
    stop("\n", path, "\ndoes not exist")
  }
  ## TO DO: verify it's a zip archive? only way I know is unix `file` command
  ## http://officeopenxml.com/anatomyofOOXML-xlsx.php
  ## https://msdn.microsoft.com/en-us/library/office/gg278316.aspx#MinWBScenario
  files <- xlsx_list_files(path)
  has_content_types <- "[Content_Types].xml" %in% files$name
  has_rels <- "_rels/.rels" %in% files$name
  has_workbook_xml <- "xl/workbook.xml" %in% files$name
  has_sheet <- any(grepl("xl/worksheets/sheet[0-9]*.xml", files$name))
  has_content_types && has_rels && has_workbook_xml && has_sheet
}
