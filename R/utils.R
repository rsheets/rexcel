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

## I know this needs a better name
## or to be unified with vcapply
vcapply2 <- function(l, nm) {
  ret <- lapply(l, `[[`, nm)
  ret <- lapply(ret, `%||%`, NA_character_)
  unlist(ret)
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

none <- function(x) !any(x)

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
  tmp <- stats::setNames(lapply(seq_along(nms), function(i) ul(i, xx[i, ])), nms)
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
      tmp[[k]] <- paste(tmp[[k]][seq_len(length(tmp[[k]]) - len[j][[k]])],
                        collapse="/")
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
  has_content_types <- "[Content_Types].xml" %in% files$Name
  has_rels <- "_rels/.rels" %in% files$Name
  has_workbook_xml <- "xl/workbook.xml" %in% files$Name
  has_sheet <- any(grepl("xl/worksheets/sheet[0-9]*.xml", files$Name))
  has_content_types && has_rels && has_workbook_xml && has_sheet
}

rm_xml_ns <- function(x) gsub(".*:(.*)", "\\1", x)

construct_xml_ns <- function(...) {
  ddd <- list(...)
  ns <- vapply(ddd, `[[`, character(1), 1)
  structure(ns, class = "xml_namespace")
}

ns_equal_to_ref <- function(xml, ref_ns) {
  if (inherits(xml, "xml_node")) {
    ns <- xml2::xml_ns(xml2::xml_root(xml))
    return(identical(ns[order(names(ns))], ref_ns[order(names(ref_ns))]))
  }
  FALSE
}

## TO DO: replace with the real parser from cellranger once it's exposed
## https://github.com/rsheets/cellranger/issues/22
extract_sheet <- function(x) {
  rx <- "^(?:(?:^\\[([^\\]]+)\\])?(?:'?([^']+)'?!)?([a-zA-Z0-9:\\-$\\[\\]]+)|(.*))$"
  ## from rematch package
  m <- regexpr(rx, x, perl = TRUE)
  res <- cbind(ifelse(m == -1, NA_character_,
                      substr(x, m, m + attr(m, "match.length") - 1)))
  res <- cbind(res,
               rbind(vapply(
                 seq_len(NCOL(attr(m, "capture.start"))),
                 function(i) {
                   start <- attr(m, "capture.start")[,i]
                   len <- attr(m, "capture.length")[,i]
                   end <- start + len - 1
                   res <- substr(x, start, end)
                   res[ start == -1 ] <- NA_character_
                   res
                 },
                 character(length(m))
               ))
  )
  res[ , 3]
}
