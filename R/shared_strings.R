## Shared string table:
##
##   [-] 18.4.1  charset (Character Set) (ignored for now)
##   [x] 18.4.2  outline (Outline) -- xlsx_ct_boolean_property
##   [-] 18.4.3  phoneticPr (Phonetic Properties)
##   [x] 18.4.4  r (Rich Text Run) -- xlsx_ct_rst
##   [-] 18.4.5  rFont (Font)
##   [-] 18.4.6  rPh (Phonetic Run)
##   [-] 18.4.7  rPr (Run Properties)
##   [x] 18.4.8  si (String Item) -- xlsx_ct_rst
##   [x] 18.4.9  sst (Shared String Table) -- xlsx_read_shared_strings
##   [x] 18.4.10 strike (Strike Through) -- xlsx_ct_boolean_property
##   [x] 18.4.11 sz (Font Size) -- xlsx_ct_font_size
##   [x] 18.4.12 t (Text) -- in xlsx_ct_rst
##   [x] 18.4.13 u (Underline) -- xlsx_ct_underline_property
##   [-] 18.4.14 vertAlign (Vertical Alignment) (ignored for now)

## If the format is <si>/<t> then we can just take the text values.
## Otherwise we'll have to parse out the RTF strings separately.
xlsx_read_shared_strings <- function(path) {
  xml <- xlsx_read_file_if_exists(path, "xl/sharedStrings.xml")
  if (is.null(xml)) {
    return(character(0))
  }
  vcapply(xml2::xml_children(xml), xlsx_ct_rst, xml2::xml_ns(xml))
}

## 18.4.8 si
##
## This is the core function that reads a string item (si).  The spec
## is a bit vague on this, but it seems most likely that the element
## can contain either a 't' or a bunch of 'r' elements, but not both.
##
## NOTE: Ignoring rPh and phoneticPr which might be part of this
## element.  Terribly anglocentric :(
xlsx_ct_rst <- function(x, ns) {
  t <- xml2::xml_find_first(x, xlsx_name("t", ns), ns)
  r <- xml2::xml_find_all(x, xlsx_name("r", ns), ns)
  if (length(r) == 0L) {
    ## Treat as plain text.
    ## 18.4.12 t -- ST_Xstring
    ##
    ## The only complication here is that we *might* contain the flag:
    ## xml:space which is a W3C defined thing indicating if whitespace
    ## is relevant.
    str <- xml2::xml_text(t)
  } else {
    ## NOTE: we totally ignore sub-string formatting.
    str <- if (inherits(t, "xml_missing")) "" else xml2::xml_text(t)
    if (length(r) > 0L) {
      str <- paste(c(str, xml2::xml_text(
                            xml2::xml_find_all(r, xlsx_name("t", ns), ns))),
                   collapse="")
    }
  }

  ## NOTE: I am still getting slightly different line endings to
  ## readxl because I need to convert \n -> \r\n to match
  ##
  ## Unescape the strings (ST_Xstring) See 22.9.2.19 [p3786]
  re <- "_x([[:xdigit:]]{4})_"
  i <- regexpr(re, str, perl = TRUE)
  len <- nchar(str)
  while (i > 0) {
    repl <- intToUtf8(as.integer(paste0("0x", substr(str, i + 2, i + 5))))
    str <- sub(re, repl, str)
    ## This bit of faffery stops an escaped '_' character being
    ## counted as an unescaped '_' character, and is tested.  This
    ## would be heaps easier in languages with char-by-char string
    ## handling
    j <- regexpr(re, substr(str, i + 1, len))
    i <- if (j > 0) i + j else j
  }

  str
}

## 18.4.11 sz (Font Size)
xlsx_ct_font_size <- function(sz) {
  as.numeric(xml2::xml_attr(sz, "val"))
}

## 18.4.13 u (Underline)
xlsx_ct_underline_property <- function(u, missing="none") {
  if (inherits(u, "xml_missing")) {
    missing
  } else {
    val <- xml2::xml_attr(u, "val")
    if (is.na(val)) "single" else val
  }
}
