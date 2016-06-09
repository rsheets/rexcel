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
  ns <- xml2::xml_ns(xml)
  ## to deal w/ less common namespacing, e.g. Ekaterinburg sheet
  alt_ns <-
    construct_xml_ns(x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  if (ns_equal_to_ref(ns, alt_ns)) {
    ns <- xml2::xml_ns_rename(ns, x = "d1")
  }
  string_items <- xml2::xml_children(xml)
  ret <- vcapply(string_items, xlsx_ct_rst, ns)
  ## these gymnastics are necessary to preserve attribute names
  at <- as.list(xml2::xml_attrs(xml)[c("count", "uniqueCount")])
  at <- lapply(at, as.integer)
  attributes(ret) <- at
  ret
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
  t <- xml2::xml_find_first(x, "d1:t", ns)
  r <- xml2::xml_find_all(x, "d1:r", ns)
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
      str <- paste(c(str, xml2::xml_text(xml2::xml_find_all(r, "d1:t", ns))),
                   collapse="")
    }
  }
  ## TODO: still need to "unescape" these.
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
