## This is done empirically rather than with the spec because I've not
## worked out how it fits into the design...
xlsx_read_rels <- function(path) {
  ## TODO: Looks like this does always exist.
  xml <- xlsx_read_file_if_exists(path, "xl/_rels/workbook.xml.rels")

  if (is.null(xml)) {
    NULL
  } else {
    ## TODO: These are allowed to be external references I think...
    rbind_df(lapply(xml2::xml_children(xml), xlsx_parse_relationship))
  }
}

xlsx_parse_relationship <- function(x) {
  at <- as.list(xml2::xml_attrs(x))
  tibble::data_frame(
    id = attr_character(at$Id),
    type = basename(at$Type),
    target = at$Target)
}
