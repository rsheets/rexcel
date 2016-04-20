xlsx_read_workbook <- function(path) {
  ## TODO: Consider what do do when rels is NULL; do we throw?
  rels <- xlsx_read_rels(path, "xl/workbook.xml")
  xml <- xlsx_read_file(path, "xl/workbook.xml")
  ns <- xml2::xml_ns(xml)
  sheets <- xlsx_ct_sheets(xml, ns, rels)

  list(rels=rels, sheets=sheets)
}

xlsx_ct_sheets <- function(xml, ns, rels) {
  dat <- process_container(xml, "d1:sheets", ns, xlsx_ct_sheet)

  if (is.null(rels)) {
    stop("FIXME")
  } else {
    i <- match(dat$ref, rels$id)
    dat <- cbind(dat, rels[i, -1L])
  }

  dat
}

xlsx_ct_sheet <- function(xml, ns) {
  at <- as.list(xml2::xml_attrs(xml))
  tibble::data_frame(
    name = attr_character(at$name),
    sheet_id = attr_integer(at$sheetId),
    state = attr_character(at$state, "visible"),
    ref = attr_character(at[["id"]]))
}
