xlsx_read_workbook <- function(path) {
  ## TODO: Consider what do do when rels is NULL; do we throw?
  rels <- xlsx_read_rels(path, "xl/workbook.xml")
  xml <- xlsx_read_file(path, "xl/workbook.xml")
  ns <- xml2::xml_ns(xml)

  defined_names <- xlsx_ct_external_defined_names(xml, ns)
  sheets <- xlsx_ct_sheets(xml, ns, rels)

  list(rels=rels, sheets=sheets, defined_names=defined_names)
}

xlsx_namespace <- function(ns) {
  url <- "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  names(ns)[[match(url, ns)]]
}
xlsx_name <- function(name, ns) {
  paste0(xlsx_namespace(ns), ":", name)
}

## 18.2.20 sheets
xlsx_ct_sheets <- function(xml, ns, rels) {
  ## Apparently new xml2 has some facilities for dealing with
  ## namespaces which might make this easier.  Or break everything in
  ## here.  Or perhaps a little of both.
  dat <- process_container(xml, xlsx_name("sheets", ns), ns, xlsx_ct_sheet)

  if (is.null(rels)) {
    stop("FIXME")
  } else {
    i <- match(dat$ref, rels$id)
    dat <- cbind(dat, rels[i, -1L])
  }

  dat
}

## 18.2.19 sheet
xlsx_ct_sheet <- function(xml, ns) {
  at <- as.list(xml2::xml_attrs(xml))
  tibble::tibble(
    name = attr_character(at$name),
    sheet_id = attr_integer(at$sheetId),
    state = attr_character(at$state, "visible"),
    ref = attr_character(at[["id"]]))
}

## 18.14.6 definedName
xlsx_ct_external_defined_names <- function(xml, ns) {
  process_container(xml, xlsx_name("definedNames", ns), ns,
                    xlsx_ct_external_defined_name, classes=TRUE)
}

## 18.14.5 definedName
xlsx_ct_external_defined_name <- function(xml, ns) {
  at <- xml_attrs_list(xml)
  tibble::tibble(
    name = attr_character(at$name),
    refers_to = attr_character(at$refersTo),
    sheet_id = attr_integer(at$sheetId))
}
