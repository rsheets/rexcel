xlsx_read_Content_Types <- function(path) {
  ct <- xlsx_read_file(path, "[Content_Types].xml")
  node_att <- lapply(xml2::xml_contents(ct), xml_attrs_list)
  ## some version of this should probably go in utils.R eventually?
  f <- function(l, nm) {
    ret <- lapply(l, `[[`, nm)
    ret <- lapply(ret, `%||%`, NA_character_)
    unlist(ret)
  }
  tibble::data_frame(
    PartName = f(node_att, "PartName"),
    Extension = f(node_att, "Extension"),
    ContentType = f(node_att, "ContentType")
  )
}

xlsx_read_workbook <- function(path) {
  ## TODO: Consider what do do when rels is NULL; do we throw?
  rels <- xlsx_read_rels(path, "xl/workbook.xml")
  xml <- xlsx_read_file(path, "xl/workbook.xml")
  ns <- xml2::xml_ns(xml)

  defined_names <- xlsx_ct_external_defined_names(xml, ns)
  sheets <- xlsx_ct_sheets(xml, ns, rels)

  list(rels=rels, sheets=sheets, defined_names=defined_names)
}

## 18.2.20 sheets
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

## 18.2.19 sheet
xlsx_ct_sheet <- function(xml, ns) {
  at <- as.list(xml2::xml_attrs(xml))
  tibble::data_frame(
    name = attr_character(at$name),
    sheet_id = attr_integer(at$sheetId),
    state = attr_character(at$state, "visible"),
    ref = attr_character(at[["id"]]))
}

## 18.14.6 definedName
xlsx_ct_external_defined_names <- function(xml, ns) {
  process_container(xml, "d1:definedNames", ns, xlsx_ct_external_defined_name,
                    classes=TRUE)
}

## 18.14.5 definedName
xlsx_ct_external_defined_name <- function(xml, ns) {
  at <- xml_attrs_list(xml)
  tibble::data_frame(
    name = attr_character(at$name),
    refers_to = attr_character(at$refersTo),
    sheet_id = attr_integer(at$sheetId))
}
