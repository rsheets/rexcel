xlsx_read_Content_Types <- function(path) {
  ct <- xlsx_read_file(path, "[Content_Types].xml")
  node_att <- lapply(xml2::xml_contents(ct), xml_attrs_list)
  tibble::data_frame(
    PartName = vcapply2(node_att, "PartName"),
    Extension = vcapply2(node_att, "Extension"),
    ContentType = vcapply2(node_att, "ContentType")
  )
}

xlsx_read_workbook_sheets <- function(path) {
  xml <- xlsx_read_file(path, "xl/workbook.xml")
  ## why do I write this weird XPath?
  ## namespace avoidance in order to handle xlsx like
  ## Ekaterinburg_IP.xlsx from here:
  ## https://github.com/hadley/readxl/issues/80
  sheets <- xml2::xml_find_first(xml, ".//*[local-name() = 'sheets']")
  sheets_att <- lapply(xml2::xml_contents(sheets), xml_attrs_list)
  tibble::data_frame(
    name = vcapply2(sheets_att, "name"),
    state = vcapply2(sheets_att, "state"),
    sheet_id = as.integer(vcapply2(sheets_att, "sheetId")),
    id = vcapply2(sheets_att, "id")
  )
}

xlsx_read_workbook_defined_names <- function(path) {
  xml <- xlsx_read_file(path, "xl/workbook.xml")

  ## 18.14.5 definedName
  ## 18.14.6 definedNames

  ## the definedName nodes can have different XML structure
  ## Rich had one in mind during his initial pass (or just worked from spec)
  ## --> all info found in attributes (name, refersTo, sheedId)
  ## which motivated
  ## xlsx_ct_external_defined_names()
  ## Jenny sees a different structure in the example sheet she created
  ## inst/sheets/defined_names.xlsx
  ## --> node value gives cell ref (Rich found this in refersTo attr)
  ## --> attributes refersTo, sheetId don't even exist
  ##
  ## this is an attempt to accomodate both but Jenny doesn't have an actual
  ## example of the first form to look at
  dn_nodes <- xml2::xml_find_all(xml, "//*[local-name() = 'definedName']")
  ## why do I write this weird XPath?
  ## namespace avoidance in order to handle xlsx like
  ## Ekaterinburg_IP.xlsx from here:
  ## https://github.com/hadley/readxl/issues/80
  if (length(dn_nodes) == 0) {
    return(NULL)
  }
  dn_att <- lapply(dn_nodes, xml_attrs_list)

  ## this is where the spec suggests you will find the cell refs/ranges
  refers_to <- vcapply2(dn_att, "refersTo")
  ## and yet Jenny finds them as node text ...
  if (all(is.na(refers_to))) {
    refers_to <- xml2::xml_text(dn_nodes)
  }

  ## may just be NAs
  sheet_id <- as.integer(vcapply2(dn_att, "sheetId"))

  tibble::data_frame(
     name = vcapply2(dn_att, "name"),
     refers_to,
     sheet_id
  )
}

xlsx_read_workbook_rels <- function(path) {
  ## do we really have to worry about this file not existing?
  xml <- xlsx_read_file_if_exists(path, "xl/_rels/workbook.xml.rels")
  if (is.null(xml)) {
    return(NULL)
  }
  rel_nodes <- xml2::xml_children(xml)
  rels <- rbind_df(lapply(rel_nodes, xml_attrs_list))
  names(rels) <- tolower(names(rels))
  rels
  ## MAYBE TODOs, if decide to do more processing:
  ##
  ## prepend target with "xl/"
  ## but check the type and don't to if an external reference
  ##
  ## just take the last bit of type, i.e.basename(type),
}

xlsx_read_worksheet_rels <- function(path) {
  manifest <- xlsx_list_files(path)
  holds_sheet_rels <-
    grepl("xl/worksheets/_rels/sheet[0-9]*.xml.rels", manifest$Name)
  if (none(holds_sheet_rels)) {
    return(NULL)
  }
  sheet_rels_fnames <- manifest$Name[holds_sheet_rels]
  nms <- gsub("xl/worksheets/_rels/(sheet[0-9]+).xml.rels", "\\1",
              sheet_rels_fnames)
  wr <- sheet_rels_fnames %>%
    purrr::map(xlsx_read_file, path = path)
  worksheet_rels <- wr %>%
    purrr::map(xml2::xml_find_all, xpath = "//d1:Relationship") %>%
    stats::setNames(nms)
  f <- function(x) {
    x %>%
      purrr::map(xml2::xml_attrs) %>%
      purrr::map(as.list) %>%
      dplyr::bind_rows()
  }
  worksheet_rels %>%
    purrr::map(f) %>%
    dplyr::bind_rows(.id = "worksheet") %>%
    stats::setNames(tolower(names(.)))
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

## 18.14.6 definedNames
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
