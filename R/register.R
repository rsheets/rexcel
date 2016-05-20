#' Low-level function to expose contents of xlsx
#'
#' This is just Jenny getting to know xlsx! Function returns alot of the same
#' information as rexcel_read_workbook() but with several notable exceptions.
#' Returns a list, not a proper linen::workbook. Much less processing is done --
#' basically only whats needed to some reasonable R object, usually a data
#' frame.
#'
#' @param path
#'
#' @return a list
#' @keywords internaln
#'
#' @examples
#' mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
#' rexcel_workbook(mini_gap_path)
#'
#' ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
#'                        package = "rexcel")
#' rexcel_workbook(ff_path)
rexcel_workbook <- function(path) {
  ## TO DO:
  ## if path is actually a workbook
  ## Recall(path$path)
  ## i.e. refresh registration of the workbook
  ## to be used when you are concerned the xlsx has changed
  is_xlsx(path)
  manifest <- xlsx_list_files(path)

  ## overview of typical manifest

   ## [Content_Types].xml
  ## _rels/.rels
  ## xl/workbook.xml
  ## xl/sharedStrings.xml
  ## xl/styles.xml
  ## xl/_rels/workbook.xml.rels

  ## xl/worksheets/sheet1.xml
  ## xl/worksheets/sheet2.xml
  ##  ... and so on
  ## xl/worksheets/_rels/sheet1.xml.rels
  ## xl/worksheets/_rels/sheet2.xml.rels
  ##  ... and so on
  ## xl/drawings/worksheetdrawing1.xml
  ## xl/drawings/worksheetdrawing2.xml
  ##  ... and so on

  ## [Content_Types].xml
  ct <- xlsx_read_file(path, "[Content_Types].xml") %>%
    xml2::xml_contents() %>%
    xml2::xml_attrs() %>%
    purrr::map(as.list) %>%
    dplyr::bind_rows() %>%
    dplyr::select(PartName, Extension, ContentType)
  #setdiff(manifest$Name, gsub("^\\/", "", ct$PartName))
  #intersect(gsub("^\\/", "", ct$PartName), manifest$Name)
  ## ct is a tbl associating content types with extensions or files
  ## for the most part, each row addresses a specific file in manifest
  ## except all "rels" files are represented by a single row
  ## and there's a row that says xml files are "application/xml"

  ## _rels/.rels
  #rels <- xlsx_read_file(path, "_rels/.rels")
  ## this appears to be always boring? omit it

  ## xl/workbook.xml
  xl_workbook <- xlsx_read_file(path, "xl/workbook.xml")
  sheets <- xl_workbook %>%
    xml2::xml_find_one("//d1:sheets", xml2::xml_ns(.)) %>%
    xml2::xml_contents() %>%
    xml2::xml_attrs() %>%
    purrr::map(as.list) %>%
    dplyr::bind_rows() %>%
    dplyr::mutate(sheetId = as.integer(sheetId))
  ## sheets is a tbl with one row per worksheet and these variables:
  ## state: "visible" (or what else ... "invisible"?)
  ## name: e.g. "Africa" (assume this is name of the tab)
  ## sheetID: integer (assume this is order perceived by user)
  ## id: character, e.g. "rId5" (a key that comes up in other tables)

  ## xl/sharedStrings.xml
  shared_strings <- xlsx_read_file(path, "xl/sharedStrings.xml")
  shared_strings_att <- xml2::xml_attrs(shared_strings) %>%
    as.list() %>%
    purrr::map(as.integer)
  shared_strings <- shared_strings %>%
    xml2::xml_find_all("//d1:t", xml2::xml_ns(.)) %>%
    purrr::map_chr(xml2::xml_text)
  attributes(shared_strings) <- shared_strings_att
  ## sh_strings is a character of shared strings
  ## with attributes count (total # of strings?), uniqueCount (its own length?)

  ## xl/styles.xml
  styles <- xlsx_read_file(path, "xl/styles.xml")
  ns <- xml2::xml_ns(styles)

  font_nodes <- styles %>%
    xml2::xml_find_all("//d1:fonts/d1:font", ns) %>%
    purrr::map(xml2::xml_children)
  f <- function(font_node, ns) {
    nms <- xml2::xml_name(font_node, ns) %>% rm_xml_ns()
    vals <- xml2::xml_attrs(font_node, ns) %>% purrr::map(unname)
    setNames(vals, nms) %>%
      purrr::keep(~length(.x) > 0)
  }
  fonts <- font_nodes %>%
    purrr::map(f, ns = ns) %>%
    dplyr::bind_rows()
  ## fonts is a tbl with one row per font and variables such as
  ## sz, color, name

  ## I'm don't feel like parsing the remaining elements of styles for now
  ## no temptation to duplicate rich's efforts there ... yikes
  fills <- NULL
  borders <- NULL
  cell_style_xfs <- NULL
  cell_xfs <- NULL
  cell_styles <- NULL
  num_fmts <- NULL
  dxfs <- NULL

  ## xl/_rels/workbook.xml.rels
  workbook_rels <- xlsx_read_file(path, "xl/_rels/workbook.xml.rels") %>%
    xml2::xml_contents() %>%
    xml2::xml_attrs() %>%
    purrr::map(as.list) %>%
    dplyr::bind_rows() %>%
    dplyr::select(Id, Target, Type)
  ## workbook_rels is a tibble, each row a file, with variables
  ## Id: character, e.g. "rId5" (a key that came up already in sheets above)
  ## Target: a file path relative to xl/
  ## Type: an long namespace-y string, the last bit of which tells you
  ## if the associated file is sharedStrings, styles, or a worksheet, e.g.,
  ## http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet

  ## xl/worksheets/_rels/sheet1.xml.rels and friends
  fnames <- manifest %>%
    dplyr::filter(grepl("rels", Name), grepl("sheet", Name)) %>%
    .$Name
  nms <- gsub("xl/worksheets/_rels/(sheet[0-9]+).xml.rels", "\\1", fnames)
  wr <- fnames %>%
    purrr::map(xlsx_read_file, path = path)
  ns <- xml2::xml_ns(wr[[1]])
  worksheet_rels <-
    wr %>%
    purrr::map(xml2::xml_find_one, xpath = "//d1:Relationship", ns = ns) %>%
    purrr::map(xml2::xml_attrs, ns = ns) %>%
    purrr::map(as.list) %>%
    dplyr::bind_rows() %>%
    dplyr::mutate(sheet = nms) %>%
    dplyr::select(sheet, Target, dplyr::everything())
  ## worksheet_rels is a tibble, each row a file ... so far one row per
  ## worksheet, though that might not hold in general, with variables
  ## sheet: character, e.g. "sheet1" (I added this!)
  ## Target: hmmm .... seems to vary
  ##    in one example: path to the corresponding worksheetdrawingX.xml file
  ##    ^^ maybe that's the default? when there's nothing else?
  ##    in another: "http://www.google.com/", which appears in the sheet
  ## Id: character, so far uniformly "rId1"
  ## Type: uniformly a long name-spacey string ending in "drawing" or "hyperlink"
  ## TargetMode: (seen in one example) "External" for the hyperlink

  ## xl/worksheets/sheet1.xml etc.
  #one_sheet <- xlsx_read_file(path, "xl/worksheets/sheet1.xml")
  ## come back here and parse enough to learn worksheet extent
  ## otherwise, I don't see anything here that belongs in top-level workbook
  ## creation

  ## xl/drawings/worksheetdrawing1.xml etc.
  #one_drawing <- xlsx_read_file(path, "xl/drawings/worksheetdrawing1.xml")
  ## I don't see anything here that belongs in top-level workbook creation
  ## also, in my toy examples with no charts, this consists only of namespaces

  ## come back here and make a new object
  ## one row per sheet
  ## everything from sheets tbl already formed
  ## workbook_rels prepend xl/ to Target
  ## join to sheets on Id
  sheets_df <- workbook_rels %>%
    dplyr::mutate(Target = file.path("xl", Target)) %>%
    dplyr::right_join(sheets, by = c("Id" = "id")) %>%
    dplyr::select(sheetId, name, Id, Target, Type)

  dplyr::lst(xlsx_path = path,
      reg_time = Sys.time(),
      manifest,
      content_types = ct,
      sheets,
      sheets_df,
      shared_strings,
      styles =
        dplyr::lst(fonts, fills, borders, cell_style_xfs, cell_xfs,
                   cell_styles, num_fmts, dxfs),
      workbook_rels,
      worksheet_rels
  )
  ## TODO: obviously this should return a workbook object!
}
