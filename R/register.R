#' Low-level function to expose contents of xlsx
#'
#' Jenny is getting to know xlsx by writing this. Maybe it will evolve into some
#' sort of an "Excel doctor" function. It covers alot of the same ground as
#' \code{\link{rexcel_read_workbook}()} with several notable exceptions. Returns
#' a list, not a proper linen::workbook. Much less processing is done --
#' basically only whats needed to some reasonable R object, usually a data
#' frame. "Read one file at a time, make one object from each file" is the
#' general philosophy.
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
  if (!is_xlsx(path)) {
    stop("`path` does not appear to point to valid xlsx:\n", path,
         call. = FALSE)
  }
  manifest <- xlsx_list_files(path)

  ## overview of typical manifest

  ## *** workbook infrastructure ***
  ## [Content_Types].xml
  ## _rels/.rels
  ## xl/workbook.xml
  ## xl/sharedStrings.xml
  ## xl/styles.xml
  ## xl/_rels/workbook.xml.rels

  ## TO DO: look into these files that appear in defined_names.xlsx
  ## docProps/core.xml
  ## docProps/app.xml

  ## ** worksheets-related ***
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
  ct <- xlsx_read_Content_Types(path)
  #setdiff(manifest$Name, gsub("^\\/", "", ct$PartName))
  #intersect(gsub("^\\/", "", ct$PartName), manifest$Name)
  ## ct is a tbl associating content types with extensions or specific files
  ## two "general" rows for the xml and rels extensions, otherwise ...
  ## each row seems to address a specific file from the manifest (but not all)

  ## _rels/.rels
  #rels <- xlsx_read_file(path, "_rels/.rels")
  ## this appears to be always boring? omit it

  ## xl/workbook.xml
  sheets <- xlsx_read_workbook_sheets(path)
  ## sheets is a tbl with one row per worksheet and these variables:
  ## name: e.g. "Africa" (assume this is name of the tab)
  ## state: "visible" (or what else ... "invisible"?), might be NA
  ## sheet_id: integer (assume this is order perceived by user)
  ## id: character, e.g. "rId5" (a key that comes up in other tables)

  defined_names <- xlsx_read_workbook_defined_names(path)
  ## defined_names is a tbl with one row per named range and these variables:
  ## name: name of the named range
  ## refers_to: string representation of the cell area reference, e.g. Sheet1!$B$2:$B$11
  ## sheet_id: integer (I can't get my hands on a sheet that has actually this)
  ## defined_names will be NULL if there are no named ranges
  ## TO DO:
  ## it's possible there should be more info, because I've seen named range xml
  ## that is more complicated
  ## https://github.com/rsheets/cellranger/issues/23#issuecomment-221898917

  ## xl/_rels/workbook.xml.rels
  workbook_rels <- xlsx_read_workbook_rels(path)
  ## workbook_rels is a tibble, each row a file, with variables
  ## id: character, e.g. "rId5" (a key that occurs elsewhere)
  ## target: a file path relative to xl/ (maybe prepend xl/?)
  ## type: a long namespace-y string, the last bit of which tells you
  ## if the associated file is sharedStrings, styles, or a worksheet, e.g.,
  ## http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet
  ##   (maybe just take the last bit?)

  ## xl/sharedStrings.xml
  shared_strings <- xlsx_read_shared_strings(path)
  ## shared_strings is a character of shared strings
  ## with attributes count (total # of strings?), uniqueCount (its own length?)

  ## xl/styles.xml
  styles <- xlsx_read_file(path, "xl/styles.xml")
  ## again, ekaterinburg has different namespace
  ## this is a mess; discuss with rich
  ns <- xml2::xml_ns(styles)
  alt_ns <-
    construct_xml_ns(x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  if (ns_equal_to_ref(styles, alt_ns)) {
    ns <- xml2::xml_ns_rename(ns, x = "d1")
  }
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

  ## I don't feel like parsing the remaining elements of styles for now
  ## no temptation to duplicate rich's efforts there ... yikes
  fills <- NULL
  borders <- NULL
  cell_style_xfs <- NULL
  cell_xfs <- NULL
  cell_styles <- NULL
  num_fmts <- NULL
  dxfs <- NULL

  ## xl/worksheets/_rels/sheet1.xml.rels and friends
  worksheet_rels <- xlsx_read_worksheet_rels(path)
  ## worksheet_rels is a tibble, each row a file ... so far one row per
  ## worksheet, though that might not hold in general, with variables
  ## worksheet: character, e.g. "sheet1" (I added this!)
  ## Id: character, so far uniformly "rId1"
  ## Type: uniformly a long name-spacey string ending in "drawing" or "hyperlink"
  ## Target: hmmm .... seems to vary
  ##    in one example: path to the corresponding worksheetdrawingX.xml file
  ##    ^^ maybe that's the default? when there's nothing else?
  ##    in another: "http://www.google.com/", which appears in the sheet
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

  ## use synthesis of the above:
  ## one row per sheet
  ## everything from sheets tbl already formed
  ## workbook_rels prepend xl/ to Target
  ## join to sheets on Id
  sheets_df <- workbook_rels %>%
    dplyr::mutate(target = file.path("xl", target)) %>%
    dplyr::right_join(sheets, by = c("id" = "id")) %>%
    dplyr::select(
      dplyr::one_of(c("sheet_id", "name", "id", "target", "state", "type"))
    )

  dplyr::lst(xlsx_path = path,
      reg_time = Sys.time(),
      manifest,
      content_types = ct,
      sheets,
      defined_names,
      workbook_rels,
      shared_strings,
      styles =
        dplyr::lst(fonts, fills, borders, cell_style_xfs, cell_xfs,
                   cell_styles, num_fmts, dxfs),
      worksheet_rels,
      sheets_df
  )
  ## TODO: obviously this should return a workbook object!
}
