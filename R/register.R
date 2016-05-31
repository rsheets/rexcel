#' Register an Excel workbook
#'
#' Experimental function, in an experimental package! It covers alot of the same
#' ground as \code{\link{rexcel_read_workbook}()} with several notable
#' exceptions.
#'
#' \itemize{
#' \item Currently returns a list, not a proper linen::workbook.
#' \item Read one file at a time, make one or more objects from it, with only
#' the processing necessary to make it a reasonable R object.
#' \item For some clearly useful things, create downstream objects via
#' processing and/or combining info across primary objects.
#' }
#'
#' @param path path to xlsx
#'
#' @return a list
#' @keywords internal
#'
#' @examples
#' mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
#' rexcel_register(mini_gap_path)
#'
#' ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
#'                        package = "rexcel")
#' rexcel_register(ff_path)
#' @export
rexcel_register <- function(path) {
  ## TO DO:
  ## if path is actually a workbook
  ## Recall(path$path)
  ## i.e. refresh registration of the workbook
  ## to be used when you are concerned the xlsx has changed
  if (!is_xlsx(path)) {
    stop("`path` does not appear to point to valid xlsx:\n", path,
         call. = FALSE)
  }
  manifest      <- xlsx_list_files(path)
  ct            <- xlsx_read_Content_Types(path) ## [Content_Types].xml
  # rels        <- xlsx_read_file(path, "_rels/.rels") # always boring? skipped
  sheets        <- xlsx_read_workbook_sheets(path) # xl/workbook.xml
  defined_names <- xlsx_read_workbook_defined_names(path)
  workbook_rels <- xlsx_read_workbook_rels(path)

  ## allow ourselves to do some synthesis of the above
  ## in workbook_rels, prepend xl/ to target
  ## join to sheets on id
  ## result has one row per worksheet
  sheets_df <- join_sheets_workbook_rels(sheets, workbook_rels)

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
    lapply(xml2::xml_children)
  f <- function(font_node, ns) {
    nms <- xml2::xml_name(font_node, ns) %>% rm_xml_ns()
    vals <- xml2::xml_attrs(font_node, ns) %>% lapply(unname)
    names(vals) <- nms
    vals[lengths(vals) > 0]
  }
  fonts <- font_nodes %>%
    lapply(f, ns = ns) %>%
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

join_sheets_workbook_rels <- function(sheets, workbook_rels) {
  suppressMessages(
    sheets_df <- workbook_rels %>%
      dplyr::mutate(target = file.path("xl", target)) %>%
      dplyr::right_join(sheets) %>%
      ## use one_of() here because not all variables exist all the time
      dplyr::select(
        dplyr::one_of(c("sheet_id", "name", "id", "target", "state", "type"))
      )
  )
  unique_type <- unique(sheets_df$type)
  if (!identical(unique_type,
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")) {
    message("New type found for a worksheet! LOOK INTO THIS.")
  } else {
    sheets_df$type <- NULL
  }
  sheets_df
}
