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

  ## back to the slog
  shared_strings <- xlsx_read_shared_strings(path)
  ## reverting to Rich's work here and god bless him for it
  styles <- xlsx_read_style(path)
  ## xl/worksheets/_rels/sheet1.xml.rels and friends
  worksheet_rels <- xlsx_read_worksheet_rels(path)

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
      styles,
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
