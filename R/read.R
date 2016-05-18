##' This function does not get the data into a usable form but at
##' least loads it up into R so we can poke about with it.  The
##' resulting loaded data can distinguish between formulae and data,
##' numbers and text.  Merged cells are detected.  Style information
##' is included, though nothing is done with it yet.  A summary of the
##' data is printed if you print the resulting object.
##'
##' @title Read an xlsx file that probably contains nontabular data
##'
##' @param path Path to the xlsx file to load.  xls files are not supported.
##'
##' @param sheet Sheet number (not name at this point).  Googlesheets
##'   exported sheets are likely not to do the right thing.
##'
##' @return An \code{xlsx} object that can be printed.  Future methods
##'   might do something sensible.  The structure is subject to
##'   complete change so is not documented here.
##' @export
rexcel_read <- function(path, sheet=1L) {
  rexcel_read_workbook(path, sheet, FALSE)$sheets[[1L]]
}

##' Read an entire workbook
##'
##' @title Read an Excel workbook
##'
##' @param path Path to the xlsx file to load.  xls files are not supported.
##'
##' @param sheets Character or integer vector of sheets to read, or
##'   \code{NULL} to read all sheets (the default)
##'
##' @param progress Display a progress bar?
##' @export
rexcel_read_workbook <- function(path, sheets=NULL, progress=TRUE) {
  if (!file.exists(path)) {
    stop(sprintf("%s does not exist", path))
  }

  dat <- xlsx_read_workbook(path)

  if (is.null(sheets)) {
    sheets <- xlsx_sheet_names(dat)
  } else if (is.numeric(sheets)) {
    sheets <- xlsx_sheet_names(dat)[sheets]
  }

  p <- progress(paste0(basename(path), " [:bar] :current / :total"),
                length(sheets), show=progress)
  p(0)

  ## Some of this will move into the worksheet and save some of the
  ## ugly options passinh here.
  strings <- xlsx_read_shared_strings(path)
  date_offset <- xlsx_date_offset(path)

  style_xlsx <- xlsx_read_style(path)
  lookup <- tibble::data_frame(
    font    = style_xlsx$cell_xfs$font_id,
    fill    = style_xlsx$cell_xfs$fill_id,
    border  = style_xlsx$cell_xfs$border_id,
    num_fmt = style_xlsx$cell_xfs$num_fmt_id)

  ## This becomes read_number_formats?
  if (nrow(style_xlsx$num_fmts) > 0L) {
    n <- max(style_xlsx$num_fmts$num_format_id)
    fmt <- rep(NA_character_, n)
    fmt[seq_along(xlsx_format_codes())] <- xlsx_format_codes()
    fmt[style_xlsx$num_fmts$num_format_id] <- style_xlsx$num_fmts$format_code
  } else {
    fmt <- xlsx_format_codes()
  }
  num_fmt <- tibble::data_frame(num_fmt = fmt)
  style <- linen::linen_style(lookup, font = style_xlsx$fonts,
                              fill = style_xlsx$fills,
                              border = style_xlsx$borders,
                              num_fmt = num_fmt)

  workbook <- linen::workbook(sheets, style, dat$defined_names)
  for (s in sheets) {
    p(1)
    rexcel_read_worksheet(path, s, workbook, dat, strings, style, date_offset)
  }

  workbook
}

## The name here is a bit of a gong show, as is the general logic.  I
## hope this will refine a bit over the next little bit.
rexcel_read_worksheet <- function(path, sheet, workbook,
                                  workbook_dat, strings, style, date_offset) {
  if (is.numeric(sheet)) {
    sheet_idx <- sheet
    sheet_name <- workbook$names[[sheet]]
  } else if (is.character(sheet)) {
    sheet_idx <- match(sheet, workbook$names)
    sheet_name <- sheet
  } else {
    stop("Invalid input for sheet")
  }

  target <- xlsx_internal_sheet_name(sheet, workbook_dat)
  rels <- xlsx_read_rels(path, target)

  xml <- xlsx_read_sheet(path, sheet_idx, workbook_dat)
  ns <- xml2::xml_ns(xml)

  merged <- xlsx_read_merged(xml, ns)
  view <- xlsx_ct_worksheet_views(xml, ns)
  cols <- xlsx_ct_cols(xml, ns) # NOTE: not used yet
  dat <- xlsx_parse_cells(xml, ns, strings, style, date_offset)
  rows <- dat$rows
  cells <- linen::cells(dat$cells$ref, dat$cells$style, dat$cells$type,
                        dat$cells$value, dat$cells$formula)

  comments <- NULL
  if (!is.null(rels)) {
    path_comments <- rels$target_abs[rels$type == "comments"]
    if (length(path_comments) == 1L) {
      comments <- xlsx_read_comments(path, path_comments)
    } else if (length(path_comments) > 1L) {
      stop("CHECK THIS") # TODO: assertion.
    }
  }

  linen::worksheet(sheet_name, cols, rows, cells, merged, view, comments,
                   workbook)
}

xlsx_read_sheet <- function(path, sheet, workbook_dat) {
  xml <- xlsx_read_file(path, xlsx_internal_sheet_name(sheet, workbook_dat))
  stopifnot(xml2::xml_name(xml) == "worksheet")
  xml
}


#' Read XML for a specific file
#'
#' Read in the XML for a specific file within the xlsx, e.g. the file
#' corresponding to a specific worksheet.
#'
#' @param path path to xlsx
#' @param file xml file corresponding to a specific worksheet
#'
#' @return an XML document
#'
#' @keywords internal
xlsx_read_file <- function(path, file) {
  tmp <- tempfile()
  dir.create(tmp)
  ## Oh boy more terrible default behaviour.
  filename <- tryCatch(utils::unzip(path, file, exdir = tmp),
                       warning = function(e) stop(e))
  on.exit(unlink(tmp, recursive = TRUE))
  xml2::read_xml(filename)
}

xlsx_read_file_if_exists <- function(path, file, missing=NULL) {
  ## TODO: Appropriate error handling here is difficult; we should
  ## check that `path` exists, but by the time that this is called we
  ## know that already.
  tmp <- tempfile()
  dir.create(tmp)
  filename <- tryCatch(utils::unzip(path, file, exdir=tmp),
                       warning=function(e) NULL,
                       error=function(e) NULL)
  if (is.null(filename)) {
    missing
  } else {
    on.exit(unlink(tmp, recursive=TRUE))
    xml2::read_xml(filename)
  }
}

## sheetData: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx
##
##   Nothing looks interesting in sheetData, and all elements must be
##   'row'.
##
## row: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.row.aspx
##   The only interesting attribute here is "hidden", but that
##   includes being collapsed by outline, so probably not that
##   interesting.
##
## cell: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.cell.aspx
##
##   Might contain
##     <f>: formula
##     <is> rich test inline
##     <v> value
##   Interesting attributes:
##     r: an A1 style reference to the locatiopn of this cell
##     s: the index of this cell's style (if colours are used as a guide)
##     t: type "an enumeration representing the cell's data type", the
##       only reference to which I can find is
##       http://mailman.vse.cz/pipermail/sc34wg4/attachments/20100428/3fc0a446/attachment-0001.pdf
##       - b: boolean
##       - d: date (ISO 8601)
##       -  e: error
##       - inlineStr: inline string in rich text format, with
##           contents in the 'is' element, not the 'v' element.
##       - n: number
##       - s: shared string
##       - str: formula string
##
## However, many numbers seem not to have a "t" attribute which is
## charming.
##
## NOTE: handling of formulae is potentially tricky as they can have an attribute "shared" which
##
## Blank cells have no children at all.
##
## See readxl/src/XlsxCell.h: XlsxCell::type()
xlsx_parse_cells <- function(xml, ns, strings, style_data, date_offset) {
  sheet_data <- xlsx_read_sheet_data(xml, ns, strings)
  cells <- sheet_data$cells
  rows <- sheet_data$rows

  ## TODO: Roll this back into the xfs parsing perhaps?  in the (not
  ## yet existing) compute style part I think.  We can have an
  ## "is_date" entry there.
  custom_date <- style_data$num_fmts$num_format_id[
    grepl("[dmyhs]", style_data$num_fmts$format_code)]
  is_date_time <- xlsx_is_date_time(style_data$cell_xfs$num_fmt_id, custom_date)

  type <- character(nrow(cells))
  type[!is.na(cells$type) & cells$type == "b"] <- "bool"
  type[!is.na(cells$type) & cells$type == "s" | cells$type == "str"] <- "text"
  i <- is.na(cells$type) | cells$type == "n"
  j <- is_date_time[cells$style[i]]
  type[i] <- ifelse(!is.na(j) & j, "date", "number")
  type[lengths(cells$value) == 0L] <- "blank"
  cells$type <- type

  i <- type == "date"
  cells$value[i] <-
    as.list(as.POSIXct(unlist(cells$value[i]) * 86400, "UTC", date_offset))

  list(cells=cells, rows=rows)
}

xlsx_sheet_names <- function(dat) {
  if (is.character(dat)) {
    dat <- xlsx_read_workbook(dat)
  }
  sheets <- dat$sheets
  sheets$name[sheets$type == "worksheet" & sheets$state != "veryHidden"]
}

## Return the filename within the bundle
xlsx_internal_sheet_name <- function(sheet, dat) {
  if (length(sheet) != 1L) {
    stop("'sheet' must be a scalar")
  }
  if (is.na(sheet)) {
    stop("'sheet' must be non-missing")
  }

  sheets <- dat$sheets
  sheets <- sheets[sheets$type == "worksheet" & sheets$state != "veryHidden", ]

  if (is.character(sheet)) {
    target <- sheets$target_abs[match(sheet, sheets$name)]
  } else if (is.numeric(sheet)) {
    target <- sheets$target_abs[[sheet]]
  } else {
    stop("invalid input")
  }
  target
}

## NOTE: Date handling will change a bit once I get the string parsing
## stuff entirely worked out.
xlsx_date_offset <- function(path) {
  ## See readxl/src/utils.h: dateOffset
  ## See readxl/src/XlsxWorkbook.h: is1904
  xml <- xlsx_read_file(path, "xl/workbook.xml")
  date1904 <- xml2::xml_find_one(xml, "/d1:workbook/d1:workbookPr/@date1904",
                                 xml2::xml_ns(xml))
  if (inherits(date1904, "xml_missing")) {
    date_is_1904 <- FALSE
  } else {
    ## TODO: in theory we should do whatever atoi would allow here
    ## (that's what Hadley uses in the C++) but I have a sheet that
    ## contains this as "false".  So I'm trying this way for now.
    value <- xml2::xml_text(date1904)
    date_is_1904 <- value == "1" || value == "true"
  }
  if (date_is_1904) "1904-01-01" else "1899-12-30"
}

xlsx_is_date_time <- function(id, custom) {
  ## See readxl's src/CellType.h: isDateTime()
  id %in% c(c(14:22, 27:36, 45:47, 50:58, 71:81), custom)
}
