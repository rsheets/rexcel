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
  ## NOTE: Some docs here:
  ##   https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.aspx
  ## though getting the actual spec would be nicer I suspect.
  if (!file.exists(path)) {
    stop(sprintf("%s does not exist", path))
  }
  xml <- xlsx_read_sheet(path, sheet)
  ns <- xml2::xml_ns(xml)
  strings <- xlsx_read_shared_strings(path)
  style <- xlsx_read_style(path)

  ## According to the spec mergeCells contains only mergeCell
  ## elements, and they contain only a "ref" attribute.  Once I track
  ## down the full schema (MS's website is a mess here) we can add
  ## correct references for this assertion.
  merged <- xlsx_read_merged(xls, ns)

  date_offset <- xlsx_date_offset(path)

  ## For the vast majority of sheets, this should be the longest step.
  ## The per-cell processing is pretty hard on the XML processing in
  ## R.
  cells <- xlsx_parse_cells(xml, ns, strings, style, date_offset)

  linen::worksheet(cells, merged, linen::workbook(style))
}

## Non api functions:
## xlsx_read_*: reads something from a file
## xlsx_parse_*: turns xml into somethig usable

xlsx_read_sheet <- function(path, sheet) {
  xml <- xlsx_read_file(path, xlsx_internal_sheet_name(path, sheet))
  stopifnot(xml2::xml_name(xml) == "worksheet")
  xml
}

xlsx_read_file <- function(path, file) {
  tmp <- tempfile()
  dir.create(tmp)
  ## Oh boy more terrible default behaviour.
  filename <- tryCatch(utils::unzip(path, file, exdir=tmp),
                       warning=function(e) stop(e))
  on.exit(unlink(tmp, recursive=TRUE))
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

## If the format is <si>/<t> then we can just take the text values.
## Otherwise we'll have to parse out the RTF strings separately.
xlsx_read_shared_strings <- function(path) {
  xml <- xlsx_read_file_if_exists(path, "xl/sharedStrings.xml")
  if (is.null(xml)) {
    return(character(0))
  }
  ns <- xml2::xml_ns(xml)
  si <- xml2::xml_find_all(xml, "d1:si", ns)
  if (length(si) == 0L) {
    return(character(0))
  }
  ## TODO: This is a bug in xls
  is_rich <- xml2::xml_find_lgl(si, "boolean(d1:r)", ns)
  ret <- character(length(si))
  ret[!is_rich] <-
    xml2::xml_text(xml2::xml_find_one(si[!is_rich], "d1:t", ns))
  ret[is_rich] <- vcapply(si[is_rich], xlsx_parse_string, ns)
  ret
}

xlsx_read_merged <- function(xml, ns) {
  merged <- xml2::xml_text(
    xml2::xml_find_all(xml, "./d1:mergeCells/d1:mergeCell/@ref", ns))
  merged <- lapply(merged, cellranger::as.cell_limits)
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
  sheet_data <- xml2::xml_find_one(xml, "d1:sheetData", ns)
  cells <- xml2::xml_find_all(sheet_data, "./d1:row/d1:c", ns)

  xml_find_if_exists <- function(x, xpath, ns) {
    i <- xml2::xml_find_lgl(x, sprintf("boolean(%s)", xpath), ns)
    ret <- vector("list", length(i))
    ret[i] <- xml2::xml_text(xml2::xml_find_one(x[i], xpath, ns))
    ret
  }

  ref <- xml2::xml_attr(cells, "r")
  style <- as.integer(xml2::xml_attr(cells, "s"))
  cells_type <- xml2::xml_attr(cells, "t")

  formula <- xml_find_if_exists(cells, "./d1:f", ns)
  value <- xml_find_if_exists(cells, "./d1:v", ns)

  ## Quick check to make sure we didn't miss anything (I think it's
  ## only is values)
  inline_string <- xml2::xml_find_lgl(cells, "boolean(./d1:is)", ns)
  if (any(inline_string)) {
    ## These would get fired through the string parsing I think.
    stop("Inline string value not yet handled")
  }

  ## TODO: Roll this back into the xfs parsing perhaps?
  if ("formatCode" %in% names(style_data$num_formats)) {
    custom_date <- style_data$num_formats$numFmtId[
      grepl("[dmyhs]", style_data$num_formats$formatCode)]
  } else {
    custom_date <- integer()
  }

  ## Might roll this back into the style?
  is_date_time <- xlsx_is_date_time(style_data$cell_xfs$numFmtId, custom_date)

  type <- character(length(value))
  type[!is.na(cells_type) & cells_type == "b"] <- "bool"
  type[!is.na(cells_type) & cells_type == "s" | cells_type == "str"] <- "text"
  i <- is.na(cells_type) | cells_type == "n"
  j <- is_date_time[style[i] + 1L]
  type[i] <- ifelse(!is.na(j) & j, "date", "number")
  type[lengths(value) == 0L] <- "blank"

  ## String substitutions:
  i <- which(cells_type == "s")
  value[i] <- strings[as.integer(unlist(value[i])) + 1L]

  i <- type == "bool" | type == "number"
  value[i] <- as.numeric(unlist(value[i]))

  i <- type == "date"
  value[i] <-
    as.list(as.POSIXct(as.numeric(unlist(value[i])) * 86400,
                       "UTC", date_offset))

  linen::cells(ref, style, value, formula, type)
}

xlsx_parse_string <- function(x, ns) {
  t <- xml2::xml_find_one(x, "d1:t", ns)
  if (inherits(t, "xml_missing")) {
    ## NOTE: it *looks* like most of the time we can do xml_text(x)
    ## here and get about the right answer.
    str <- character()
  } else {
    str <- xml2::xml_text(t)
  }
  r <- xml2::xml_find_all(x, "d1:r", ns)
  if (length(r) > 0L) {
    str <- paste(c(str, xml2::xml_text(r)), collapse="")
  }
  ## TODO: still need to "unescape" these.
  str
}

xlsx_sheet_names <- function(filename) {
  xml <- xlsx_read_file(filename, "xl/workbook.xml")
  ns <- xml2::xml_ns(xml)
  xml2::xml_text(xml2::xml_find_all(xml, "d1:sheets/d1:sheet/@name", ns))
}

## Return the filename within the bundle
xlsx_internal_sheet_name <- function(filename, sheet) {
  if (length(sheet) != 1L) {
    stop("'sheet' must be a scalar")
  }
  if (is.na(sheet)) {
    stop("'sheet' must be non-missing")
  }
  if (is.character(sheet)) {
    sheet <- match(sheet, xlsx_sheet_names(filename))
  } else if (!(is.integer(sheet) || is.numeric(sheet))) {
    stop("'sheet' must be an integer or a string")
  }

  ## TODO: Looks like this does always exist.
  rels <- xlsx_read_file_if_exists(filename, "xl/_rels/workbook.xml.rels")
  if (is.null(rels)) {
    target <- sprintf("xl/worksheets/sheet%d.xml", sheet)
  } else {
    ## This might fail with a cryptic error if my assumptions are
    ## incorrect.
    xml <- xlsx_read_file(filename, "xl/workbook.xml")
    xpath <- sprintf("d1:sheets/d1:sheet[%d]", sheet)
    node <- xml2::xml_find_one(xml, xpath, xml2::xml_ns(xml))
    id <- xml2::xml_attr(node, "id")
    ## This _should_ work but I don't see it:
    ##   xpath <- sprintf("string(d1:sheets/d1:sheet[%d]/@id)", sheet)
    ##   xml2::xml_find_chr(xml, xpath, ns) # --> ""
    xpath <- sprintf('/d1:Relationships/d1:Relationship[@Id = "%s"]/@Target',
                     id)
    target <- xml2::xml_text(xml2::xml_find_one(rels, xpath,
                                                xml2::xml_ns(rels)))
    ## NOTE: these are _relative_ paths so must be qualified here:
    target <- file.path("xl", target)
  }
  target
}

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
