## ------------------------------------------------------------------------
library(rprojroot)
devtools::load_all(find_package_root_file())

## ------------------------------------------------------------------------
mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
rexcel_workbook(mini_gap_path)

## ------------------------------------------------------------------------
ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
                       package = "rexcel")
rexcel_workbook(ff_path)

## ------------------------------------------------------------------------
(ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
                        package = "rexcel"))

## enter rexcel_read_workbook()
path <- ff_path
sheets <- 1L

## this gets info about the files inside the zip archive
dat <- xlsx_read_workbook(path)
dat$rels   ## ?files in the zip archive?
dat$sheets ## ?files corresponding to worksheets?
(sheets <- xlsx_sheet_names(dat)[sheets])

(strings <- xlsx_read_shared_strings(path))
(date_offset <- xlsx_date_offset(path))

style_xlsx <- xlsx_read_style(path)
str(style_xlsx, max.level = 1)
(lookup <- tibble::data_frame(
  font    = style_xlsx$cell_xfs$font_id,
  fill    = style_xlsx$cell_xfs$fill_id,
  border  = style_xlsx$cell_xfs$border_id,
  num_fmt = style_xlsx$cell_xfs$num_fmt_id))

## numeric formatting
n <- max(style_xlsx$num_fmts$num_format_id)
fmt <- rep(NA_character_, n)
fmt[seq_along(xlsx_format_codes())] <- xlsx_format_codes()
fmt[style_xlsx$num_fmts$num_format_id] <- style_xlsx$num_fmts$format_code
num_fmt <- tibble::data_frame(num_fmt = fmt)
style <- linen::linen_style(lookup, font = style_xlsx$fonts,
                            fill = style_xlsx$fills,
                            border = style_xlsx$borders,
                            num_fmt = num_fmt)

(workbook <- linen::workbook(sheets, style, dat$defined_names))

## ------------------------------------------------------------------------
## enter rexcel_read_worksheet()
## rexcel_read_worksheet(path, s, workbook, dat, strings, style, date_offset)
(sheet <- sheets[1])
(sheet_idx <- match(sheet, workbook$names))
(sheet_name <- sheet)

(target <- xlsx_internal_sheet_name(sheet, dat))
(rels <- xlsx_read_rels(path, target))

## ------------------------------------------------------------------------
## enter xlsx_read_sheet()
(file <- xlsx_internal_sheet_name(sheet_idx, dat))
xml <- xlsx_read_file(path, file) ## at last! the xml! w00t!
(ns <- xml2::xml_ns(xml)) ## much less w00t now :(

(merged <- xlsx_read_merged(xml, ns))
(view <- xlsx_ct_worksheet_views(xml, ns))
(cols <- xlsx_ct_cols(xml, ns)) # NOTE: not used yet

## this is where it's at!
(cell_dat <- xlsx_parse_cells(xml, ns, strings, style, date_offset))

## not even sure what this is
(rows <- cell_dat$rows)

## this is where cells come from  
(cells <- linen::cells(cell_dat$cells$ref, cell_dat$cells$style,
                       cell_dat$cells$type, cell_dat$cells$value,
                       cell_dat$cells$formula))

## in real life and in other sheets, it's possible comments will be populated
## but not in this sheet
comments <- NULL

## ------------------------------------------------------------------------
(ws <- linen::worksheet(sheet_name, cols, rows, cells, merged, view, comments,
                        workbook))

