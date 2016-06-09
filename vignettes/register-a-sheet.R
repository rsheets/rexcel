## ----setup, include = FALSE, cache = FALSE-------------------------------
knitr::opts_chunk$set(error = TRUE)  

## ------------------------------------------------------------------------
library(rprojroot)
devtools::load_all(find_package_root_file())

## ------------------------------------------------------------------------
mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
                       package = "rexcel")
ek_path <- system.file("sheets", "Ekaterinburg_IP_9.xlsx", package = "rexcel")
ek2_path <- system.file("sheets", "Ekaterinburg_IP_9-RESAVED.xlsx",
                        package = "rexcel")
dn_path <- system.file("sheets", "defined-names.xlsx", package = "rexcel")
gabe_path <- system.file("sheets", "gabe.xlsx", package = "rexcel")

## ------------------------------------------------------------------------
is_xlsx(mini_gap_path)

## ------------------------------------------------------------------------
manifest <- xlsx_list_files(mini_gap_path)
print(manifest, n = Inf)

## ------------------------------------------------------------------------
print(xlsx_list_files(ek_path), n = Inf)

## ------------------------------------------------------------------------
print(xlsx_list_files(ek2_path), n = Inf)

## ------------------------------------------------------------------------
(ct <- as.data.frame(xlsx_read_Content_Types(mini_gap_path)))
#setdiff(manifest$name, gsub("^\\/", "", ct$part_name))
#intersect(gsub("^\\/", "", ct$part_name), manifest$name)

## ------------------------------------------------------------------------
(sheets <- xlsx_read_workbook_sheets(mini_gap_path))

## ------------------------------------------------------------------------
xlsx_read_workbook_sheets(ff_path)
xlsx_read_workbook_sheets(ek_path)
xlsx_read_workbook_sheets(ek2_path)
xlsx_read_workbook_sheets(dn_path)
xlsx_read_workbook_sheets(gabe_path)

## ------------------------------------------------------------------------
xlsx_read_workbook_defined_names(mini_gap_path)
xlsx_read_workbook_defined_names(dn_path)
xlsx_read_workbook_defined_names(gabe_path)

## ------------------------------------------------------------------------
(workbook_rels <- xlsx_read_workbook_rels(mini_gap_path))
xlsx_read_workbook_rels(ek_path)
xlsx_read_workbook_rels(ek2_path)

## ------------------------------------------------------------------------
(sheets_df <- join_sheets_workbook_rels(sheets, workbook_rels))

## ------------------------------------------------------------------------
(shared_strings <- xlsx_read_shared_strings(mini_gap_path))

## ------------------------------------------------------------------------
xlsx_read_shared_strings(ff_path)
str(xlsx_read_shared_strings(ek_path))
str(xlsx_read_shared_strings(ek2_path))
xlsx_read_shared_strings(gabe_path)

## ------------------------------------------------------------------------
(styles <- xlsx_read_style(mini_gap_path))

## ------------------------------------------------------------------------
xlsx_read_style(ff_path)
## as explained above, namespace crazy means this won't work
#xlsx_read_style(ek_path)
xlsx_read_style(ek2_path)
xlsx_read_style(dn_path)
xlsx_read_style(gabe_path)

## ------------------------------------------------------------------------
as.data.frame(worksheet_rels <- xlsx_read_worksheet_rels(mini_gap_path))
as.data.frame(xlsx_read_worksheet_rels(ff_path))
as.data.frame(xlsx_read_worksheet_rels(ek_path))
as.data.frame(xlsx_read_worksheet_rels(ek2_path))
as.data.frame(xlsx_read_worksheet_rels(dn_path))
as.data.frame(xlsx_read_worksheet_rels(gabe_path))

## ------------------------------------------------------------------------
subset(xlsx_list_files(mini_gap_path), grepl("drawing", name))

## ------------------------------------------------------------------------
mini_gap_workbook <- rexcel_register(mini_gap_path)
str(mini_gap_workbook, max.level = 1)
mini_gap_workbook

## ------------------------------------------------------------------------
path <- mini_gap_path

dat <- xlsx_read_workbook(path)
str(dat, max.level = 1)
dat$rels
dat$sheets
dat$defined_names

## let's look at defined names in a sheet that actually has them
xlsx_read_workbook(dn_path)$defined_names

## ------------------------------------------------------------------------
(strings <- xlsx_read_shared_strings(path))
(date_offset <- xlsx_date_offset(path))

## ------------------------------------------------------------------------
style_xlsx <- xlsx_read_style(path)
str(style_xlsx, max.level = 1)
(lookup <- tibble::data_frame(
  font    = style_xlsx$cell_xfs$font_id,
  fill    = style_xlsx$cell_xfs$fill_id,
  border  = style_xlsx$cell_xfs$border_id,
  num_fmt = style_xlsx$cell_xfs$num_fmt_id))

## ------------------------------------------------------------------------
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

sheets <- 1L
(sheets <- xlsx_sheet_names(dat)[sheets])

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

