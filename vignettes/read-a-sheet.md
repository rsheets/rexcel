# Read an Excel Sheet
Jennifer Bryan  
`r Sys.Date()`  


```r
library(rprojroot)
```

```
## Warning: package 'rprojroot' was built under R version 3.2.4
```

```r
devtools::load_all(find_package_root_file())
```

```
## Loading rexcel
```

*Using a function I wrote while exploring all the files that make up an xlsx.*

Apply it to mini gapminder.


```r
mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
rexcel_workbook(mini_gap_path)
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/mini-gap.xlsx"
## 
## $reg_time
## [1] "2016-05-20 15:18:34 PDT"
## 
## $manifest
## Source: local data frame [21 x 3]
## 
##                                   Name Length                Date
##                                  <chr>  <dbl>              <time>
## 1             xl/worksheets/sheet1.xml   2136 2015-04-25 12:00:00
## 2  xl/worksheets/_rels/sheet1.xml.rels    307 2015-04-25 12:00:00
## 3             xl/worksheets/sheet2.xml   2136 2015-04-25 12:00:00
## 4  xl/worksheets/_rels/sheet2.xml.rels    307 2015-04-25 12:00:00
## 5             xl/worksheets/sheet3.xml   2146 2015-04-25 12:00:00
## 6  xl/worksheets/_rels/sheet3.xml.rels    307 2015-04-25 12:00:00
## 7             xl/worksheets/sheet4.xml   2136 2015-04-25 12:00:00
## 8  xl/worksheets/_rels/sheet4.xml.rels    307 2015-04-25 12:00:00
## 9             xl/worksheets/sheet5.xml   2144 2015-04-25 12:00:00
## 10 xl/worksheets/_rels/sheet5.xml.rels    307 2015-04-25 12:00:00
## ..                                 ...    ...                 ...
## 
## $content_types
## Source: local data frame [15 x 3]
## 
##                              PartName Extension
##                                 <chr>     <chr>
## 1                                <NA>      rels
## 2                                <NA>       xml
## 3  /xl/drawings/worksheetdrawing4.xml      <NA>
## 4  /xl/drawings/worksheetdrawing2.xml      <NA>
## 5  /xl/drawings/worksheetdrawing1.xml      <NA>
## 6  /xl/drawings/worksheetdrawing3.xml      <NA>
## 7  /xl/drawings/worksheetdrawing5.xml      <NA>
## 8                      /xl/styles.xml      <NA>
## 9               /xl/sharedStrings.xml      <NA>
## 10                   /xl/workbook.xml      <NA>
## 11          /xl/worksheets/sheet5.xml      <NA>
## 12          /xl/worksheets/sheet3.xml      <NA>
## 13          /xl/worksheets/sheet1.xml      <NA>
## 14          /xl/worksheets/sheet4.xml      <NA>
## 15          /xl/worksheets/sheet2.xml      <NA>
## Variables not shown: ContentType <chr>.
## 
## $sheets
## Source: local data frame [5 x 4]
## 
##     state     name sheetId    id
##     <chr>    <chr>   <int> <chr>
## 1 visible   Africa       1  rId3
## 2 visible Americas       2  rId4
## 3 visible     Asia       3  rId5
## 4 visible   Europe       4  rId6
## 5 visible  Oceania       5  rId7
## 
## $sheets_df
## Source: local data frame [5 x 5]
## 
##   sheetId     name    Id                   Target
##     <int>    <chr> <chr>                    <chr>
## 1       1   Africa  rId3 xl/worksheets/sheet4.xml
## 2       2 Americas  rId4 xl/worksheets/sheet3.xml
## 3       3     Asia  rId5 xl/worksheets/sheet5.xml
## 4       4   Europe  rId6 xl/worksheets/sheet1.xml
## 5       5  Oceania  rId7 xl/worksheets/sheet2.xml
## Variables not shown: Type <chr>.
## 
## $shared_strings
##  [1] "country"                "continent"             
##  [3] "year"                   "lifeExp"               
##  [5] "pop"                    "gdpPercap"             
##  [7] "Algeria"                "Africa"                
##  [9] "Angola"                 "Albania"               
## [11] "Europe"                 "Benin"                 
## [13] "Austria"                "Argentina"             
## [15] "Americas"               "Belgium"               
## [17] "Australia"              "Oceania"               
## [19] "Bolivia"                "Bosnia and Herzegovina"
## [21] "New Zealand"            "Bulgaria"              
## [23] "Brazil"                 "Canada"                
## [25] "Afghanistan"            "Asia"                  
## [27] "Bahrain"                "Chile"                 
## [29] "Bangladesh"             "Botswana"              
## [31] "Cambodia"               "China"                 
## [33] "Burkina Faso"          
## attr(,"count")
## [1] 80
## attr(,"uniqueCount")
## [1] 33
## 
## $styles
## $styles$fonts
## Source: local data frame [1 x 3]
## 
##      sz    color  name
##   <chr>    <chr> <chr>
## 1  10.0 FF000000 Arial
## 
## $styles$fills
## NULL
## 
## $styles$borders
## NULL
## 
## $styles$cell_style_xfs
## NULL
## 
## $styles$cell_xfs
## NULL
## 
## $styles$cell_styles
## NULL
## 
## $styles$num_fmts
## NULL
## 
## $styles$dxfs
## NULL
## 
## 
## $workbook_rels
## Source: local data frame [7 x 3]
## 
##      Id                Target
##   <chr>                 <chr>
## 1  rId2     sharedStrings.xml
## 2  rId1            styles.xml
## 3  rId4 worksheets/sheet3.xml
## 4  rId3 worksheets/sheet4.xml
## 5  rId6 worksheets/sheet1.xml
## 6  rId5 worksheets/sheet5.xml
## 7  rId7 worksheets/sheet2.xml
## Variables not shown: Type <chr>.
## 
## $worksheet_rels
## $worksheet_rels[[1]]
## {xml_nodeset (1)}
## [1] <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/offi ...
## 
## $worksheet_rels[[2]]
## {xml_nodeset (1)}
## [1] <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/offi ...
## 
## $worksheet_rels[[3]]
## {xml_nodeset (1)}
## [1] <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/offi ...
## 
## $worksheet_rels[[4]]
## {xml_nodeset (1)}
## [1] <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/offi ...
## 
## $worksheet_rels[[5]]
## {xml_nodeset (1)}
## [1] <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/offi ...
```

Apply it to formula and formatting sheet.


```r
ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
                       package = "rexcel")
rexcel_workbook(ff_path)
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/gs-test-formula-formatting.xlsx"
## 
## $reg_time
## [1] "2016-05-20 15:18:34 PDT"
## 
## $manifest
## Source: local data frame [9 x 3]
## 
##                                  Name Length                Date
##                                 <chr>  <dbl>              <time>
## 1            xl/worksheets/sheet1.xml  60580 2016-05-03 02:56:00
## 2 xl/worksheets/_rels/sheet1.xml.rels    471 2016-05-03 02:56:00
## 3   xl/drawings/worksheetdrawing1.xml    494 2016-05-03 02:56:00
## 4                xl/sharedStrings.xml    407 2016-05-03 02:56:00
## 5                       xl/styles.xml   3014 2016-05-03 02:56:00
## 6                     xl/workbook.xml    731 2016-05-03 02:56:00
## 7          xl/_rels/workbook.xml.rels    565 2016-05-03 02:56:00
## 8                         _rels/.rels    296 2016-05-03 02:56:00
## 9                 [Content_Types].xml    945 2016-05-03 02:56:00
## 
## $content_types
## Source: local data frame [7 x 3]
## 
##                             PartName Extension
##                                <chr>     <chr>
## 1                               <NA>       xml
## 2                               <NA>      rels
## 3          /xl/worksheets/sheet1.xml      <NA>
## 4              /xl/sharedStrings.xml      <NA>
## 5 /xl/drawings/worksheetdrawing1.xml      <NA>
## 6                     /xl/styles.xml      <NA>
## 7                   /xl/workbook.xml      <NA>
## Variables not shown: ContentType <chr>.
## 
## $sheets
## Source: local data frame [1 x 4]
## 
##     state   name sheetId    id
##     <chr>  <chr>   <int> <chr>
## 1 visible Sheet1       1  rId3
## 
## $sheets_df
## Source: local data frame [1 x 5]
## 
##   sheetId   name    Id                   Target
##     <int>  <chr> <chr>                    <chr>
## 1       1 Sheet1  rId3 xl/worksheets/sheet1.xml
## Variables not shown: Type <chr>.
## 
## $shared_strings
##  [1] "integer"           "number_formatted"  "number_rounded"   
##  [4] "character"         "formula"           "formula_formatted"
##  [7] "one"               "three"             "four"             
## [10] "five"             
## attr(,"count")
## [1] 10
## attr(,"uniqueCount")
## [1] 10
## 
## $styles
## $styles$fonts
## Source: local data frame [3 x 3]
## 
##      sz    color        name
##   <chr>    <chr>       <chr>
## 1  10.0 FF000000       Arial
## 2  <NA> FF0000FF        <NA>
## 3  <NA>     <NA> Courier New
## 
## $styles$fills
## NULL
## 
## $styles$borders
## NULL
## 
## $styles$cell_style_xfs
## NULL
## 
## $styles$cell_xfs
## NULL
## 
## $styles$cell_styles
## NULL
## 
## $styles$num_fmts
## NULL
## 
## $styles$dxfs
## NULL
## 
## 
## $workbook_rels
## Source: local data frame [3 x 3]
## 
##      Id                Target
##   <chr>                 <chr>
## 1  rId1            styles.xml
## 2  rId2     sharedStrings.xml
## 3  rId3 worksheets/sheet1.xml
## Variables not shown: Type <chr>.
## 
## $worksheet_rels
## $worksheet_rels[[1]]
## {xml_nodeset (2)}
## [1] <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/offi ...
## [2] <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/offi ...
```

*Here I'm exploring the existing sheet reading functionality, using the vignette format. This is not an actual vignette!*

Peeling the many-layered onion that is `rexcel_read()` until I get at the XML for a worksheet. Wish me luck.

We'll work with an example sheet created for `googlesheets` that has alot of formulas and formatting going.

Objective 1: create a `linen::workbook` object. Dropping into code inside `rexcel_read_workbook()`.


```r
(ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
                        package = "rexcel"))
```

```
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/gs-test-formula-formatting.xlsx"
```

```r
## enter rexcel_read_workbook()
path <- ff_path
sheets <- 1L

## this gets info about the files inside the zip archive
dat <- xlsx_read_workbook(path)
dat$rels   ## ?files in the zip archive?
```

```
## Source: local data frame [3 x 4]
## 
##      id          type                target               target_abs
##   <chr>         <chr>                 <chr>                    <chr>
## 1  rId1        styles            styles.xml            xl/styles.xml
## 2  rId2 sharedStrings     sharedStrings.xml     xl/sharedStrings.xml
## 3  rId3     worksheet worksheets/sheet1.xml xl/worksheets/sheet1.xml
```

```r
dat$sheets ## ?files corresponding to worksheets?
```

```
##     name sheet_id   state  ref      type                target
## 1 Sheet1        1 visible rId3 worksheet worksheets/sheet1.xml
##                 target_abs
## 1 xl/worksheets/sheet1.xml
```

```r
(sheets <- xlsx_sheet_names(dat)[sheets])
```

```
## [1] "Sheet1"
```

```r
(strings <- xlsx_read_shared_strings(path))
```

```
##  [1] "integer"           "number_formatted"  "number_rounded"   
##  [4] "character"         "formula"           "formula_formatted"
##  [7] "one"               "three"             "four"             
## [10] "five"
```

```r
(date_offset <- xlsx_date_offset(path))
```

```
## [1] "1899-12-30"
```

```r
style_xlsx <- xlsx_read_style(path)
str(style_xlsx, max.level = 1)
```

```
## List of 7
##  $ fonts         :Classes 'tbl_df', 'tbl' and 'data.frame':	4 obs. of  13 variables:
##  $ fills         :Classes 'tbl_df', 'tbl' and 'data.frame':	2 obs. of  4 variables:
##  $ borders       :Classes 'tbl_df', 'tbl' and 'data.frame':	1 obs. of  19 variables:
##  $ cell_style_xfs:Classes 'tbl_df', 'tbl' and 'data.frame':	1 obs. of  16 variables:
##  $ cell_xfs      :Classes 'tbl_df', 'tbl' and 'data.frame':	16 obs. of  16 variables:
##  $ cell_styles   :Classes 'tbl_df', 'tbl' and 'data.frame':	1 obs. of  6 variables:
##  $ num_fmts      :Classes 'tbl_df', 'tbl' and 'data.frame':	1 obs. of  2 variables:
```

```r
(lookup <- tibble::data_frame(
  font    = style_xlsx$cell_xfs$font_id,
  fill    = style_xlsx$cell_xfs$fill_id,
  border  = style_xlsx$cell_xfs$border_id,
  num_fmt = style_xlsx$cell_xfs$num_fmt_id))
```

```
## Source: local data frame [16 x 4]
## 
##     font  fill border num_fmt
##    <int> <int>  <int>   <int>
## 1      1    NA     NA      NA
## 2      2    NA     NA      NA
## 3      2    NA     NA       4
## 4      2    NA     NA       5
## 5      3    NA     NA      NA
## 6      2    NA     NA      12
## 7      2    NA     NA      11
## 8      2    NA     NA       5
## 9      2    NA     NA      11
## 10     2    NA     NA      12
## 11     4    NA     NA      NA
## 12     2    NA     NA       3
## 13     2    NA     NA      13
## 14     2    NA     NA     165
## 15     2    NA     NA       4
## 16     2    NA     NA       4
```

```r
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
```

```
## <workbook>
##   Public:
##     add_sheet: function (sheet) 
##     clone: function (deep = FALSE) 
##     defined_names: tbl_df, tbl, data.frame
##     initialize: function (names, style, defined_names) 
##     names: Sheet1
##     sheets: list
##     style: linen_style
```

Objective 2: Visit and extract information for all requested worksheets.

In this case, I'm just reading the first and only sheet. This loop appears in `rexcel_read_workbook()` and calls `rexcel_read_worksheet()` for each requested worksheet. This is the loop and function we eventually exit from and this `workbook` object is what's returned.


```r
## enter rexcel_read_worksheet()
## rexcel_read_worksheet(path, s, workbook, dat, strings, style, date_offset)
(sheet <- sheets[1])
```

```
## [1] "Sheet1"
```

```r
(sheet_idx <- match(sheet, workbook$names))
```

```
## [1] 1
```

```r
(sheet_name <- sheet)
```

```
## [1] "Sheet1"
```

```r
(target <- xlsx_internal_sheet_name(sheet, dat))
```

```
## [1] "xl/worksheets/sheet1.xml"
```

```r
(rels <- xlsx_read_rels(path, target))
```

```
## Source: local data frame [2 x 4]
## 
##      id      type                            target
##   <chr>     <chr>                             <chr>
## 1  rId1 hyperlink            http://www.google.com/
## 2  rId2   drawing ../drawings/worksheetdrawing1.xml
## Variables not shown: target_abs <chr>.
```

Now we drop down into a lower-level non-exported function, `xlsx_read_sheet()`.


```r
## enter xlsx_read_sheet()
(file <- xlsx_internal_sheet_name(sheet_idx, dat))
```

```
## [1] "xl/worksheets/sheet1.xml"
```

```r
xml <- xlsx_read_file(path, file) ## at last! the xml! w00t!
(ns <- xml2::xml_ns(xml)) ## much less w00t now :(
```

```
## d1    <-> http://schemas.openxmlformats.org/spreadsheetml/2006/main
## r     <-> http://schemas.openxmlformats.org/officeDocument/2006/relationships
## mx    <-> http://schemas.microsoft.com/office/mac/excel/2008/main
## mc    <-> http://schemas.openxmlformats.org/markup-compatibility/2006
## mv    <-> urn:schemas-microsoft-com:mac:vml
## x14   <-> http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
## x14ac <-> http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac
## xm    <-> http://schemas.microsoft.com/office/excel/2006/main
```

```r
(merged <- xlsx_read_merged(xml, ns))
```

```
## list()
```

```r
(view <- xlsx_ct_worksheet_views(xml, ns))
```

```
## NULL
```

```r
(cols <- xlsx_ct_cols(xml, ns)) # NOTE: not used yet
```

```
## Source: local data frame [6 x 9]
## 
##   best_fit collapsed custom_width hidden   min   max outline_level style
##      <lgl>     <lgl>        <lgl>  <lgl> <int> <int>         <int> <int>
## 1    FALSE     FALSE         TRUE  FALSE     1     1            NA    NA
## 2    FALSE     FALSE         TRUE  FALSE     2     2            NA    NA
## 3    FALSE     FALSE         TRUE  FALSE     3     3            NA    NA
## 4    FALSE     FALSE         TRUE  FALSE     4     4            NA    NA
## 5    FALSE     FALSE         TRUE  FALSE     5     5            NA    NA
## 6    FALSE     FALSE         TRUE  FALSE     6     6            NA    NA
## Variables not shown: width <dbl>.
```

```r
## this is where it's at!
(cell_dat <- xlsx_parse_cells(xml, ns, strings, style, date_offset))
```

```
## $cells
## Source: local data frame [2,022 x 5]
## 
##      ref style   type formula     value
##    <chr> <int>  <chr>   <chr>    <list>
## 1     A1     2   text    <NA> <chr [1]>
## 2     B1     3   text    <NA> <chr [1]>
## 3     C1     4   text    <NA> <chr [1]>
## 4     D1     2   text    <NA> <chr [1]>
## 5     E1     2   text    <NA> <chr [1]>
## 6     F1     2   text    <NA> <chr [1]>
## 7     A2     2 number    <NA> <dbl [1]>
## 8     B2     3 number    <NA> <dbl [1]>
## 9     C2     4 number    <NA> <dbl [1]>
## 10    D2     2   text    <NA> <chr [1]>
## ..   ...   ...    ...     ...       ...
## 
## $rows
## Source: local data frame [1,000 x 11]
## 
##        r spans     s custom_format    ht hidden custom_height
##    <int> <chr> <int>         <lgl> <dbl>  <lgl>         <lgl>
## 1      1  <NA>    NA         FALSE    NA  FALSE            NA
## 2      2  <NA>    NA         FALSE    NA  FALSE            NA
## 3      3  <NA>    NA         FALSE    NA  FALSE            NA
## 4      4  <NA>    NA         FALSE    NA  FALSE            NA
## 5      5  <NA>    NA         FALSE    NA  FALSE            NA
## 6      6  <NA>    NA         FALSE    NA  FALSE            NA
## 7      7  <NA>    NA         FALSE    NA  FALSE            NA
## 8      8  <NA>    NA         FALSE    NA  FALSE            NA
## 9      9  <NA>    NA         FALSE    NA  FALSE            NA
## 10    10  <NA>    NA         FALSE    NA  FALSE            NA
## ..   ...   ...   ...           ...   ...    ...           ...
## Variables not shown: outline_level <int>, collapsed <lgl>, thick_top
##   <lgl>, thick_bot <lgl>.
```

```r
## not even sure what this is
(rows <- cell_dat$rows)
```

```
## Source: local data frame [1,000 x 11]
## 
##        r spans     s custom_format    ht hidden custom_height
##    <int> <chr> <int>         <lgl> <dbl>  <lgl>         <lgl>
## 1      1  <NA>    NA         FALSE    NA  FALSE            NA
## 2      2  <NA>    NA         FALSE    NA  FALSE            NA
## 3      3  <NA>    NA         FALSE    NA  FALSE            NA
## 4      4  <NA>    NA         FALSE    NA  FALSE            NA
## 5      5  <NA>    NA         FALSE    NA  FALSE            NA
## 6      6  <NA>    NA         FALSE    NA  FALSE            NA
## 7      7  <NA>    NA         FALSE    NA  FALSE            NA
## 8      8  <NA>    NA         FALSE    NA  FALSE            NA
## 9      9  <NA>    NA         FALSE    NA  FALSE            NA
## 10    10  <NA>    NA         FALSE    NA  FALSE            NA
## ..   ...   ...   ...           ...   ...    ...           ...
## Variables not shown: outline_level <int>, collapsed <lgl>, thick_top
##   <lgl>, thick_bot <lgl>.
```

```r
## this is where cells come from  
(cells <- linen::cells(cell_dat$cells$ref, cell_dat$cells$style,
                       cell_dat$cells$type, cell_dat$cells$value,
                       cell_dat$cells$formula))
```

```
## Source: local data frame [2,022 x 12]
## 
##      ref style     value formula   type is_formula is_value is_blank
##    <chr> <int>    <list>   <chr>  <chr>      <lgl>    <lgl>    <lgl>
## 1     A1     2 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## 2     B1     3 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## 3     C1     4 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## 4     D1     2 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## 5     E1     2 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## 6     F1     2 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## 7     A2     2 <dbl [1]>    <NA> number      FALSE     TRUE    FALSE
## 8     B2     3 <dbl [1]>    <NA> number      FALSE     TRUE    FALSE
## 9     C2     4 <dbl [1]>    <NA> number      FALSE     TRUE    FALSE
## 10    D2     2 <chr [1]>    <NA>   text      FALSE     TRUE    FALSE
## ..   ...   ...       ...     ...    ...        ...      ...      ...
## Variables not shown: is_bool <lgl>, is_number <lgl>, is_text <lgl>,
##   is_date <lgl>.
```

```r
## in real life and in other sheets, it's possible comments will be populated
## but not in this sheet
comments <- NULL
```

Now we gather everything we've learned about this worksheet into a `linen::worksheet` object.


```r
(ws <- linen::worksheet(sheet_name, cols, rows, cells, merged, view, comments,
                        workbook))
```

```
## <worksheet: 1000 x 6>
##  : ABCDEF
## 1: aaaaaa
## 2: 000a$$
## 3: 000 $$
## 4: 000a$$
## 5:  00a$$
## 6: 000a $
```

If we had other sheets to read, that would be done now. Ultimately this `workbook` is returned.
