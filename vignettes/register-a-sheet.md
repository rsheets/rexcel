# Register an Excel Sheet
Jennifer Bryan  
`r Sys.Date()`  




```r
library(rprojroot)
devtools::load_all(find_package_root_file())
```

```
## Loading rexcel
```

A walk through the code in `rexcel_register()`. I wrote this function to make myself better acquainted with the files that make up an xlsx and with `rexcel`. It could evolve into something useful and/or some of this could be incorporated into `rexcel_read_workbook()` or `rexcel_read_worksheet()`.

Philosophy, at least in theory:

  * Don't error unless it's an invalid xlsx file. Try to do something useful, even if it's mostly informative messages about why something can't be read in.
  * Inner functions are typically of this form: `xlsx_read_*()`.
    * They always take `path` to xlsx as primary or only argument.
    - They read from exactly one file.
    - We have one to read every file that's there. *probably very untrue, in practice!*
  * Process data read from xlsx files only enough to create the minimal R object.
  * Inner functions always return a single object.
  * Why all of this? So it's easy to "drop in" on a problematic xlsx or to work on `rexcel`. I find it hard to do this when low-level functions work with lists that combine different objects and objects that come from different files. If you haven't run a bunch of other internal code first (or if some of that failed), it's hard to get into a good place for figuring out what's wrong.

### Example sheets

We illustrate different xlsx features using different example sheets. Pre-store their paths to make this easier.


```r
mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
ff_path <- system.file("sheets", "gs-test-formula-formatting.xlsx",
                       package = "rexcel")
ek_path <- system.file("sheets", "Ekaterinburg_IP_9.xlsx", package = "rexcel")
ek2_path <- system.file("sheets", "Ekaterinburg_IP_9-RESAVED.xlsx",
                        package = "rexcel")
dn_path <- system.file("sheets", "defined-names.xlsx", package = "rexcel")
gabe_path <- system.file("sheets", "gabe.xlsx", package = "rexcel")
```

### Walk through `rexcel_register()`

`path` is path to the xlsx and is the only argument. We'll use different example sheets as we go. Let's start with mini Gapminder.

`is_xlsx()` runs sanity checks to make sure this looks like valid xlsx.


```r
is_xlsx(mini_gap_path)
```

```
## [1] TRUE
```

### `manifest` = file list

`manifest` holds a list of files in the zip archive.


```r
manifest <- xlsx_list_files(mini_gap_path)
print(manifest, n = Inf)
```

```
## Source: local data frame [21 x 3]
## 
##                                   name length                date
##                                  <chr>  <dbl>              <time>
## 1                  [Content_Types].xml   2005 2015-04-25 12:00:00
## 2                          _rels/.rels    296 2015-04-25 12:00:00
## 3           xl/_rels/workbook.xml.rels   1129 2015-04-25 12:00:00
## 4    xl/drawings/worksheetdrawing1.xml    494 2015-04-25 12:00:00
## 5    xl/drawings/worksheetdrawing2.xml    494 2015-04-25 12:00:00
## 6    xl/drawings/worksheetdrawing3.xml    494 2015-04-25 12:00:00
## 7    xl/drawings/worksheetdrawing4.xml    494 2015-04-25 12:00:00
## 8    xl/drawings/worksheetdrawing5.xml    494 2015-04-25 12:00:00
## 9                 xl/sharedStrings.xml    942 2015-04-25 12:00:00
## 10                       xl/styles.xml   1070 2015-04-25 12:00:00
## 11                     xl/workbook.xml    980 2015-04-25 12:00:00
## 12 xl/worksheets/_rels/sheet1.xml.rels    307 2015-04-25 12:00:00
## 13 xl/worksheets/_rels/sheet2.xml.rels    307 2015-04-25 12:00:00
## 14 xl/worksheets/_rels/sheet3.xml.rels    307 2015-04-25 12:00:00
## 15 xl/worksheets/_rels/sheet4.xml.rels    307 2015-04-25 12:00:00
## 16 xl/worksheets/_rels/sheet5.xml.rels    307 2015-04-25 12:00:00
## 17            xl/worksheets/sheet1.xml   2136 2015-04-25 12:00:00
## 18            xl/worksheets/sheet2.xml   2136 2015-04-25 12:00:00
## 19            xl/worksheets/sheet3.xml   2146 2015-04-25 12:00:00
## 20            xl/worksheets/sheet4.xml   2136 2015-04-25 12:00:00
## 21            xl/worksheets/sheet5.xml   2144 2015-04-25 12:00:00
```

Overview of manifests I've seen. *TO DO: cross-check/enhance this with some actual research in the spec or other resources we like.*

  * Workbook infrastructure
    - [Content_Types].xml
    - _rels/.rels *boring? I don't currently process*
    - xl/workbook.xml
    - xl/_rels/workbook.xml.rels
    - xl/sharedStrings.xml *doesn't necessarily exist*
    - xl/styles.xml
    - docProps/core.xml *have never looked at one of these; we have one in defined_names.xlsx*
    - docProps/app.xml *ditto*
  * Worksheet: one main file per sheet
    - xl/worksheets/sheet1.xml, xl/worksheets/sheet2.xml, ... is typical
    - but xl/worksheets/sheet.xml is also possible!
    - this is where sheet data actually lives, up to complications like the shared strings
  * Worksheet: a `rels` file for each sheet
    - xl/worksheets/_rels/sheet1.xml.rels and so on
  * Worksheet: a file of drawings?
    - xl/drawings/worksheetdrawing1.xml and so on
    - *you don't necessarily have these, but I see them even when there are no "drawings"*

The Ekaterinburg sheet from [readxl/#80](https://github.com/hadley/readxl/issues/80) has unusual structure. It was created by an undisclosed BI system but I include it here because the R packages that [wrap the Apache POI](https://poi.apache.org/spreadsheet/index.html) can read it just fine. So we should be able to return something informative for it.

Note the single sheet is referred to as `sheet`, not `sheet1`, and there is no associated `drawings` file.


```r
print(xlsx_list_files(ek_path), n = Inf)
```

```
## Source: local data frame [8 x 3]
## 
##                                 name  length                date
##                                <chr>   <dbl>              <time>
## 1                [Content_Types].xml     736 2014-10-23 16:38:00
## 2                        _rels/.rels     296 2014-10-23 16:38:00
## 3         xl/_rels/workbook.xml.rels     603 2014-10-23 16:38:00
## 4               xl/sharedStrings.xml   41835 2014-10-23 16:38:00
## 5                      xl/styles.xml    4614 2014-10-23 16:38:00
## 6                    xl/workbook.xml     314 2014-10-23 16:38:00
## 7 xl/worksheets/_rels/sheet.xml.rels     322 2014-10-23 16:38:00
## 8            xl/worksheets/sheet.xml 1922092 2014-10-23 16:38:00
```

If you open that workbook in Excel and resave it, things look different.


```r
print(xlsx_list_files(ek2_path), n = Inf)
```

```
## Source: local data frame [11 x 3]
## 
##                                   name  length       date
##                                  <chr>   <dbl>     <time>
## 1                  [Content_Types].xml    1168 1980-01-01
## 2                          _rels/.rels     588 1980-01-01
## 3                     docProps/app.xml     795 1980-01-01
## 4                    docProps/core.xml     589 1980-01-01
## 5           xl/_rels/workbook.xml.rels     698 1980-01-01
## 6                 xl/sharedStrings.xml  589203 1980-01-01
## 7                        xl/styles.xml    2856 1980-01-01
## 8                  xl/theme/theme1.xml    6788 1980-01-01
## 9                      xl/workbook.xml    1336 1980-01-01
## 10 xl/worksheets/_rels/sheet1.xml.rels     324 1980-01-01
## 11            xl/worksheets/sheet1.xml 1072131 1980-01-01
```

We gain `docProps/app.xml`, `docProps/core.xml`, and `xl/theme/theme1.xml` and the single sheet is now referred to as `sheet1`, not just `sheet`. Still no drawings, though.

Conclusion: there is a lot of variety in the manifest for valid xlsx.

### [Content_Types].xml

The `ct` object created from `[Content_Types].xml` is a tibble associating content types with extensions or specific files:

  * two "general" rows for the extensions `.xml` and `.rels`
  * other rows for specific files seen in the manifest *I gather these override the general types associated with extensions?*


```r
(ct <- as.data.frame(xlsx_read_Content_Types(mini_gap_path)))
```

```
##                             part_name extension
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
##                                                                     content_type
## 1                       application/vnd.openxmlformats-package.relationships+xml
## 2                                                                application/xml
## 3                      application/vnd.openxmlformats-officedocument.drawing+xml
## 4                      application/vnd.openxmlformats-officedocument.drawing+xml
## 5                      application/vnd.openxmlformats-officedocument.drawing+xml
## 6                      application/vnd.openxmlformats-officedocument.drawing+xml
## 7                      application/vnd.openxmlformats-officedocument.drawing+xml
## 8         application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml
## 9  application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml
## 10    application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml
## 11     application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml
## 12     application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml
## 13     application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml
## 14     application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml
## 15     application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml
```

```r
#setdiff(manifest$name, gsub("^\\/", "", ct$part_name))
#intersect(gsub("^\\/", "", ct$part_name), manifest$name)
```

### `sheets` info from xl/workbook.xml

Among many other things in `xl/workbook.xml`, there is one xml node per worksheet, which we use to make `sheets`. It's a tibble with one row per worksheet and these variables:

  * `name`: e.g., "Africa" *I assume this is name of the tab*
  * `state`: "visible" *or what else ... "invisible"?; might be `NA`*
  * `sheet_id`: integer *I assume this is order perceived by user*
  * `id`: character, e.g., `"rId5"` *a key that comes up elsewhere*


```r
(sheets <- xlsx_read_workbook_sheets(mini_gap_path))
```

```
## Source: local data frame [5 x 4]
## 
##       name   state sheet_id    id
##      <chr>   <chr>    <int> <chr>
## 1   Africa visible        1  rId3
## 2 Americas visible        2  rId4
## 3     Asia visible        3  rId5
## 4   Europe visible        4  rId6
## 5  Oceania visible        5  rId7
```

Quick tour of similar info for other example sheets:


```r
xlsx_read_workbook_sheets(ff_path)
```

```
## Source: local data frame [1 x 4]
## 
##     name   state sheet_id    id
##    <chr>   <chr>    <int> <chr>
## 1 Sheet1 visible        1  rId3
```

```r
xlsx_read_workbook_sheets(ek_path)
```

```
## Source: local data frame [1 x 4]
## 
##             name state sheet_id                id
##            <chr> <chr>    <int>             <chr>
## 1 СПАРК - Список  <NA>        1 R6082ddd3e995440f
```

```r
xlsx_read_workbook_sheets(ek2_path)
```

```
## Source: local data frame [1 x 4]
## 
##             name state sheet_id    id
##            <chr> <chr>    <int> <chr>
## 1 СПАРК - Список  <NA>        1  rId1
```

```r
xlsx_read_workbook_sheets(dn_path)
```

```
## Source: local data frame [1 x 4]
## 
##     name state sheet_id    id
##    <chr> <chr>    <int> <chr>
## 1 Sheet1  <NA>        1  rId1
```

```r
xlsx_read_workbook_sheets(gabe_path)
```

```
## Source: local data frame [2 x 4]
## 
##             name state sheet_id    id
##            <chr> <chr>    <int> <chr>
## 1 Gabe's S''heet  <NA>        2  rId1
## 2     HiGabe!!!!  <NA>        1  rId2
```

### Named ranges from xl/workbook.xml

Cell ranges can be named in Excel and subsequently used in formulas. These are described in `xl/workbook.xml` in the `definedName` nodes. Jenny has seen one node structure in the example she created, which differs from what Rich anticipated (presumably based on the spec?). Furthermore, Ekaterinburg has novel namespacing as well. See the comments in source for `xlsx_read_workbook_defined_names()` for details. Expect future pain here.

If a workbook has no named ranges, this will be `NULL`, e.g., as for mini Gapminder. Currently we have two example sheets with named ranges, `defined-names.xlsx` and `gabe.xlsx`. `defined-names.xlsx` was a planned example sheet. `gabe.xlsx` was intended just to explore weird worksheet names but, since it was copied from `defined-names.xlsx` and then worksheets got copied again, it incidentally shows what happens when there are replicated range names.


```r
xlsx_read_workbook_defined_names(mini_gap_path)
```

```
## NULL
```

```r
xlsx_read_workbook_defined_names(dn_path)
```

```
## Source: local data frame [6 x 4]
## 
##        name         refers_to sheet_id local_sheet_id
##       <chr>             <chr>    <int>          <int>
## 1 continent Sheet1!$B$2:$B$11       NA             NA
## 2   country Sheet1!$A$2:$A$11       NA             NA
## 3 gdpPercap Sheet1!$F$2:$F$11       NA             NA
## 4   lifeExp Sheet1!$D$2:$D$11       NA             NA
## 5       pop Sheet1!$E$2:$E$11       NA             NA
## 6      year Sheet1!$C$2:$C$11       NA             NA
```

```r
xlsx_read_workbook_defined_names(gabe_path)
```

```
## Source: local data frame [4 x 4]
## 
##      name                     refers_to sheet_id local_sheet_id
##     <chr>                         <chr>    <int>          <int>
## 1   A_one      'Gabe''s S''''heet'!$A$1       NA              0
## 2   A_one             'HiGabe!!!!'!$A$1       NA             NA
## 3 numbers 'Gabe''s S''''heet'!$A$2:$A$6       NA              0
## 4 numbers        'HiGabe!!!!'!$A$2:$A$6       NA             NA
```

`defined_names` is a tibble with one row per named range and these variables:

  * `name`: name of the named range
  * `refers_to`: string representation of the cell (area) reference, e.g., `Sheet1!$B$2:$B$11`
  * `sheet_id`: integer *I can't get my hands on a sheet that has actually this*
  * `local_sheet_id`: integer *appears when names are replicated*

### Workbook rels from xl/_rels/workbook.xml.rels

*We should probably cook up a more interesting example here?*

Mini Gapminder is interesting because you see the sheets aren't numbered exactly as you'd expect (which is [`hadley/readxl#104`](https://github.com/hadley/readxl/issues/104)). Ekaterinburg is also interesting because `target` has a leading slash and includes the `xl` subdirectory. But the re-saved version looks more conventional.


```r
(workbook_rels <- xlsx_read_workbook_rels(mini_gap_path))
```

```
## Source: local data frame [7 x 3]
## 
##                  target    id
##                   <chr> <chr>
## 1     sharedStrings.xml  rId2
## 2            styles.xml  rId1
## 3 worksheets/sheet3.xml  rId4
## 4 worksheets/sheet4.xml  rId3
## 5 worksheets/sheet1.xml  rId6
## 6 worksheets/sheet5.xml  rId5
## 7 worksheets/sheet2.xml  rId7
## Variables not shown: type <chr>.
```

```r
xlsx_read_workbook_rels(ek_path)
```

```
## Source: local data frame [3 x 3]
## 
##                     target                id
##                      <chr>             <chr>
## 1    /xl/sharedStrings.xml Rfd88d28f71e84f97
## 2           /xl/styles.xml Ra71ceb88d7f8404d
## 3 /xl/worksheets/sheet.xml R6082ddd3e995440f
## Variables not shown: type <chr>.
```

```r
xlsx_read_workbook_rels(ek2_path)
```

```
## Source: local data frame [4 x 3]
## 
##                  target    id
##                   <chr> <chr>
## 1            styles.xml  rId3
## 2     sharedStrings.xml  rId4
## 3 worksheets/sheet1.xml  rId1
## 4      theme/theme1.xml  rId2
## Variables not shown: type <chr>.
```

`workbook_rels` is a tibble, each row a file, with variables

  * `target`: a file path relative to `xl/` *maybe I should prepend xl/?*
  * `id`: character, e.g., `"rId5"` *a key that occurs elsewhere*
  * `type`: a long namespace-y string, the last bit of which tells you
if the associated file is `sharedStrings`, styles, or a worksheet, e.g.,
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`
*maybe I should retain just the last bit?*

### Derived object: `sheets_df`

We depart from the philosophy a bit here to create `sheets_df`, a new tibble with one row per worksheet. It joins `sheets`, which came from `workbook.xml`, to `workbook_rels`, which came from `workbook.xml.rels`. This is where we finally get sheet id, sheet name, id, and xml target file all in one place.


```r
(sheets_df <- join_sheets_workbook_rels(sheets, workbook_rels))
```

```
## Source: local data frame [5 x 5]
## 
##   sheet_id     name    id                   target   state
##      <int>    <chr> <chr>                    <chr>   <chr>
## 1        1   Africa  rId3 xl/worksheets/sheet4.xml visible
## 2        2 Americas  rId4 xl/worksheets/sheet3.xml visible
## 3        3     Asia  rId5 xl/worksheets/sheet5.xml visible
## 4        4   Europe  rId6 xl/worksheets/sheet1.xml visible
## 5        5  Oceania  rId7 xl/worksheets/sheet2.xml visible
```

### Shared strings from xl/sharedStrings.xml

Strings do not appear in the main sheet data files but rather appear exactly once in `sharedStrings.xml` and are then referenced.


```r
(shared_strings <- xlsx_read_shared_strings(mini_gap_path))
```

```
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
```

`shared_strings` is a character vector, with attributes `count` (total # of strings?) and `uniqueCount` (its own length?).

Let's look at some others. **I have no idea what's going on with Ekaterinburg original vs re-saved!** Why do `count` and `uniqueCount` both blow up?


```r
xlsx_read_shared_strings(ff_path)
```

```
##  [1] "integer"           "number_formatted"  "number_rounded"   
##  [4] "character"         "formula"           "formula_formatted"
##  [7] "one"               "three"             "four"             
## [10] "five"             
## attr(,"count")
## [1] 10
## attr(,"uniqueCount")
## [1] 10
```

```r
str(xlsx_read_shared_strings(ek_path))
```

```
##  atomic [1:165]   Исключен из Статрегистра ИП от 31.01.2014   Исключен из Статрегистра ИП от 31.01.2014   Исключен из Статрегистра ИП от 01.07.2014  ...
##  - attr(*, "count")= int 2191
##  - attr(*, "uniqueCount")= int 165
```

```r
str(xlsx_read_shared_strings(ek2_path))
```

```
##  atomic [1:11354]  Исключен из Статрегистра ИП от 31.01.2014 Исключен из Статрегистра ИП от 31.01.2014 Исключен из Статрегистра ИП от 01.07.2014 ...
##  - attr(*, "count")= int 24114
##  - attr(*, "uniqueCount")= int 11354
```

```r
xlsx_read_shared_strings(gabe_path)
```

```
## [1] "numbers"
## attr(,"count")
## [1] 2
## attr(,"uniqueCount")
## [1] 1
```



```r
mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
mini_gap_workbook <- rexcel_register(mini_gap_path)
str(mini_gap_workbook, max.level = 1)
```

```
## List of 11
##  $ xlsx_path     : chr "/Users/jenny/rrr/rexcel/inst/sheets/mini-gap.xlsx"
##  $ reg_time      : POSIXct[1:1], format: "2016-06-01 23:41:18"
##  $ manifest      :Classes 'tbl_df', 'tbl' and 'data.frame':	21 obs. of  3 variables:
##  $ content_types :Classes 'tbl_df', 'tbl' and 'data.frame':	15 obs. of  3 variables:
##  $ sheets        :Classes 'tbl_df', 'tbl' and 'data.frame':	5 obs. of  4 variables:
##  $ defined_names : NULL
##  $ workbook_rels :Classes 'tbl_df', 'tbl' and 'data.frame':	7 obs. of  3 variables:
##  $ shared_strings: atomic [1:33] country continent year lifeExp ...
##   ..- attr(*, "count")= int 80
##   ..- attr(*, "uniqueCount")= int 33
##  $ styles        :List of 8
##  $ worksheet_rels:Classes 'tbl_df', 'tbl' and 'data.frame':	5 obs. of  4 variables:
##  $ sheets_df     :Classes 'tbl_df', 'tbl' and 'data.frame':	5 obs. of  5 variables:
```

```r
mini_gap_workbook
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/mini-gap.xlsx"
## 
## $reg_time
## [1] "2016-06-01 23:41:18 PDT"
## 
## $manifest
## Source: local data frame [21 x 3]
## 
##                                 name length                date
##                                <chr>  <dbl>              <time>
## 1                [Content_Types].xml   2005 2015-04-25 12:00:00
## 2                        _rels/.rels    296 2015-04-25 12:00:00
## 3         xl/_rels/workbook.xml.rels   1129 2015-04-25 12:00:00
## 4  xl/drawings/worksheetdrawing1.xml    494 2015-04-25 12:00:00
## 5  xl/drawings/worksheetdrawing2.xml    494 2015-04-25 12:00:00
## 6  xl/drawings/worksheetdrawing3.xml    494 2015-04-25 12:00:00
## 7  xl/drawings/worksheetdrawing4.xml    494 2015-04-25 12:00:00
## 8  xl/drawings/worksheetdrawing5.xml    494 2015-04-25 12:00:00
## 9               xl/sharedStrings.xml    942 2015-04-25 12:00:00
## 10                     xl/styles.xml   1070 2015-04-25 12:00:00
## ..                               ...    ...                 ...
## 
## $content_types
## Source: local data frame [15 x 3]
## 
##                             part_name extension
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
## Variables not shown: content_type <chr>.
## 
## $sheets
## Source: local data frame [5 x 4]
## 
##       name   state sheet_id    id
##      <chr>   <chr>    <int> <chr>
## 1   Africa visible        1  rId3
## 2 Americas visible        2  rId4
## 3     Asia visible        3  rId5
## 4   Europe visible        4  rId6
## 5  Oceania visible        5  rId7
## 
## $defined_names
## NULL
## 
## $workbook_rels
## Source: local data frame [7 x 3]
## 
##                  target    id
##                   <chr> <chr>
## 1     sharedStrings.xml  rId2
## 2            styles.xml  rId1
## 3 worksheets/sheet3.xml  rId4
## 4 worksheets/sheet4.xml  rId3
## 5 worksheets/sheet1.xml  rId6
## 6 worksheets/sheet5.xml  rId5
## 7 worksheets/sheet2.xml  rId7
## Variables not shown: type <chr>.
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
## $worksheet_rels
## Source: local data frame [5 x 4]
## 
##   worksheet    id
##       <chr> <chr>
## 1    sheet1  rId1
## 2    sheet2  rId1
## 3    sheet3  rId1
## 4    sheet4  rId1
## 5    sheet5  rId1
## Variables not shown: type <chr>, target <chr>.
## 
## $sheets_df
## Source: local data frame [5 x 5]
## 
##   sheet_id     name    id                   target   state
##      <int>    <chr> <chr>                    <chr>   <chr>
## 1        1   Africa  rId3 xl/worksheets/sheet4.xml visible
## 2        2 Americas  rId4 xl/worksheets/sheet3.xml visible
## 3        3     Asia  rId5 xl/worksheets/sheet5.xml visible
## 4        4   Europe  rId6 xl/worksheets/sheet1.xml visible
## 5        5  Oceania  rId7 xl/worksheets/sheet2.xml visible
```

What's here?

  * `xlsx_path`: path to the xlsx
  * `reg_time`: time xlsx was processed
  * `manifest`: file list for the xlsx zip archive
  * `content_types`: tbl representing `[Content_Types].xml`
  * `sheets`: tbl representing `xl/workbook.xml`
  * `sheets_df`:
    - This is really the only thing I created.
    - A tbl with one row per worksheet, from joining `sheets` and `workbook_rels`
  * `shared_strings`: character vector representing `xl/sharedStrings.xml`
  * `styles`: list of tbls from `xl/styles.xml` (ok, I admit, I stopped after parsing fonts)
  * `workbook_rels`:
    - tbl that links target files to `Id`s, also gives file type
    - example: tells you that `Id = rId4` refers to `Target` file `xl/worksheets/sheetX.xml`
    - comes from `xl/_rels/workbook.xml.rels`
  * `worksheet_rels`:
    - *I'm still figuring this one out but it's about files or external resources (potentially) referred to from worksheets*
    - comes from files like` xl/worksheets/_rels/(sheet[0-9]+).xml.rels`
  


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
## attr(,"count")
## [1] 10
## attr(,"uniqueCount")
## [1] 10
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
