# Read an Excel Sheet
Jennifer Bryan  
`r Sys.Date()`  




```r
library(rprojroot)
devtools::load_all(find_package_root_file())
```

```
## Loading rexcel
```

*Using a function I wrote while exploring all the files that make up an xlsx.*

Apply it to mini gapminder.


```r
mini_gap_path <- system.file("sheets", "mini-gap.xlsx", package = "rexcel")
mini_gap_workbook <- rexcel_workbook(mini_gap_path)
str(mini_gap_workbook, max.level = 1)
```

```
## List of 11
##  $ xlsx_path     : chr "/Users/jenny/rrr/rexcel/inst/sheets/mini-gap.xlsx"
##  $ reg_time      : POSIXct[1:1], format: "2016-05-30 15:23:09"
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
##  $ sheets_df     :Classes 'tbl_df', 'tbl' and 'data.frame':	5 obs. of  6 variables:
```

```r
mini_gap_workbook
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/mini-gap.xlsx"
## 
## $reg_time
## [1] "2016-05-30 15:23:09 PDT"
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
##      id
##   <chr>
## 1  rId2
## 2  rId1
## 3  rId4
## 4  rId3
## 5  rId6
## 6  rId5
## 7  rId7
## Variables not shown: type <chr>, target <chr>.
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
## Source: local data frame [5 x 6]
## 
##   sheet_id     name    id                   target   state
##      <int>    <chr> <chr>                    <chr>   <chr>
## 1        1   Africa  rId3 xl/worksheets/sheet4.xml visible
## 2        2 Americas  rId4 xl/worksheets/sheet3.xml visible
## 3        3     Asia  rId5 xl/worksheets/sheet5.xml visible
## 4        4   Europe  rId6 xl/worksheets/sheet1.xml visible
## 5        5  Oceania  rId7 xl/worksheets/sheet2.xml visible
## Variables not shown: type <chr>.
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
## [1] "2016-05-30 15:23:09 PDT"
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
##     name   state sheet_id    id
##    <chr>   <chr>    <int> <chr>
## 1 Sheet1 visible        1  rId3
## 
## $defined_names
## NULL
## 
## $workbook_rels
## Source: local data frame [3 x 3]
## 
##      id
##   <chr>
## 1  rId1
## 2  rId2
## 3  rId3
## Variables not shown: type <chr>, target <chr>.
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
## $worksheet_rels
## Source: local data frame [2 x 5]
## 
##   worksheet    id
##       <chr> <chr>
## 1    sheet1  rId1
## 2    sheet1  rId2
## Variables not shown: type <chr>, target <chr>, targetmode <chr>.
## 
## $sheets_df
## Source: local data frame [1 x 6]
## 
##   sheet_id   name    id                   target   state
##      <int>  <chr> <chr>                    <chr>   <chr>
## 1        1 Sheet1  rId3 xl/worksheets/sheet1.xml visible
## Variables not shown: type <chr>.
```


Apply it to the Ekaterinburg sheet from [readxl/#80](https://github.com/hadley/readxl/issues/80) and the "resaved in Excel" version.


```r
ek_path <- system.file("sheets", "Ekaterinburg_IP_9.xlsx", package = "rexcel")
rexcel_workbook(ek_path)
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/Ekaterinburg_IP_9.xlsx"
## 
## $reg_time
## [1] "2016-05-30 15:23:09 PDT"
## 
## $manifest
## Source: local data frame [8 x 3]
## 
##                                 Name  Length                Date
##                                <chr>   <dbl>              <time>
## 1                    xl/workbook.xml     314 2014-10-23 16:38:00
## 2                        _rels/.rels     296 2014-10-23 16:38:00
## 3               xl/sharedStrings.xml   41835 2014-10-23 16:38:00
## 4         xl/_rels/workbook.xml.rels     603 2014-10-23 16:38:00
## 5                      xl/styles.xml    4614 2014-10-23 16:38:00
## 6            xl/worksheets/sheet.xml 1922092 2014-10-23 16:38:00
## 7 xl/worksheets/_rels/sheet.xml.rels     322 2014-10-23 16:38:00
## 8                [Content_Types].xml     736 2014-10-23 16:38:00
## 
## $content_types
## Source: local data frame [5 x 3]
## 
##                   PartName Extension
##                      <chr>     <chr>
## 1                     <NA>       xml
## 2                     <NA>      rels
## 3    /xl/sharedStrings.xml      <NA>
## 4           /xl/styles.xml      <NA>
## 5 /xl/worksheets/sheet.xml      <NA>
## Variables not shown: ContentType <chr>.
## 
## $sheets
## Source: local data frame [1 x 4]
## 
##             name state sheet_id                id
##            <chr> <chr>    <int>             <chr>
## 1 СПАРК - Список  <NA>        1 R6082ddd3e995440f
## 
## $defined_names
## NULL
## 
## $workbook_rels
## Source: local data frame [3 x 3]
## 
##                                                                          type
##                                                                         <chr>
## 1 http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedS
## 2  http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles
## 3 http://schemas.openxmlformats.org/officeDocument/2006/relationships/workshe
## Variables not shown: target <chr>, id <chr>.
## 
## $shared_strings
##   [1] ""                                           
##   [2] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [3] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [4] " Исключен из Статрегистра ИП от 01.07.2014 "
##   [5] " Исключен из Статрегистра ИП от 01.07.2014 "
##   [6] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [7] " Исключен из Статрегистра ИП от 31.03.2014 "
##   [8] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [9] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [10] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [11] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [12] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [13] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [14] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [15] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [16] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [17] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [18] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [19] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [20] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [21] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [22] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [23] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [24] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [25] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [26] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [27] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [28] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [29] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [30] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [31] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [32] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [33] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [34] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [35] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [36] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [37] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [38] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [39] " Исключен из Статрегистра ИП от 27.02.2013 "
##  [40] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [41] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [42] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [43] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [44] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [45] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [46] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [47] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [48] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [49] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [50] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [51] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [52] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [53] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [54] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [55] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [56] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [57] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [58] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [59] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [60] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [61] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [62] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [63] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [64] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [65] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [66] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [67] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [68] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [69] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [70] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [71] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [72] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [73] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [74] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [75] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [76] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [77] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [78] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [79] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [80] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [81] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [82] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [83] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [84] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [85] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [86] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [87] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [88] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [89] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [90] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [91] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [92] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [93] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [94] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [95] " Исключен из Статрегистра ИП от 27.02.2013 "
##  [96] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [97] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [98] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [99] " Исключен из Статрегистра ИП от 28.02.2014 "
## [100] " Исключен из Статрегистра ИП от 01.09.2014 "
## [101] " Исключен из Статрегистра ИП от 01.09.2014 "
## [102] " Исключен из Статрегистра ИП от 28.02.2014 "
## [103] " Исключен из Статрегистра ИП от 01.07.2014 "
## [104] " Исключен из Статрегистра ИП от 01.07.2014 "
## [105] " Исключен из Статрегистра ИП от 31.12.2013 "
## [106] " Исключен из Статрегистра ИП от 31.03.2014 "
## [107] " Исключен из Статрегистра ИП от 31.03.2014 "
## [108] " Исключен из Статрегистра ИП от 31.01.2014 "
## [109] " Исключен из Статрегистра ИП от 31.12.2013 "
## [110] " Исключен из Статрегистра ИП от 31.12.2013 "
## [111] " Исключен из Статрегистра ИП от 01.09.2014 "
## [112] " Исключен из Статрегистра ИП от 01.09.2014 "
## [113] " Исключен из Статрегистра ИП от 31.01.2014 "
## [114] " Исключен из Статрегистра ИП от 31.01.2014 "
## [115] " Исключен из Статрегистра ИП от 31.12.2013 "
## [116] " Исключен из Статрегистра ИП от 28.02.2014 "
## [117] " Исключен из Статрегистра ИП от 31.12.2013 "
## [118] " Исключен из Статрегистра ИП от 31.12.2013 "
## [119] " Исключен из Статрегистра ИП от 31.01.2014 "
## [120] " Исключен из Статрегистра ИП от 01.09.2014 "
## [121] " Исключен из Статрегистра ИП от 27.02.2013 "
## [122] " Исключен из Статрегистра ИП от 31.01.2014 "
## [123] " Исключен из Статрегистра ИП от 31.03.2014 "
## [124] " Исключен из Статрегистра ИП от 31.12.2013 "
## [125] " Исключен из Статрегистра ИП от 01.09.2014 "
## [126] " Исключен из Статрегистра ИП от 31.12.2013 "
## [127] " Исключен из Статрегистра ИП от 01.07.2014 "
## [128] " Исключен из Статрегистра ИП от 31.03.2014 "
## [129] " Исключен из Статрегистра ИП от 31.12.2013 "
## [130] " Исключен из Статрегистра ИП от 31.12.2013 "
## [131] " Исключен из Статрегистра ИП от 31.12.2013 "
## [132] " Исключен из Статрегистра ИП от 31.12.2013 "
## [133] " Исключен из Статрегистра ИП от 27.02.2013 "
## [134] " Исключен из Статрегистра ИП от 31.03.2014 "
## [135] " Исключен из Статрегистра ИП от 01.09.2014 "
## [136] " Исключен из Статрегистра ИП от 01.09.2014 "
## [137] " Исключен из Статрегистра ИП от 01.07.2014 "
## [138] " Исключен из Статрегистра ИП от 31.12.2013 "
## [139] " Исключен из Статрегистра ИП от 01.09.2014 "
## [140] " Исключен из Статрегистра ИП от 01.07.2014 "
## [141] " Исключен из Статрегистра ИП от 31.05.2007 "
## [142] " Исключен из Статрегистра ИП от 31.12.2013 "
## [143] " Исключен из Статрегистра ИП от 01.09.2014 "
## [144] " Исключен из Статрегистра ИП от 31.03.2014 "
## [145] " Исключен из Статрегистра ИП от 31.01.2014 "
## [146] " Исключен из Статрегистра ИП от 31.03.2014 "
## [147] " Исключен из Статрегистра ИП от 01.09.2014 "
## [148] " Исключен из Статрегистра ИП от 01.09.2014 "
## [149] " Исключен из Статрегистра ИП от 31.12.2013 "
## [150] " Исключен из Статрегистра ИП от 31.03.2014 "
## [151] " Исключен из Статрегистра ИП от 31.01.2014 "
## [152] " Исключен из Статрегистра ИП от 01.09.2014 "
## [153] " Исключен из Статрегистра ИП от 31.03.2014 "
## [154] " Исключен из Статрегистра ИП от 31.03.2014 "
## [155] " Исключен из Статрегистра ИП от 27.02.2013 "
## [156] " Исключен из Статрегистра ИП от 31.12.2013 "
## [157] " Исключен из Статрегистра ИП от 28.02.2014 "
## [158] " Исключен из Статрегистра ИП от 01.09.2014 "
## [159] " Исключен из Статрегистра ИП от 01.09.2014 "
## [160] " Исключен из Статрегистра ИП от 28.02.2014 "
## [161] " Исключен из Статрегистра ИП от 31.12.2013 "
## [162] " Исключен из Статрегистра ИП от 01.09.2014 "
## [163] " Исключен из Статрегистра ИП от 31.12.2013 "
## [164] " Исключен из Статрегистра ИП от 01.09.2014 "
## [165] " Исключен из Статрегистра ИП от 01.07.2014 "
## attr(,"count")
## [1] 2191
## attr(,"uniqueCount")
## [1] 165
## 
## $styles
## $styles$fonts
## Source: local data frame [5 x 6]
## 
##      sz    color    name family charset scheme
##   <chr>    <chr>   <chr>  <chr>   <chr>  <chr>
## 1    10        1 Calibri      2     204  minor
## 2    10        1 Calibri      2     204  minor
## 3    14        3 Calibri      2     204  minor
## 4    10 FF0000FF Calibri      2     204  minor
## 5    10        1 Calibri      2     204  minor
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
## Source: local data frame [1 x 5]
## 
##   worksheet
##       <chr>
## 1     sheet
## Variables not shown: type <chr>, target <chr>, targetmode <chr>, id <chr>.
## 
## $sheets_df
## Source: local data frame [1 x 6]
## 
##   sheet_id           name                id                      target
##      <int>          <chr>             <chr>                       <chr>
## 1        1 СПАРК - Список R6082ddd3e995440f xl//xl/worksheets/sheet.xml
## Variables not shown: state <chr>, type <chr>.
```

```r
ek2_path <- system.file("sheets", "Ekaterinburg_IP_9-RESAVED.xlsx",
                        package = "rexcel")
rexcel_workbook(ek_path)
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/Ekaterinburg_IP_9.xlsx"
## 
## $reg_time
## [1] "2016-05-30 15:23:09 PDT"
## 
## $manifest
## Source: local data frame [8 x 3]
## 
##                                 Name  Length                Date
##                                <chr>   <dbl>              <time>
## 1                    xl/workbook.xml     314 2014-10-23 16:38:00
## 2                        _rels/.rels     296 2014-10-23 16:38:00
## 3               xl/sharedStrings.xml   41835 2014-10-23 16:38:00
## 4         xl/_rels/workbook.xml.rels     603 2014-10-23 16:38:00
## 5                      xl/styles.xml    4614 2014-10-23 16:38:00
## 6            xl/worksheets/sheet.xml 1922092 2014-10-23 16:38:00
## 7 xl/worksheets/_rels/sheet.xml.rels     322 2014-10-23 16:38:00
## 8                [Content_Types].xml     736 2014-10-23 16:38:00
## 
## $content_types
## Source: local data frame [5 x 3]
## 
##                   PartName Extension
##                      <chr>     <chr>
## 1                     <NA>       xml
## 2                     <NA>      rels
## 3    /xl/sharedStrings.xml      <NA>
## 4           /xl/styles.xml      <NA>
## 5 /xl/worksheets/sheet.xml      <NA>
## Variables not shown: ContentType <chr>.
## 
## $sheets
## Source: local data frame [1 x 4]
## 
##             name state sheet_id                id
##            <chr> <chr>    <int>             <chr>
## 1 СПАРК - Список  <NA>        1 R6082ddd3e995440f
## 
## $defined_names
## NULL
## 
## $workbook_rels
## Source: local data frame [3 x 3]
## 
##                                                                          type
##                                                                         <chr>
## 1 http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedS
## 2  http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles
## 3 http://schemas.openxmlformats.org/officeDocument/2006/relationships/workshe
## Variables not shown: target <chr>, id <chr>.
## 
## $shared_strings
##   [1] ""                                           
##   [2] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [3] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [4] " Исключен из Статрегистра ИП от 01.07.2014 "
##   [5] " Исключен из Статрегистра ИП от 01.07.2014 "
##   [6] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [7] " Исключен из Статрегистра ИП от 31.03.2014 "
##   [8] " Исключен из Статрегистра ИП от 31.01.2014 "
##   [9] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [10] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [11] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [12] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [13] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [14] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [15] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [16] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [17] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [18] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [19] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [20] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [21] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [22] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [23] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [24] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [25] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [26] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [27] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [28] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [29] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [30] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [31] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [32] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [33] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [34] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [35] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [36] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [37] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [38] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [39] " Исключен из Статрегистра ИП от 27.02.2013 "
##  [40] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [41] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [42] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [43] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [44] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [45] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [46] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [47] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [48] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [49] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [50] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [51] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [52] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [53] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [54] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [55] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [56] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [57] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [58] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [59] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [60] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [61] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [62] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [63] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [64] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [65] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [66] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [67] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [68] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [69] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [70] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [71] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [72] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [73] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [74] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [75] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [76] " Исключен из Статрегистра ИП от 28.02.2014 "
##  [77] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [78] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [79] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [80] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [81] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [82] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [83] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [84] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [85] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [86] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [87] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [88] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [89] " Исключен из Статрегистра ИП от 31.12.2013 "
##  [90] " Исключен из Статрегистра ИП от 31.01.2014 "
##  [91] " Исключен из Статрегистра ИП от 01.07.2014 "
##  [92] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [93] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [94] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [95] " Исключен из Статрегистра ИП от 27.02.2013 "
##  [96] " Исключен из Статрегистра ИП от 01.09.2014 "
##  [97] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [98] " Исключен из Статрегистра ИП от 31.03.2014 "
##  [99] " Исключен из Статрегистра ИП от 28.02.2014 "
## [100] " Исключен из Статрегистра ИП от 01.09.2014 "
## [101] " Исключен из Статрегистра ИП от 01.09.2014 "
## [102] " Исключен из Статрегистра ИП от 28.02.2014 "
## [103] " Исключен из Статрегистра ИП от 01.07.2014 "
## [104] " Исключен из Статрегистра ИП от 01.07.2014 "
## [105] " Исключен из Статрегистра ИП от 31.12.2013 "
## [106] " Исключен из Статрегистра ИП от 31.03.2014 "
## [107] " Исключен из Статрегистра ИП от 31.03.2014 "
## [108] " Исключен из Статрегистра ИП от 31.01.2014 "
## [109] " Исключен из Статрегистра ИП от 31.12.2013 "
## [110] " Исключен из Статрегистра ИП от 31.12.2013 "
## [111] " Исключен из Статрегистра ИП от 01.09.2014 "
## [112] " Исключен из Статрегистра ИП от 01.09.2014 "
## [113] " Исключен из Статрегистра ИП от 31.01.2014 "
## [114] " Исключен из Статрегистра ИП от 31.01.2014 "
## [115] " Исключен из Статрегистра ИП от 31.12.2013 "
## [116] " Исключен из Статрегистра ИП от 28.02.2014 "
## [117] " Исключен из Статрегистра ИП от 31.12.2013 "
## [118] " Исключен из Статрегистра ИП от 31.12.2013 "
## [119] " Исключен из Статрегистра ИП от 31.01.2014 "
## [120] " Исключен из Статрегистра ИП от 01.09.2014 "
## [121] " Исключен из Статрегистра ИП от 27.02.2013 "
## [122] " Исключен из Статрегистра ИП от 31.01.2014 "
## [123] " Исключен из Статрегистра ИП от 31.03.2014 "
## [124] " Исключен из Статрегистра ИП от 31.12.2013 "
## [125] " Исключен из Статрегистра ИП от 01.09.2014 "
## [126] " Исключен из Статрегистра ИП от 31.12.2013 "
## [127] " Исключен из Статрегистра ИП от 01.07.2014 "
## [128] " Исключен из Статрегистра ИП от 31.03.2014 "
## [129] " Исключен из Статрегистра ИП от 31.12.2013 "
## [130] " Исключен из Статрегистра ИП от 31.12.2013 "
## [131] " Исключен из Статрегистра ИП от 31.12.2013 "
## [132] " Исключен из Статрегистра ИП от 31.12.2013 "
## [133] " Исключен из Статрегистра ИП от 27.02.2013 "
## [134] " Исключен из Статрегистра ИП от 31.03.2014 "
## [135] " Исключен из Статрегистра ИП от 01.09.2014 "
## [136] " Исключен из Статрегистра ИП от 01.09.2014 "
## [137] " Исключен из Статрегистра ИП от 01.07.2014 "
## [138] " Исключен из Статрегистра ИП от 31.12.2013 "
## [139] " Исключен из Статрегистра ИП от 01.09.2014 "
## [140] " Исключен из Статрегистра ИП от 01.07.2014 "
## [141] " Исключен из Статрегистра ИП от 31.05.2007 "
## [142] " Исключен из Статрегистра ИП от 31.12.2013 "
## [143] " Исключен из Статрегистра ИП от 01.09.2014 "
## [144] " Исключен из Статрегистра ИП от 31.03.2014 "
## [145] " Исключен из Статрегистра ИП от 31.01.2014 "
## [146] " Исключен из Статрегистра ИП от 31.03.2014 "
## [147] " Исключен из Статрегистра ИП от 01.09.2014 "
## [148] " Исключен из Статрегистра ИП от 01.09.2014 "
## [149] " Исключен из Статрегистра ИП от 31.12.2013 "
## [150] " Исключен из Статрегистра ИП от 31.03.2014 "
## [151] " Исключен из Статрегистра ИП от 31.01.2014 "
## [152] " Исключен из Статрегистра ИП от 01.09.2014 "
## [153] " Исключен из Статрегистра ИП от 31.03.2014 "
## [154] " Исключен из Статрегистра ИП от 31.03.2014 "
## [155] " Исключен из Статрегистра ИП от 27.02.2013 "
## [156] " Исключен из Статрегистра ИП от 31.12.2013 "
## [157] " Исключен из Статрегистра ИП от 28.02.2014 "
## [158] " Исключен из Статрегистра ИП от 01.09.2014 "
## [159] " Исключен из Статрегистра ИП от 01.09.2014 "
## [160] " Исключен из Статрегистра ИП от 28.02.2014 "
## [161] " Исключен из Статрегистра ИП от 31.12.2013 "
## [162] " Исключен из Статрегистра ИП от 01.09.2014 "
## [163] " Исключен из Статрегистра ИП от 31.12.2013 "
## [164] " Исключен из Статрегистра ИП от 01.09.2014 "
## [165] " Исключен из Статрегистра ИП от 01.07.2014 "
## attr(,"count")
## [1] 2191
## attr(,"uniqueCount")
## [1] 165
## 
## $styles
## $styles$fonts
## Source: local data frame [5 x 6]
## 
##      sz    color    name family charset scheme
##   <chr>    <chr>   <chr>  <chr>   <chr>  <chr>
## 1    10        1 Calibri      2     204  minor
## 2    10        1 Calibri      2     204  minor
## 3    14        3 Calibri      2     204  minor
## 4    10 FF0000FF Calibri      2     204  minor
## 5    10        1 Calibri      2     204  minor
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
## Source: local data frame [1 x 5]
## 
##   worksheet
##       <chr>
## 1     sheet
## Variables not shown: type <chr>, target <chr>, targetmode <chr>, id <chr>.
## 
## $sheets_df
## Source: local data frame [1 x 6]
## 
##   sheet_id           name                id                      target
##      <int>          <chr>             <chr>                       <chr>
## 1        1 СПАРК - Список R6082ddd3e995440f xl//xl/worksheets/sheet.xml
## Variables not shown: state <chr>, type <chr>.
```

Apply it to a sheet created to play with defined ranges.


```r
dn_path <- system.file("sheets", "defined-names.xlsx", package = "rexcel")
rexcel_workbook(dn_path)
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/defined-names.xlsx"
## 
## $reg_time
## [1] "2016-05-30 15:23:09 PDT"
## 
## $manifest
## Source: local data frame [10 x 3]
## 
##                          Name Length       Date
##                         <chr>  <dbl>     <time>
## 1         [Content_Types].xml   1168 1980-01-01
## 2                 _rels/.rels    588 1980-01-01
## 3  xl/_rels/workbook.xml.rels    698 1980-01-01
## 4             xl/workbook.xml   1699 1980-01-01
## 5        xl/sharedStrings.xml    574 1980-01-01
## 6         xl/theme/theme1.xml   6788 1980-01-01
## 7               xl/styles.xml   1399 1980-01-01
## 8    xl/worksheets/sheet1.xml   4062 1980-01-01
## 9           docProps/core.xml    589 1980-01-01
## 10           docProps/app.xml    776 1980-01-01
## 
## $content_types
## Source: local data frame [9 x 3]
## 
##                    PartName Extension
##                       <chr>     <chr>
## 1                      <NA>       xml
## 2                      <NA>      rels
## 3          /xl/workbook.xml      <NA>
## 4 /xl/worksheets/sheet1.xml      <NA>
## 5      /xl/theme/theme1.xml      <NA>
## 6            /xl/styles.xml      <NA>
## 7     /xl/sharedStrings.xml      <NA>
## 8        /docProps/core.xml      <NA>
## 9         /docProps/app.xml      <NA>
## Variables not shown: ContentType <chr>.
## 
## $sheets
## Source: local data frame [1 x 4]
## 
##     name state sheet_id    id
##    <chr> <chr>    <int> <chr>
## 1 Sheet1  <NA>        1  rId1
## 
## $defined_names
## Source: local data frame [6 x 3]
## 
##        name         refers_to sheet_id
##       <chr>             <chr>    <int>
## 1 continent Sheet1!$B$2:$B$11       NA
## 2   country Sheet1!$A$2:$A$11       NA
## 3 gdpPercap Sheet1!$F$2:$F$11       NA
## 4   lifeExp Sheet1!$D$2:$D$11       NA
## 5       pop Sheet1!$E$2:$E$11       NA
## 6      year Sheet1!$C$2:$C$11       NA
## 
## $workbook_rels
## Source: local data frame [4 x 3]
## 
##      id
##   <chr>
## 1  rId3
## 2  rId4
## 3  rId1
## 4  rId2
## Variables not shown: type <chr>, target <chr>.
## 
## $shared_strings
##  [1] "country"      "continent"    "year"         "lifeExp"     
##  [5] "pop"          "gdpPercap"    "Algeria"      "Africa"      
##  [9] "Angola"       "Benin"        "Argentina"    "Americas"    
## [13] "Bolivia"      "Brazil"       "Canada"       "Chile"       
## [17] "Botswana"     "Burkina Faso"
## attr(,"count")
## [1] 26
## attr(,"uniqueCount")
## [1] 18
## 
## $styles
## $styles$fonts
## Source: local data frame [2 x 3]
## 
##      sz    color  name
##   <chr>    <chr> <chr>
## 1    10 FF000000 Arial
## 2    10     <NA> Arial
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
## NULL
## 
## $sheets_df
## Source: local data frame [1 x 6]
## 
##   sheet_id   name    id                   target state
##      <int>  <chr> <chr>                    <chr> <chr>
## 1        1 Sheet1  rId1 xl/worksheets/sheet1.xml  <NA>
## Variables not shown: type <chr>.
```

Apply it to a sheet created to explore tricky worksheet names.


```r
gabe_path <- system.file("sheets", "gabe.xlsx", package = "rexcel")
rexcel_workbook(gabe_path)
```

```
## $xlsx_path
## [1] "/Users/jenny/rrr/rexcel/inst/sheets/gabe.xlsx"
## 
## $reg_time
## [1] "2016-05-30 15:23:09 PDT"
## 
## $manifest
## Source: local data frame [12 x 3]
## 
##                          Name Length       Date
##                         <chr>  <dbl>     <time>
## 1         [Content_Types].xml   1356 1980-01-01
## 2                 _rels/.rels    733 1980-01-01
## 3  xl/_rels/workbook.xml.rels    839 1980-01-01
## 4             xl/workbook.xml   1702 1980-01-01
## 5               xl/styles.xml   1260 1980-01-01
## 6    xl/worksheets/sheet2.xml   1103 1980-01-01
## 7         xl/theme/theme1.xml   6788 1980-01-01
## 8    xl/worksheets/sheet1.xml   1171 1980-01-01
## 9     docProps/thumbnail.jpeg  14312 1980-01-01
## 10       xl/sharedStrings.xml    183 1980-01-01
## 11           docProps/app.xml    834 1980-01-01
## 12          docProps/core.xml    635 1980-01-01
## 
## $content_types
## Source: local data frame [11 x 3]
## 
##                     PartName Extension
##                        <chr>     <chr>
## 1                       <NA>       xml
## 2                       <NA>      rels
## 3                       <NA>      jpeg
## 4           /xl/workbook.xml      <NA>
## 5  /xl/worksheets/sheet1.xml      <NA>
## 6  /xl/worksheets/sheet2.xml      <NA>
## 7       /xl/theme/theme1.xml      <NA>
## 8             /xl/styles.xml      <NA>
## 9      /xl/sharedStrings.xml      <NA>
## 10        /docProps/core.xml      <NA>
## 11         /docProps/app.xml      <NA>
## Variables not shown: ContentType <chr>.
## 
## $sheets
## Source: local data frame [2 x 4]
## 
##             name state sheet_id    id
##            <chr> <chr>    <int> <chr>
## 1 Gabe's S''heet  <NA>        2  rId1
## 2     HiGabe!!!!  <NA>        1  rId2
## 
## $defined_names
## Source: local data frame [4 x 3]
## 
##      name                     refers_to sheet_id
##     <chr>                         <chr>    <int>
## 1   A_one      'Gabe''s S''''heet'!$A$1       NA
## 2   A_one             'HiGabe!!!!'!$A$1       NA
## 3 numbers 'Gabe''s S''''heet'!$A$2:$A$6       NA
## 4 numbers        'HiGabe!!!!'!$A$2:$A$6       NA
## 
## $workbook_rels
## Source: local data frame [5 x 3]
## 
##      id
##   <chr>
## 1  rId3
## 2  rId4
## 3  rId5
## 4  rId1
## 5  rId2
## Variables not shown: type <chr>, target <chr>.
## 
## $shared_strings
## [1] "numbers"
## attr(,"count")
## [1] 2
## attr(,"uniqueCount")
## [1] 1
## 
## $styles
## $styles$fonts
## Source: local data frame [1 x 5]
## 
##      sz color    name family scheme
##   <chr> <chr>   <chr>  <chr>  <chr>
## 1    12     1 Calibri      2  minor
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
## NULL
## 
## $sheets_df
## Source: local data frame [2 x 6]
## 
##   sheet_id           name    id                   target state
##      <int>          <chr> <chr>                    <chr> <chr>
## 1        2 Gabe's S''heet  rId1 xl/worksheets/sheet1.xml  <NA>
## 2        1     HiGabe!!!!  rId2 xl/worksheets/sheet2.xml  <NA>
## Variables not shown: type <chr>.
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
