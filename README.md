# rexcel
[![Build Status](https://travis-ci.org/rsheets/rexcel.svg?branch=master)](https://travis-ci.org/rsheets/rexcel)

**Warning: This project is an experiment; do not use for anything other than amusement/frustration purposes.**

## Design

This package implements a **very slow**, but thorough, Excel (xlsx) reader.  If you have a rectangular region of cells to read you will be better off with another Excel reading package such as: [`readxl`](http://cran.r-project.org/package=readxl),
[`openxlsx`](http://cran.r-project.org/package=openxlsx),
[`XLConnect`](http://cran.r-project.org/package=XLConnect),
[`xlsx`](http://cran.r-project.org/package=xlsx),
[`gdata`](http://cran.r-project.org/package=gdata), [`RODBC`](http://cran.r-project.org/package=RODBC), or possibly even
[`excel.link`](http://cran.r-project.org/package=excel.link), [`WriteXLS`](http://cran.r-project.org/package=WriteXLS), [`table1xlsx`](http://cran.r-project.org/package=table1xlsx), [`tablaxlsx`](http://cran.r-project.org/package=tablaxlsx) (not clear how current these last 4 are). Mango Solutions has a nice review article, [R: the Excel Connection](http://www.mango-solutions.com/wp/2015/05/r-the-excel-connection/), in which they compare several of the above packages, with a special emphasis on those that can both read and write Excel files (XLConnect, xlsx, openxlsx, excel.link).

Compared with the above packages, `rexcel` tries to read all the data from an Excel sheet using [`linen`](https://github.com/rsheets/linen) as an intermediate representation in R. The eventual goal is to provide a common receptacle for detailed spreadsheet information from both Excel and Google Sheets.  Rather than trying to create a single data.frame in one shot, it allows access to data, formulae and formatting information.  Excel type information is preserved, especially for heterogeneous columns.  It has no non-R dependencies (e.g. on Perl or Java) and should run on any platform regardless of whether Excel is installed.

## Installation

Requires the development version of `xml2`, as the newer version changes the behaviour of matching functions in fairly large ways.  For terminal printing we use the most recent copy of `crayon` which includes more accurate rendering.  And we use `linen` as the R spreadsheet representation.

```r
devtools::install_github("hadley/xml2")
devtools::install_github("gaborcsardi/crayon")
devtools::install_github("rsheets/linen")
devtools::install_github("rsheets/rexcel")
```

## Formatting preserved

* [x] Cell fill colour
* [x] Cell patterns
* [ ] Cell _gradients_
* [x] Text colour
* [x] Text bold, italic, underline, strikethrough, outline, shadow, condense, extend
* [x] Text font
* [x] Text size
* [x] Text alignment (horizontal, vertical)
* [x] Column/row visibility
* [x] Column/row width/height
* [ ] Styles applied at the column level (though the spec seems vague about if that's a real thing - compare p. 1600 and 1596)
* [ ] Conditional formatting (the rule and the outcome)
* [x] Borders (position and colour)
* [ ] Numeric/date formatting, possibly also formatted text?
* [ ] Table styles (e.g. for pivot tables)

Of particular concern:

> A cell can have both direct formatting (e.g., bold) and a cell style (e.g., Explanatory) applied to it. Therefore, both the cell style xf records and cell xf records shall be read to understand the full set of formatting applied to a cell.

(18.8.10, p.1754)

Also

> When the color palette is modified, the indexedColors collection is written. When a custom color has been selected, the mruColors collection is written.

So we should find a sheet that includes this and see what this looks like empirically.

## Other worthwhile things to get

* [x] named ranges
* [x] comments (author, ref, rich text, visibility)
* [ ] graphs
* [ ] other drawings
* [ ] pivot tables
* [ ] frozen rows / split panes
* [ ] calculation chain
* [ ] header/footer

## The Excel XML Spec

Page and section numbers refer to the "ECMA Office Open XML Part 1 - Fundamentals and Markup Language Reference" document; a massive and fairly unweidly 5026 PDF!  The spreadsheet material is mostly in section 18 (pages 1518 - 2508).

## Writing Excel files

Writing workbooks is not currently supported.  Before implementing it, we want to see how much information we can preserve while reading.  Once we know that, we can start working towards seeing how much can be written and how lossy a read/write cycle will be (it is likely to be *very* lossy as there is an enormous number of things that might be stored in a worksheet.
