# rexcel

## Design

This package implements a **very slow**, but thorough, Excel (xlsx) reader.  If you have a rectangular region of cells to read you will be better off with another Excel reading package such as: [`readxl`](http://cran.r-project.org/package=readxl), [`gdata`](http://cran.r-project.org/package=gdata), [`openxlsx`](http://cran.r-project.org/package=readxl), [`xlsx`](http://cran.r-project.org/package=xlsx),
[`XLConnect`](http://cran.r-project.org/package=XLConnect), or possibly even [`excel.link`](http://cran.r-project.org/package=excel.link), or possibly [`WriteXLS`](http://cran.r-project.org/package=WriteXLS), [`table1xlsx`](http://cran.r-project.org/package=table1xlsx), [`tablaxlsx`](http://cran.r-project.org/package=tablaxlsx) (I have not checked the last 4 to see if they offer any real read functionality).

Compared with the many other packages, `rexcel` tries to read all the data from an Excel sheet using `linen` as an intermediate representation in R, which allows non-tabular data to be explored.  Rather than trying to create a single data.frame in one shot, it allows access to data, formulae and formatting information.  Excel type information is preserved, especially for heterogeneous columns.  It has no non-R dependencies (e.g. on Perl or Java) and should run on any platform regardless of whether Excel is installed.

## Installation

Requires the development version of xml2 (for `xml_find_lgl`):

```r
devtools::install_github("hadley/xml2")
devtools::install_github("jennybc/rexcel")
```

## Formatting preserved

* [x] Cell fill colour
* [ ] Cell patterns
* [ ] Cell _gradients_
* [ ] Text colour
* [ ] Text bold, italic, underline, strikethrough
* [ ] Text font
* [ ] Text size
* [ ] Text alignment (horizontal, vertical)
* [ ] Column/row visibility
* [ ] Column/row width/height
* [ ] Styles applied at the column level (though the spec seems vague about if that's a real thing - compare p. 1600 and 1596)
* [ ] Conditional formatting (the rule and the outcome)
* [ ] Borders

Of particular concern:

> A cell can have both direct formatting (e.g., bold) and a cell style (e.g., Explanatory) applied to it. Therefore, both the cell style xf records and cell xf records shall be read to understand the full set of formatting applied to a cell.

(18.8.10, p.1754)

Also

> When the color palette is modified, the indexedColors collection is written. When a custom color has been selected, the mruColors collection is written.

So we should find a sheet that includes this and see what this looks like empirically.

## Other worthwhile things to get

* [ ] named ranges
* [ ] comments (author, ref, rich text, visibility)
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
