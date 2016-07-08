context("xlsx utilities")

iris_path <- file.path(get_readxl(), "tests", "testthat", "iris-excel.xlsx")

test_that("xlsx is detected as such and vice versa", {

  expect_true(is_xlsx(iris_path))

  expect_error(is_xlsx("nonexistent_path"), "does not exist")
  expect_error(is_xlsx(system.file(package = "linen")), "cannot be opened")
  zf <- tempfile()
  utils::zip(zf,
             files = dir(system.file(package = "linen"), full.names = TRUE),
             extras = "--quiet")
  expect_false(is_xlsx(paste0(zf, ".zip")))
})

test_that("xlsx files are listed", {
  iris_files <- xlsx_list_files(iris_path)$name
  ref <- c("_rels/.rels", "[Content_Types].xml", "docProps/app.xml",
           "docProps/core.xml", "xl/_rels/workbook.xml.rels",
           "xl/sharedStrings.xml", "xl/styles.xml", "xl/theme/theme1.xml",
           "xl/workbook.xml", "xl/worksheets/sheet1.xml")
  expect_identical(intersect(iris_files, ref), iris_files)
})

