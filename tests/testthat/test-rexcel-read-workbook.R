## Can we read some universe of sheets without error?
## Later we could store reference objects, but just seems frustrating now.
context("read workbooks")

sheets <- dir(system.file("sheets", package = "rexcel"),
              pattern = "\\.xlsx$", full.names = TRUE)
sheets <- setNames(sheets, basename(sheets))

## Jenny: I propose we skip Ekaterinburg until we have a decent, general
## solution to the non-standard namespacing problem
## Jenny: I'm skipping both because even resaved one is large. In due course,
## we should create a scaled down version for testing.
## Jenny: also skipping exp.xlsx for now
sheets <-
  sheets[grep("^Ekaterinburg_IP_9|^exp", names(sheets), invert = TRUE)]

for (sh in sheets) {
  test_that(basename(sh), {
    expect_silent(rexcel_read_workbook(sh))
  })
}

test_that("read one sheet - by name", {
  filename <- sheets[["mini-gap.xlsx"]]
  d <- rexcel_read_workbook(filename)
  for (s in d$names) {
    expect_equal(d$sheets[[s]]$cells, rexcel_read(filename, s)$cells)
  }
})

test_that("read one sheet - by index", {
  filename <- sheets[["mini-gap.xlsx"]]
  d <- rexcel_read_workbook(filename)
  for (s in seq_along(d$names)) {
    expect_equal(d$sheets[[s]]$cells, rexcel_read(filename, s)$cells)
  }
})
