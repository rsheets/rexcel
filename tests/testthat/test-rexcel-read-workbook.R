## Can we read some universe of sheets without error?
## Later we could store reference objects, but just seems frustrating now.
context("read workbooks")

sheets <- dir(system.file("sheets", package = "rexcel"),
              pattern = "\\.xlsx$", full.names = TRUE)
sheets <- setNames(sheets, basename(sheets))

for (sh in sheets) {
  test_that(basename(sh), {
    expect_silent(rexcel_read_workbook(sh))
  })
}
