## Test reading against readxl
context("readxl")

## TODO: I don't see the inline string version opening correctly in
## numbers so this might be a little beyond our needs to open here.
test_that("agree with readxl", {
  files <- dir(file.path(get_readxl(), "tests/testthat"),
               "\\.xlsx$", full.names=TRUE)
  for (f in files) {
    if (grepl("^(inlineStr)", basename(f))) {
      next
    }
    ## The as.data.frame here works around something deeply weird with
    ## all.equal and tbl_dfs
    cmp <- as.data.frame(readxl::read_excel(f), stringsAsFactors=FALSE)
    dat <- as.data.frame(rexcel_readxl(f))
    expect_equal(dat, cmp)
  }
})
