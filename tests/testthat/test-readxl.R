## Test reading against readxl
context("readxl")

files <- dir(file.path(get_readxl(), "tests", "testthat"),
             pattern = "\\.xlsx$", full.names = TRUE)

## Rich says: I don't see the inline string version opening correctly in
## numbers so this might be a little beyond our needs to open here.
files <- files[grep("^(inlineStr)", basename(files), invert = TRUE)]
files <- setNames(files, basename(files))

for (f in files) {
  test_that(basename(f), {
    ## Rich says: the as.data.frame here works around something deeply weird
    ## with all.equal and tbl_dfs
    readxl <- as.data.frame(readxl::read_excel(f), stringsAsFactors = FALSE)
    us <- as.data.frame(rexcel_readxl(f))
    expect_equal(us, readxl, label = paste("our import of", basename(f)))
  })
}
