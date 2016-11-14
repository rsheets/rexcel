context("view")

## NOTE: These are really tests of *linen*

test_that("basic view", {
  filename <- here("inst", "sheets", "mini-gap.xlsx")
  d <- rexcel::rexcel_read(filename)
  tbl <- d$table()
  expect_equal(nrow(tbl), 5)
  expect_equal(names(tbl)[[1]], "country")

  xr <- cellranger::cell_limits(c(4, 1), c(d$dim[[1]], d$dim[[2]]))
  v <- linen::worksheet_view(d, xr)
  expect_error(v$table(), "header information not convertable to col_names")

  v2 <- linen::worksheet_view(d, xr, header = letters[seq_len(d$dim[[1]])])
  tbl_v2 <- v2$table()
  expect_equal(nrow(tbl_v2), 3)
  expect_equal(names(tbl_v2), letters[1:6])

  tmp <- unname(tbl[3:5, ])
  rownames(tmp) <- NULL
  expect_equal(unname(tbl_v2), tmp)
})
