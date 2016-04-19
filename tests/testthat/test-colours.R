context("colours")

test_that("rgb2hsl", {
  ## These are random colours, ground-truthed with
  ##   http://www.rapidtables.com/convert/color/rgb-to-hsl.htm
  ## Because that website returns things as [degrees, percent,
  ## percent] to 3sf I have a little wrapper here.
  f <- function(x) {
    round(x * c(360, 100, 100), c(0, 1, 1))
  }

  expect_identical(f(rgb2hsl(c(106, 90, 205))), rbind(h=248, s=53.5, l=57.8))

  cols <- c("gray22", "gray92", "indianred4", "slateblue",
            "dodgerblue4", "sienna2", "steelblue", "lightskyblue",
            "lightcyan4", "burlywood4")
  ans <- f(rgb2hsl(col2rgb(cols)))
  cmp <- rbind(h=c(0, 0, 0, 248, 210, 19, 207, 203, 180, 33),
               s=c(0, 0, 41.1, 53.5, 79.4, 83.5, 44, 92, 6.8, 24.1),
               l=c(22, 92.2, 38.6, 57.8, 30.4, 59.6, 49, 75.5, 51.2, 43.9))
  expect_identical(ans, cmp)

  ## Alpha handling:
  m <- rbind(col2rgb(cols), alpha=runif(length(cols), 0, 255))
  ans2 <- rgb2hsl(m)
  expect_identical(ans2[1:3, ], rgb2hsl(col2rgb(cols)))
  expect_equal(ans2[4, ], m[4, ] / 255)

  expect_equal(dim(rgb2hsl(matrix(numeric(), 3, 0))), c(3, 0))
  expect_equal(dim(rgb2hsl(matrix(numeric(), 4, 0))), c(4, 0))

  ## And the reverse
  m <- col2rgb(cols)
  expect_equal(hsl2rgb(rgb2hsl(m)), m)

  ## Do them _all_
  mm <- col2rgb(colors())
  expect_equal(hsl2rgb(rgb2hsl(mm)), mm)
})
