col_apply_tint <- function(col, tint) {
  if (length(tint) == 1L && length(col) > 1L) {
    tint <- rep(tint, length(col))
  }
  i <- tint < 0
  hsl <- rgb2hsl(col2rgb(col))
  if (any(i)) {
    hsl[3L, i] <- hsl[3L, i] * (1 + tint)
  }
  i <- !i
  if (any(i)) {
    hsl[3L, i] <- hsl[3L, i] * (1 - tint) + tint
  }
  rgb2col(hsl2rgb(hsl))
}

## NOTE: the spec is unfortunately a little vague about the
## interpretation of the alpha channel; in the example colours
## (p. 1763) they use 00 to indicate opacity but empirically (and
## conventionally) FF is used.
argb2rgb <- function(x) {
  a <- substr(x, 1L, 2L)
  rgb <- paste0("#", substr(x, 3L, 8L))
  if (a == "FF") rgb else paste0(rgb, a)
}

check_col_matrix <- function(m) {
  if (!is.matrix(m)) {
    if (length(m) == 3L) {
      m <- matrix(m, 3L, 1L)
    } else if (length(m) == 4L) {
      m <- matrix(m, 3L, 1L)
    } else {
      stop("Invalid input for m")
    }
  }
  if (!any(nrow(m) == c(3, 4))) {
    stop("Invalid input for m")
  }
  m
}

rgb2hsl <- function(m) {
  m <- check_col_matrix(m)
  if (length(m) == 0L) {
    nms <- c("h", "s", "l", if (nrow(m) == 4L) "a")
    ret <- matrix(numeric(), length(nms), 0)
    rownames(ret) <- nms
    return(ret)
  }
  m <- m / 255
  r <- apply(m[1:3, , drop=FALSE], 2, range)
  c_min <- r[1L, ]
  c_max <- r[2L, ]
  delta <- c_max - c_min

  l <- (c_max + c_min) / 2
  s <- delta / (1 - abs(2 * l - 1))

  r <- m[1L, ]
  g <- m[2L, ]
  b <- m[3L, ]

  i <- apply(m[1:3, , drop=FALSE], 2, which.max)
  h <- numeric(length(s))
  j <- i == 1L
  h[j] <- (g[j] - b[j]) / delta[j] %% 6
  j <- i == 2L
  h[j] <- (b[j] - r[j]) / delta[j] + 2
  j <- i == 3L
  h[j] <- (r[j] - g[j]) / delta[j] + 4

  h <- h / 6

  i <- delta == 0
  h[i] <- s[i] <- 0

  i <- h < 0
  h[i] <- h[i] %% 1

  if (nrow(m) == 4L) {
    rbind(h, s, l, a=m[4L, ])
  } else {
    rbind(h, s, l)
  }
}

hsl2rgb <- function(m) {
  m <- check_col_matrix(m)

  h <- m[1L, ] * 360 / 60
  s <- m[2L, ]
  l <- m[3L, ]

  C <- (1 - abs(2 * l - 1)) * s
  X <- C * (1 - abs(h %% 2 - 1))
  H <- floor(h)

  cx <- rbind(C, X)
  rgb <- array(0, dim(m))
  rgb[1:2,     H == 0] <- cx[,    H == 0]
  rgb[1:2,     H == 1] <- cx[2:1, H == 1]
  rgb[2:3,     H == 2] <- cx[,    H == 2]
  rgb[2:3,     H == 3] <- cx[2:1, H == 3]
  rgb[c(1, 3), H == 4] <- cx[2:1, H == 4]
  rgb[c(1, 3), H == 5] <- cx[,    H == 5]

  mm <- l - C / 2
  ret <- (rgb + rep(mm, each=3)) * 255

  rownames(ret) <- c("red", "green", "blue")
  if (nrow(m) == 4L) {
    ret <- rbind(ret, alpha=m[4L, ] * 255)
  }
  ret
}

rgb2col <- function(m) {
  rgb(m[1, ], m[2, ], m[3, ], if (nrow(m) == 4L) m[4, ], maxColorValue=255)
}
