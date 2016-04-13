xlsx_read_style <- function(path) {
  xml <- xlsx_read_file(path, "xl/styles.xml")
  ns <- xml2::xml_ns(xml)

  theme <- xlsx_read_theme(path)
  index <- xlsx_indexed_cols()

  fonts <- xlsx_read_style_fonts(xml, ns, theme, index)
  fills <- xlsx_read_style_fills(xml, ns, theme, index)
  borders <- xlsx_read_style_borders(xml, ns)

  ## XFS is "cell formatting".  The s="<int>" tag refers to an entry
  ## in cellXfs, so this is _probably_ the most useful.
  cell_style_xfs <- xlsx_read_cell_style_xfs(xml, ns)
  cell_xfs <- xlsx_read_cell_xfs(xml, ns)
  cell_styles <- xlsx_read_cell_styles(xml, ns)
  num_formats <- xlsx_read_num_formats(xml, ns)

  ## The cell_styles and cell_style_xfs go together; the cell_styes
  ## table is what other things hit I think, and these translate
  ## through to various formattings in cell_style_xfs.  So we could
  ## probably join those straightaway here.  I don't have a sheet that
  ## uses a style yet.

  list(fonts=fonts, fills=fills, borders=borders,
       cell_style_xfs=cell_style_xfs,
       cell_xfs=cell_xfs,
       cell_styles=cell_styles,
       num_formats=num_formats)
}

xlsx_read_theme <- function(path) {
  ## TODO: Strictly, the theme information should come from the
  ## workbook.rels.xml file by looking to see which file has the
  ## appropriate officeDocument/2006/relationships/theme entry, but this
  ## should be fine for now.
  ##
  ## NOTE: MSDN suggests that this will always be theme1.xml for Excel
  ## and only n>1 for PowerPoint.
  xml <- xlsx_read_file(path, "xl/theme/theme1.xml")
  ns <- xml2::xml_ns(xml)
  tmp <- xml2::xml_find_one(xml, "/a:theme/a:themeElements/a:clrScheme", ns)

  ## Empirical ordering, based on one random website.  I have not
  ## found the support for this in the actual spec yet and have seen a
  ## few variants on the ordering listed there incl dk1/lt1/dk2/lt2/accent...

  nms <- c("lt1", "dk1", "lt2", "dk2",
           paste0("accent", 1:6),
           "hlink", "folHlink")
  f <- function(x, xml, ns) {
    tmp <- xml2::xml_find_one(xml, paste0(".//a:", x), ns)
    nd <- xml2::xml_children(tmp)[[1L]]
    at <- switch(xml2::xml_name(nd), sysClr="lastClr", srgbClr="val")
    paste0("#", xml2::xml_attr(nd, at, ns))
  }
  pal <- vcapply(nms, f, xml, ns)

  list(palette=pal)
}

## TODO: expand this out to include all attributes that might be
## present in the XML based on the schema rather than using just what
## is here.
xlsx_read_style_fonts <- function(xml, ns, theme, index) {
  fonts <- xml2::xml_children(xml2::xml_find_one(xml, "d1:fonts", ns))
  dat <- lapply(fonts, xlsx_read_style_font, ns, theme, index)
  do.call("rbind", dat, quote=TRUE)
}

## Getting the definition of this from the spec is proving difficult:
##
## On p. 1759, 18.8.22 font just says "CT_Font is in A.2"
##
## The link is broken, but p. 3930, l 3797 looks good.  Beware of the
## similar but different CT_Font probably for Word's XML.
##
## Possible tags (all optional but at most one of each present)
##
##   name (CT_FontName)
##   charset (CT_IntProperty)
##   family (CT_FontFamily)
##   b, i, strike, outline, shadow, condense, extend (CT_BooleanProperty)
##   color (CT_Color)
##   sz (CT_FontSize)
##   u (CT_UnderlineProperty)
##   vertAlign (CT_VerticalAlignFontProperty) - not actually vertical alignment
##   scheme (CT_FontScheme)
##
## Looks like horizontal alignment comes through with the xf element
## in cellxfs, but I think I ignore that at the moment.  Seems like an
## odd place tbh.
xlsx_read_style_font <- function(x, ns, theme, index) {
  name <- xml2::xml_text(xml2::xml_find_one(x, "d1:name/@val", ns))
  ## ignoring charset
  family <- xlsx_st_font_family(xml2::xml_find_one(x, "d1:family", ns))

  b <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:b", ns))
  i <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:i", ns))
  strike <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:strike", ns))
  outline <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:outline", ns))
  shadow <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:shadow", ns))
  condense <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:condense", ns))
  extend <- xlsx_ct_boolean_property(xml2::xml_find_one(x, "d1:extend", ns))

  color <- xlsx_ct_color(xml2::xml_find_one(x, "d1:color", ns), theme, index)
  sz <- as.integer(xml2::xml_text(xml2::xml_find_one(x, "d1:sz/@val", ns)))

  u <- xlsx_ct_underline_property(xml2::xml_find_one(x, "d1:u", ns))
  ## This one here is either baseline, superscript or subscript.  So probably not terribly useful and fairly confuse-able with _actual_ vertical alignment.

  ## vertAlign <- xml2::xml_text(xml2::xml_find_one(x, "d1:vertAlign/@val", ns))
  scheme <- xml2::xml_text(xml2::xml_find_one(x, "d1:scheme/@val", ns))

  tibble::data_frame(name, family,
                     b, i, strike, outline, shadow, condense, extend,
                     color, sz, u, scheme)

}

xlsx_st_font_family <- function(f, missing=NA_character_) {
  pos <- c(NA_character_, "Roman", "Swiss", "Modern", "Script", "Decorative",
           rep("<<reserved>>", 9))
  if (inherits(f, "xml_missing")) {
    missing
  } else {
    pos[[as.integer(xml2::xml_attr(f, "val")) + 1L]]
  }
}

xlsx_ct_boolean_property <- function(b, missing=FALSE) {
  if (inherits(b, "xml_missing")) {
    missing
  } else {
    val <- xml2::xml_attr(b, "val")
    if (is.na(val)) TRUE else as.logical(as.integer(val))
  }
}

xlsx_ct_underline_property <- function(u, missing="none") {
  if (inherits(u, "xml_missing")) {
    missing
  } else {
    val <- xml2::xml_attr(u, "val")
    if (is.na(val)) "single" else val
  }
}

xlsx_read_style_fills <- function(xml, ns, theme, index) {
  fills <- xml2::xml_children(xml2::xml_find_one(xml, "d1:fills", ns))
  dat <- lapply(fills, xlsx_read_style_fill, ns, theme, index)
  ## TODO: In the case where not all of these are "pattern" (i.e., we
  ## have a gradient fill) this will not work correctly because we
  ## need totally different things here.  I think what we'll return
  ## there is type=gradient, and then a lookup to a gradient table, so
  ## this will expand by one more column with gradient_id perhaps.
  as.data.frame(do.call("rbind", dat), stringsAsFactors=FALSE)
}

xlsx_read_style_fill <- function(x, ns, theme, index) {
  ## The only options here, according to the xsd (A.2, p. 3925,
  ## l. 3498) is a single element of patternFill or gradientFill
  xk <- xml2::xml_children(x)[[1L]]
  if (xml2::xml_name(xk) == "patternFill") {
    xlsx_read_style_pattern_fill(xk, ns, theme, index)
  } else {
    xlsx_read_style_gradient_fill(xk, ns, theme, index)
  }
}

xlsx_read_style_pattern_fill <- function(x, ns, theme, index) {
  ## This is very weird because all of the attribute patternType,
  ## fgColor and bgColor are optional.
  pattern_type <- xml2::xml_attr(x, "patternType")
  fg <- xlsx_ct_color(xml2::xml_find_one(x, "./d1:fgColor", ns), theme, index)
  bg <- xlsx_ct_color(xml2::xml_find_one(x, "./d1:bgColor", ns), theme, index)
  c(type="pattern", pattern_type=pattern_type, fg=fg, bg=bg)
}

xlsx_read_style_gradient_fill <- function(x, ns, theme, index) {
  ## zero or more stop elements, plus attributes type, degree, left,
  ## right, bottom, all of which are optional.  I think that
  ## realistically we'll have to dump these into a separate lookup
  ## table or something.
  ##
  ## It will be interesting to see what is used in the main corpus.
  ## Even with the terrible things that people do to spreadsheets I'd
  ## hope that this is not actually common.
  stop("Ignoring gradient fill")
}

## See 18.8.19, p. 1757
xlsx_ct_color <- function(x, theme, index) {
  if (inherits(x, "xml_missing")) {
    NA_character_
  } else {
    ## The schema is vague on this point but let's make the assumption
    ## that only one of the following is present:
    ## auto, indexed, rgb, theme
    tmp <- xml2::xml_attrs(x)
    types <- c("auto", "indexed", "rgb", "theme")
    i <- types %in% names(tmp)
    if (!any(i)) {
      return(NA_character_)
    }
    t <- types[i][[1L]]
    v <- tmp[[t]]
    col <- switch(
      t,
      auto=stop("(TODO) I don't actually know what 'auto' means here..."),
      indexed=index[[as.integer(v) + 1L]],
      rgb=argb2rgb(v),
      theme=theme$palette[[as.integer(v) + 1L]])
    if ("tint" %in% names(tmp)) {
      col <- col_apply_tint(col, as.numeric(tmp[["tint"]]))
    }
    col
  }
}

xlsx_read_style_borders <- function(xml, ns) {
  borders <- xml2::xml_find_one(xml, "d1:borders", ns)
  borders_f <- xml2::xml_children(borders)

  pos <- sort(unique(unlist(lapply(borders_f, function(el)
    xml2::xml_name(kids <- xml2::xml_children(el))))))
  known <- c("bottom", "left", "top", "right", "diagonal")
  unk <- setdiff(pos, known)
  if (length(unk) > 0L) {
    message("Skipping unhandled border tags: ", paste(unk, collapse=", "))
  }

  ## NOTE: Not taking any actual style information here (e.g., weight,
  ## colour), aside from the presence of the border.  It's possible
  ## that this does not get the correct style in all cases though --
  ## if style == "none" perhaps this is wrong?  Is that a valid style?
  data.frame(
    bottom=xml2::xml_find_lgl(borders_f, "boolean(d1:bottom/@style)", ns),
    left=xml2::xml_find_lgl(borders_f, "boolean(d1:left/@style)", ns),
    top=xml2::xml_find_lgl(borders_f, "boolean(d1:top/@style)", ns),
    right=xml2::xml_find_lgl(borders_f, "boolean(d1:right/@style)", ns),
    diagonal=xml2::xml_find_lgl(borders_f, "boolean(d1:diagonal/@style)", ns),
    stringsAsFactors=FALSE)
}

xlsx_read_cell_style_xfs <- function(xml, ns) {
  csx <- xml2::xml_find_one(xml, "d1:cellStyleXfs", ns)
  as.data.frame(attrs_to_matrix(xml2::xml_children(csx), "integer"))
}

xlsx_read_cell_xfs <- function(xml, ns) {
  cx <- xml2::xml_find_one(xml, "d1:cellXfs", ns)
  cx_kids <- xml2::xml_children(cx)
  ret <- as.data.frame(attrs_to_matrix(cx_kids, "integer"))
  ret_align <- attrs_to_matrix(xml2::xml_find_one(cx_kids, "d1:alignment", ns))
  ret_align <- data.frame(ret_align, stringsAsFactors=FALSE)
  if ("wrapText" %in% names(ret_align)) {
    ret_align$wrapText <- as.logical(ret_align$wrapText)
  }
  cbind(ret, ret_align)
}

xlsx_read_cell_styles <- function(xml, ns) {
  cs <- xml2::xml_find_one(xml, "d1:cellStyles", ns)
  ret <- data.frame(attrs_to_matrix(xml2::xml_children(cs)),
                    stringsAsFactors=FALSE)
  if ("xfId" %in% names(ret)) {
    ret$xfId <- as.integer(ret$xfId)
  }
  if ("builtinId" %in% names(ret)) {
    ret$builtinId <- as.integer(ret$builtinId)
  }
  if ("hidden" %in% names(ret)) {
    ret$hidden <- as.logical(ret$hidden)
  }
  ret
}

xlsx_read_num_formats <- function(xml, ns) {
  dat <- xml2::xml_find_one(xml, "d1:numFmts", ns)
  ret <- as.data.frame(attrs_to_matrix(xml2::xml_children(dat)))
  if ("numFmtId" %in% names(ret)) {
    ret$numFmtId <- as.integer(ret$numFmtId)
  }
  ret
}

col_apply_tint <- function(col, tint) {
  if (length(tint) == 1L && length(col) > 1L) {
    tint <- rep(tint, length(col))
  }
  i <- tint < 0
  hsv <- col2hsv(col)
  if (any(i)) {
    hsv[3L, i] <- hsv[3L, i] * (1 + tint)
  }
  if (!all(i)) {
    j <- !i
    hsv[3L, j] <- hsv[3L, j] * (1 - tint) + tint
  }
  hsv2rgb(hsv)
}

## These come from the ECMA Open XML definition, p 1763 (18.8.27).
## The spec describes this as a "legacy indexing scheme for colors
## that is still required for some records, and for backwards
## compatibility with legacy formats" but this seems to be far more
## widespread than that (and from things generated with Microsoft's
## current software I think).
##
## Indecies 64 and 65 (the 65th and 66th elements) should be treated
## specially as system foreground and background colour respectively, but
xlsx_indexed_cols <- function() {
  c("#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF",
    "#FFFF00", "#FF00FF", "#00FFFF", "#000000", "#FFFFFF",
    "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
    "#00FFFF", "#800000", "#008000", "#000080", "#808000",
    "#800080", "#008080", "#C0C0C0", "#808080", "#9999FF",
    "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080",
    "#0066CC", "#CCCCFF", "#000080", "#FF00FF", "#FFFF00",
    "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF",
    "#FF99CC", "#CC99FF", "#FFCC99", "#3366FF", "#33CCCC",
    "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699",
    "#969696", "#003366", "#339966", "#003300", "#333300",
    "#993300", "#993366", "#333399", "#333333",
    ## Special:
    "black", "white")
}

## See 18.8.30, p. 1767
xlsx_format_codes <- function() {
  ## "Ids not specified in the listing, such as 5, 6, 7, and 8, shall
  ## follow the number format specified by the formatCode attribute."
  c("General",
    "0",
    "0.00",
    "#,##0",
    "#,##0.00",
    ## missing 4-8 incl
    rep(NA, length(4:8)),
    "0%",
    "0.00%",
    "0.00E+00",
    "# ?/?",
    "# ??/??",
    "mm-dd-yy",
    "d-mmm-yy",
    "d-mmm mmm-yy",
    "h:mm AM/PM",
    "h:mm:ss AM/PM",
    "h:mm",
    "h:mm:ss",
    "m/d/yy h:mm",
    ## missing 23-36 incl
    rep(NA, length(23:36)),
    "#,##0 ;(#,##0)",
    "#,##0 ;[Red](#,##0)",
    "#,##0.00;(#,##0.00)",
    "#,##0.00;[Red](#,##0.00)",
    ## missing 41-44 incl
    rep(NA, length(41:44)),
    "mm:ss",
    "[h]:mm:ss",
    "mmss.0",
    "##0.0E+0",
    "@")
}

## See 18.18.55, p. 2462
xlsx_pattern_type <- function() {
  c(## Can process these two
    "none",  # ignores both fgColor and bgColor
    "solid", # renders only the fgColor
    ## but not these:
    "darkDown",
    "darkGray",
    "darkGrid",
    "darkHorizontal",
    "darkTrellis",
    "darkUp",
    "darkVertical",
    "gray0625",
    "gray125",
    "lightDown",
    "lightGray",
    "lightGrid",
    "lightHorizontal",
    "lightTrellis",
    "lightUp",
    "lightVertical",
    "mediumGray")
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

col2hsv <- function(col) {
  rgb2hsv(col2rgb(col))
}

hsv2rgb <- function(m) {
  rgb(m[1, ], m[2, ], m[3, ], if (nrow(m) == 4L) m[4, ])
}
