## Basically everything here follows from "Styles": section 18.8 (p. 1744).
##
## I have been fairly thorough at pulling things in, but it's not
## really complete.  There are enumeration types that I have not
## turned into factors (arguments go either way, really).  There are a
## few things that are going to be very hard to deal with, too.
##
##   [ ] 18.8.1  alignment (Alignment)
##   [ ] 18.8.2  b (Bold)
##   [ ] 18.8.3  bgColor (Background Color)
##   [ ] 18.8.4  border (Border)
##   [ ] 18.8.5  borders (Borders)
##   [ ] 18.8.6  bottom (Bottom Border)
##   [x] 18.8.7  cellStyle (Cell Style) -- xlsx_ct_cell_style
##   [x] 18.8.8  cellStyles (Cell Styles) -- xlsx_ct_cell_styles
##   [ ] 18.8.9  cellStyleXfs (Formatting Records)
##   [ ] 18.8.10 cellXfs (Cell Formats)
##   [ ] 18.8.11 colors (Colors)
##   [ ] 18.8.12 condense (Condense)
##   [ ] 18.8.13 diagonal (Diagonal)
##   [ ] 18.8.14 dxf (Formatting)
##   [ ] 18.8.15 dxfs (Formats)
##   [ ] 18.8.16 end (Trailing Edge Border)
##   [ ] 18.8.17 extend (Extend)
##   [ ] 18.8.18 family (Font Family)
##   [ ] 18.8.19 fgColor (Foreground Color)
##   [ ] 18.8.20 fill (Fill)
##   [ ] 18.8.21 fills (Fills)
##   [x] 18.8.22 font (Font) -- xlsx_ct_font
##   [x] 18.8.23 fonts (Fonts) -- xlsx_ct_fonts
##   [ ] 18.8.24 gradientFill (Gradient)
##   [ ] 18.8.25 horizontal (Horizontal Inner Borders)
##   [ ] 18.8.26 i (Italic)
##   [ ] 18.8.27 indexedColors (Color Indexes)
##   [ ] 18.8.28 mruColors (MRU Colors)
##   [ ] 18.8.29 name (Font Name)
##   [x] 18.8.30 numFmt (Number Format) -- xlsx_ct_num_fmt
##   [x] 18.8.31 numFmts (Number Formats) -- xlsx_ct_num_fmts
##   [ ] 18.8.32 patternFill (Pattern)
##   [ ] 18.8.33 protection (Protection Properties)
##   [ ] 18.8.34 rgbColor (RGB Color)
##   [ ] 18.8.35 scheme (Scheme)
##   [ ] 18.8.36 shadow (Shadow)
##   [ ] 18.8.37 start (Leading Edge Border)
##   [ ] 18.8.38 stop (Gradient Stop)
##   [ ] 18.8.39 styleSheet (Style Sheet)
##   [ ] 18.8.40 tableStyle (Table Style)
##   [ ] 18.8.41 tableStyleElement (Table Style)
##   [ ] 18.8.42 tableStyles (Table Styles)
##   [ ] 18.8.43 top (Top Border)
##   [ ] 18.8.44 vertical (Vertical Inner Border)
##   [ ] 18.8.45 xf (Format)
##
## I don't have a good naming scheme for all of these yet because I
## don't really have a sense for how the open XML authors named
## things.  There are elements and at some point things get called a
## CT_<thing> (for "Complex Type"), e.g., CT_Color.  At that point
## things are called xlsx_ct_boolean_property or similar.
##
## There are no strong argument conventions either; if anything can
## contain a colour the both theme and index are passed through, as
## these contain information to convert colour types into an RGB
## triplet.  If any Xpath query is used then the namespace is passed
## along as ns.
##
## From the point of view of the rest of the package, the only entry
## point to use is xlsx_read_style which will return a list of a great
## many data.frames (the format here will get cleaned up soon).

xlsx_read_style <- function(path) {
  xml <- xlsx_read_file(path, "xl/styles.xml")
  ns <- xml2::xml_ns(xml)

  theme <- xlsx_read_theme(path)
  index <- xlsx_indexed_cols()

  fonts <- xlsx_ct_fonts(xml, ns, theme, index)
  fills <- xlsx_read_style_fills(xml, ns, theme, index)
  borders <- xlsx_read_style_borders(xml, ns, theme, index)

  cell_style_xfs <- xlsx_read_style_cell_style_xfs(xml, ns)
  cell_xfs <- xlsx_read_style_cell_xfs(xml, ns)
  cell_styles <- xlsx_ct_cell_styles(xml, ns)
  num_fmts <- xlsx_ct_num_fmts(xml, ns)

  list(fonts=fonts,
       fills=fills,
       borders=borders,
       cell_style_xfs=cell_style_xfs,
       cell_xfs=cell_xfs,
       cell_styles=cell_styles,
       num_fmts=num_fmts)
}

## NOTE: this only reads the the colour information from the theme as
## nothing else looks that exciting in there, really.
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

xlsx_ct_fonts <- function(xml, ns, theme, index) {
  process_container(xml, "d1:fonts", ns, xlsx_ct_font, theme, index)
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
xlsx_ct_font <- function(x, ns, theme, index) {
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
    ## TODO: I can't find any information indicating what "auto" means
    ## in this context.  The spec says (at least for fgColor in
    ## 18.8.19, p. 1757, but similar words are used elsewhere):
    ##
    ## > auto: A boolean value indicating the color is automatic and
    ## > system color dependent.
    ##
    ## So it probably depends on exactly _where_ the colour is used
    ## (e.g. if it tends to be a fg or a bg colour).  So I will return
    ## "auto" I think, at least for now.  Probably I could return
    ## "black" but that's going to be quite lossy.  This way I can
    ## transform into a sensible colour at use.
    col <- switch(
      t,
      auto="auto",
      indexed=index[[as.integer(v) + 1L]],
      rgb=argb2rgb(v),
      theme=theme$palette[[as.integer(v) + 1L]])
    if ("tint" %in% names(tmp)) {
      col <- col_apply_tint(col, as.numeric(tmp[["tint"]]))
    }
    col
  }
}

xlsx_read_style_borders <- function(xml, ns, theme, index) {
  borders <- xml2::xml_children(xml2::xml_find_one(xml, "d1:borders", ns))
  dat <- lapply(borders, xlsx_read_style_border, ns, theme, index)
  do.call("rbind", dat, quote=TRUE)
}

## See
##   * 18.8.4 (p. 1749)
##   * 18.8.5 (p. 1750)
##   * A.2 l. 3460 (p. 3924)
##
## Unfortunately, note that the xsd talks about start / end but the
## *example* has begin / end.  And neither of them indicates what on
## earth these are for (though the text in the example suggests that
## end is the right border in that context).  In the sheets I am
## looking at I mostly see left / right / top / bottom / diagonal.
xlsx_read_style_border <- function(x, ns, theme, index) {
  ## NOTE: I am skipping attributes diagonalUp and diagonalDown along
  ## with the element diagonal - it's not the only bit of formatting
  ## trivia we won't handle, but it's a fairly unusual thing to see, I
  ## believe.
  outline <- attr_bool(xml2::xml_attr(x, "outline"))

  f <- function(path) {
    xlsx_ct_border_pr(xml2::xml_find_one(x, path, ns), ns, theme, index)
  }

  tmp <- list(list(outline = outline),
              start = f("d1:start"),
              end = f("d1:end"),
              left = f("d1:left"),
              right = f("d1:right"),
              top = f("d1:top"),
              bottom = f("d1:bottom"))
  tmp <- unlist(tmp, FALSE)
  names(tmp) <- sub(".", "_", names(tmp), fixed=TRUE)
  tibble::as_data_frame(tmp)
}

## style (ST_BorderStyle) can be one of (18.18.3, p. 2428):
##
##   * dashDot
##   * dashDotDot
##   * dashed
##   * dotted
##   * double
##   * hair
##   * medium
##   * mediumDashDot
##   * mediumDashDotDot
##   * mediumDashed
##   * none
##   * slantDashDot
##   * thick
##   * thin
##
## Note that the various combinations do not cross with one another.
xlsx_ct_border_pr <- function(x, ns, theme, index) {
  present <- !inherits(x, "xml_missing")
  if (present) {
    style <- xml2::xml_attr(x, "style")
    color <- xlsx_ct_color(xml2::xml_find_one(x, "d1:color", ns), theme, index)
  } else {
    color <- style <- NA_character_
  }
  list(present=present, style=style, color=color)
}

xlsx_read_style_cell_style_xfs <- function(xml, ns) {
  xfs <- xml2::xml_children(xml2::xml_find_one(xml, "d1:cellStyleXfs", ns))
  dat <- lapply(xfs, xlsx_read_style_xf, ns)
  tibble::as_data_frame(do.call("rbind", dat, quote=TRUE))
}

xlsx_read_style_cell_xfs <- function(xml, ns) {
  xfs <- xml2::xml_children(xml2::xml_find_one(xml, "d1:cellXfs", ns))
  dat <- lapply(xfs, xlsx_read_style_xf, ns)
  tibble::as_data_frame(do.call("rbind", dat, quote=TRUE))
}

xlsx_read_style_xf <- function(x, ns) {
  at <- as.list(xml2::xml_attrs(x))
  xf <- tibble::data_frame(
    ## Booleans, indicating if things are applied:
    apply_alignment = attr_bool(at$applyAlignment),
    apply_border = attr_bool(at$applyBorder),
    apply_fill = attr_bool(at$applyFill),
    apply_font = attr_bool(at$applyFont),
    apply_number_format = attr_bool(at$applyNumberFormat),
    apply_protection = attr_bool(at$applyProtection),

    ## References to actual formats (all base 0)
    border_id = attr_integer(at$borderId),
    fill_id = attr_integer(at$fillId),
    font_id = attr_integer(at$fontId),
    num_fmt_id = attr_integer(at$numFmtId),

    pivot_button = attr_bool(at$pivotButton),
    quote_prefix = attr_bool(at$quotePrefix),

    ## This is a reference against cellStyleXfs
    xf_id = attr_integer(at$xfId))
  alignment <- xlsx_read_style_alignment(
    xml2::xml_find_one(x, "d1:alignment", ns))
  cbind(xf, alignment)
}

## horizontal: center | centerContinuous | distributed | fill |
##   general | justify | right
##
## vertical: bottom | center | distributed | justify | top
xlsx_read_style_alignment <- function(x, ns) {
  at <- as.list(xml2::xml_attrs(x))
  tibble::data_frame(
    horizontal=attr_character(at$horizontal),
    vertical=attr_character(at$vertical),
    indent=attr_integer(at$indent),
    justify_last_line=attr_bool(at$justifyLastLine),
    reading_order=attr_integer(at$readingOrder),
    ## relativeIndent [used only in a dxf element]
    shrink_to_fit=attr_bool(at$shrinkToFit),
    text_rotation=attr_integer(at$text_rotation),
    text_wrap=attr_bool(at$textWrap))
}

xlsx_ct_cell_styles <- function(xml, ns) {
  cs <- xml2::xml_children(xml2::xml_find_one(xml, "d1:cellStyles", ns))
  dat <- lapply(cs, xlsx_ct_cell_style, ns)
  tibble::as_data_frame(do.call("rbind", dat, quote=TRUE))
}

xlsx_ct_cell_style <- function(x, ns) {
  ## NOTE: Getting this right is really hard because the Annex (G.2)
  ## lists information about "built-in" styles but these vary with all
  ## things like row position, but no actual information about the
  ## styles is given in the annex.  So it's not really obvious what we
  ## can do here.

  ## NOTE: This element can contain "extension list" elements which
  ## are reserved for future use.  But we can skip that.

  ## NOTE: xf_id: Zero-based index referencing an xf record in the
  ## cellStyleXfs collection. This is used to determine the formatting
  ## defined for this named cell style.

  at <- as.list(xml2::xml_attrs(x))
  tibble::data_frame(
    builtin_id = attr_integer(at$builtinId),
    custom_builtin = attr_bool(at$customBuiltin),
    hidden = attr_bool(at$hidden),
    i_level = attr_integer(at$iLevel),
    name = attr_character(at$name),
    xf_id = attr_integer(at$xfId))
}

xlsx_ct_num_fmts <- function(xml, ns) {
  process_container(xml, "d1:numFmts", ns, xlsx_ct_num_fmt)
}

xlsx_ct_num_fmt <- function(x, ns) {
  at <- as.list(xml2::xml_attrs(x))
  tibble::data_frame(
    num_format_id = attr_integer(at$numFmtId),
    format_code = attr_character(at$formatCode))
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
