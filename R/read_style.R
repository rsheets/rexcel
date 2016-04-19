## Basically everything here follows from "Styles": section 18.8 (p. 1744).
##
## I have been fairly thorough at pulling things in, but it's not
## really complete.  There are enumeration types that I have not
## turned into factors (arguments go either way, really).  There are a
## few things that are going to be very hard to deal with, too.
##
##   [x] 18.8.1  alignment (Alignment) -- xlsx_ct_alignment
##   [x] 18.8.2  b (Bold) -- xlsx_ct_boolean_property
##   [x] 18.8.3  bgColor (Background Color) -- xlsx_ct_color
##   [x] 18.8.4  border (Border) -- xlsx_ct_border
##   [x] 18.8.5  borders (Borders) -- xlsx_ct_borders
##   [x] 18.8.6  bottom (Bottom Border) -- xlsx_ct_border_pr
##   [x] 18.8.7  cellStyle (Cell Style) -- xlsx_ct_cell_style
##   [x] 18.8.8  cellStyles (Cell Styles) -- xlsx_ct_cell_styles
##   [x] 18.8.9  cellStyleXfs (Formatting Records) -- xlsx_ct_cell_style_xfs
##   [x] 18.8.10 cellXfs (Cell Formats) -- xlsx_ct_cell_xfs
##   [ ] 18.8.11 colors (Colors)
##   [x] 18.8.12 condense (Condense) -- xlsx_ct_boolean_property
##   [-] 18.8.13 diagonal (Diagonal) -- (xlsx_ct_border_pr)
##   [ ] 18.8.14 dxf (Formatting)
##   [ ] 18.8.15 dxfs (Formats)
##   [x] 18.8.16 end (Trailing Edge Border) -- xlsx_ct_border_pr
##   [x] 18.8.17 extend (Extend) -- xlsx_ct_boolean_property
##   [x] 18.8.18 family (Font Family) -- xlsx_st_font_family
##   [x] 18.8.19 fgColor (Foreground Color) -- xlsx_ct_color
##   [x] 18.8.20 fill (Fill) -- xlsx_ct_fill
##   [x] 18.8.21 fills (Fills) -- xlsx_ct_fills
##   [x] 18.8.22 font (Font) -- xlsx_ct_font
##   [x] 18.8.23 fonts (Fonts) -- xlsx_ct_fonts
##   [x] 18.8.24 gradientFill (Gradient) -- xlsx_ct_gradient_fill
##   [ ] 18.8.25 horizontal (Horizontal Inner Borders)
##   [x] 18.8.26 i (Italic) -- xlsx_ct_boolean_property
##   [ ] 18.8.27 indexedColors (Color Indexes)
##   [ ] 18.8.28 mruColors (MRU Colors)
##   [x] 18.8.29 name (Font Name) -- (plain text handling in xlsx_ct_font)
##   [x] 18.8.30 numFmt (Number Format) -- xlsx_ct_num_fmt
##   [x] 18.8.31 numFmts (Number Formats) -- xlsx_ct_num_fmts
##   [x] 18.8.32 patternFill (Pattern) -- xlsx_ct_pattern_fill
##   [ ] 18.8.33 protection (Protection Properties)
##   [ ] 18.8.34 rgbColor (RGB Color)
##   [x] 18.8.35 scheme (Scheme) -- (plain text handling in xlsx_ct_font)
##   [x] 18.8.36 shadow (Shadow) -- xlsx_ct_boolean_property
##   [x] 18.8.37 start (Leading Edge Border) -- xlsx_ct_border_pr
##   [ ] 18.8.38 stop (Gradient Stop)
##   [x] 18.8.39 styleSheet (Style Sheet) -- xlsx_read_style
##   [ ] 18.8.40 tableStyle (Table Style)
##   [ ] 18.8.41 tableStyleElement (Table Style)
##   [ ] 18.8.42 tableStyles (Table Styles)
##   [x] 18.8.43 top (Top Border) -- xlsx_ct_border_pr
##   [ ] 18.8.44 vertical (Vertical Inner Border)
##   [x] 18.8.45 xf (Format) -- xlsx_ct_xf
##
## Most elements at some point things get called a CT_<thing> (for
## "Complex Type"), e.g., CT_Color; at that point the processing thing
## is called xlsx_ct_boolean_property or similar.  Note that this
## drives off the *type*, not off the element name.  These are often,
## but not always, the same.
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
##
## Some of the needed functions here come from the shared string table.
##
##
xlsx_read_style <- function(path) {
  xml <- xlsx_read_file(path, "xl/styles.xml")
  ns <- xml2::xml_ns(xml)

  theme <- xlsx_read_theme(path)
  index <- xlsx_indexed_cols()

  fonts <- xlsx_ct_fonts(xml, ns, theme, index)
  fills <- xlsx_ct_fills(xml, ns, theme, index)
  borders <- xlsx_ct_borders(xml, ns, theme, index)

  cell_style_xfs <- xlsx_ct_cell_style_xfs(xml, ns)
  cell_xfs <- xlsx_ct_cell_xfs(xml, ns)
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
  xml <- xlsx_read_file_if_exists(path, "xl/theme/theme1.xml")
  if (is.null(xml)) {
    return(NULL)
  }
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

## 18.8.23 fonts
xlsx_ct_fonts <- function(xml, ns, theme, index) {
  process_container(xml, "d1:fonts", ns, xlsx_ct_font, theme, index)
}

## 18.8.22 font
##
## The link to the actual definition is broken, but p. 3930, l 3797
## looks good.  Beware of the similar but different CT_Font probably
## for Word's XML.
##
## Possible daughter elements (all optional but at most one of each present)
##
##   name (CT_FontName)
##   charset (CT_IntProperty)
##   family (CT_FontFamily)
##   b, i, strike, outline, shadow, condense, extend (CT_BooleanProperty)
##   color (CT_Color)
##   sz (CT_FontSize)
##   u (CT_UnderlineProperty)
##   vertAlign (CT_VerticalAlignFontProperty) - subscript / superscript
##   scheme (CT_FontScheme)
##
## Looks like horizontal alignment comes through with the xf element
## in cellxfs, but I think I ignore that at the moment.  Seems like an
## odd place tbh.
##
## Despite most elements being CT_*, most of this is just that if the
## element is present a "val" attribute is required.
##
## Note that some of the elements here are defined in the "Shared
## Strings" section of the spec.  Others I have not tracked down yet.
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
  sz <- xlsx_ct_font_size(xml2::xml_find_one(x, "d1:sz", ns))

  u <- xlsx_ct_underline_property(xml2::xml_find_one(x, "d1:u", ns))
  ## This one here is either baseline, superscript or subscript.  So
  ## probably not terribly useful and fairly confuse-able with
  ## _actual_ vertical alignment.

  ## vertAlign <- xml2::xml_text(xml2::xml_find_one(x, "d1:vertAlign/@val", ns))
  scheme <- xml2::xml_text(xml2::xml_find_one(x, "d1:scheme/@val", ns))

  tibble::data_frame(name, family,
                     b, i, strike, outline, shadow, condense, extend,
                     color, sz, u, scheme)
}

## 18.8.18 family
xlsx_st_font_family <- function(f, missing=NA_character_) {
  pos <- c(NA_character_, "Roman", "Swiss", "Modern", "Script", "Decorative",
           rep("<<reserved>>", 9))
  if (inherits(f, "xml_missing")) {
    missing
  } else {
    pos[[as.integer(xml2::xml_attr(f, "val")) + 1L]]
  }
}

## Used by a bunch of things.  The actual definition is on l 3751 of
## A.2 (p. 3929).  Note that the xsd defines that if the element is
## present but @val is empty it defaults to TRUE.
xlsx_ct_boolean_property <- function(b, missing=FALSE) {
  if (inherits(b, "xml_missing")) {
    missing
  } else {
    val <- xml2::xml_attr(b, "val")
    if (is.na(val)) TRUE else as.logical(as.integer(val))
  }
}

## 18.8.21 fills
xlsx_ct_fills <- function(xml, ns, theme, index) {
  process_container(xml, "d1:fills", ns, xlsx_ct_fill, theme, index)
}

## 18.8.20 fill
xlsx_ct_fill <- function(x, ns, theme, index) {
  ## TODO: In the case where not all of these are "pattern" (i.e., we
  ## have a gradient fill) this will not work correctly because we
  ## need totally different things here.  I think what we'll return
  ## there is type=gradient, and then a lookup to a gradient table, so
  ## this will expand by one more column with gradient_id perhaps.

  ## The only options here, according to the xsd (A.2, p. 3925,
  ## l. 3498) is a single element of patternFill or gradientFill
  xk <- xml2::xml_children(x)[[1L]]
  if (xml2::xml_name(xk) == "patternFill") {
    xlsx_ct_pattern_fill(xk, ns, theme, index)
  } else {
    xlsx_ct_gradient_fill(xk, ns, theme, index)
  }
}

## 18.8.32 patternFill
xlsx_ct_pattern_fill <- function(x, ns, theme, index) {
  ## This is very weird because all of the attribute patternType,
  ## fgColor and bgColor are optional.
  pattern_type <- xml2::xml_attr(x, "patternType")
  fg <- xlsx_ct_color(xml2::xml_find_one(x, "./d1:fgColor", ns), theme, index)
  bg <- xlsx_ct_color(xml2::xml_find_one(x, "./d1:bgColor", ns), theme, index)
  c(type="pattern", pattern_type=pattern_type, fg=fg, bg=bg)
}

## 18.8.24 gradientFill
xlsx_ct_gradient_fill <- function(x, ns, theme, index) {
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

## 18.8.3  bgColor
## 18.8.19 fgColor
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

## 18.8.5  borders
xlsx_ct_borders <- function(xml, ns, theme, index) {
  process_container(xml, "d1:borders", ns, xlsx_ct_border, theme, index)
}

## 18.8.4  border
##
## See also
##   * 18.8.5 (p. 1750)
##   * A.2 l. 3460 (p. 3924)
##
## Unfortunately, note that the xsd talks about start / end but the
## *example* has begin / end.  And neither of them indicates what on
## earth these are for (though the text in the example suggests that
## end is the right border in that context).  In the sheets I am
## looking at I mostly see left / right / top / bottom / diagonal.
xlsx_ct_border <- function(x, ns, theme, index) {
  ## NOTE: I am skipping attributes diagonalUp and diagonalDown along
  ## with the element diagonal - it's not the only bit of formatting
  ## trivia we won't handle, but it's a fairly unusual thing to see, I
  ## believe.
  outline <- attr_bool(xml2::xml_attr(x, "outline"), FALSE)

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
##
## This handles:
##   * 18.8.6  bottom
##   * 18.8.16 end
##   * 18.8.37 start
##   * 18.8.43 top
## as well as left and right which aren't given section numbers in the spec.
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

## 18.8.9  cellStyleXfs
xlsx_ct_cell_style_xfs <- function(xml, ns) {
  process_container(xml, "d1:cellStyleXfs", ns, xlsx_ct_xf)
}

## 18.8.10 cellXfs
xlsx_ct_cell_xfs <- function(xml, ns) {
  process_container(xml, "d1:cellXfs", ns, xlsx_ct_xf)
}

## 18.8.45 xf (format)
xlsx_ct_xf <- function(x, ns) {
  at <- as.list(xml2::xml_attrs(x))
  xf <- tibble::data_frame(
    ## Booleans, indicating if things are applied:
    apply_alignment = attr_bool(at$applyAlignment, FALSE),
    apply_border = attr_bool(at$applyBorder, FALSE),
    apply_fill = attr_bool(at$applyFill, FALSE),
    apply_font = attr_bool(at$applyFont, FALSE),
    apply_number_format = attr_bool(at$applyNumberFormat, FALSE),
    apply_protection = attr_bool(at$applyProtection, FALSE),

    ## References to actual formats (all base 0)
    border_id = attr_integer(at$borderId),
    fill_id = attr_integer(at$fillId),
    font_id = attr_integer(at$fontId),
    num_fmt_id = attr_integer(at$numFmtId),

    pivot_button = attr_bool(at$pivotButton, FALSE),
    quote_prefix = attr_bool(at$quotePrefix, FALSE),

    ## This is a reference against cellStyleXfs
    xf_id = attr_integer(at$xfId))
  alignment <- xlsx_ct_alignment(xml2::xml_find_one(x, "d1:alignment", ns))
  cbind(xf, alignment)
}

## 18.8.1  alignment
##
## horizontal: center | centerContinuous | distributed | fill |
##   general | justify | right
##
## vertical: bottom | center | distributed | justify | top
xlsx_ct_alignment <- function(x, ns) {
  at <- as.list(xml2::xml_attrs(x))
  tibble::data_frame(
    horizontal=attr_character(at$horizontal),
    vertical=attr_character(at$vertical),
    indent=attr_integer(at$indent),
    justify_last_line=attr_bool(at$justifyLastLine, FALSE),
    reading_order=attr_integer(at$readingOrder),
    ## relativeIndent [used only in a dxf element]
    shrink_to_fit=attr_bool(at$shrinkToFit, FALSE),
    text_rotation=attr_integer(at$text_rotation),
    text_wrap=attr_bool(at$textWrap, FALSE))
}

## 18.8.8  cellStyles
xlsx_ct_cell_styles <- function(xml, ns) {
  process_container(xml, "d1:cellStyles", ns, xlsx_ct_cell_style)
}

## 18.8.7  cellStyle
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
    custom_builtin = attr_bool(at$customBuiltin, FALSE),
    hidden = attr_bool(at$hidden, FALSE),
    i_level = attr_integer(at$iLevel),
    name = attr_character(at$name),
    xf_id = attr_integer(at$xfId))
}

## 18.8.31 numFmts
xlsx_ct_num_fmts <- function(xml, ns) {
  process_container(xml, "d1:numFmts", ns, xlsx_ct_num_fmt)
}

## 18.8.30 numFmt
xlsx_ct_num_fmt <- function(x, ns) {
  at <- as.list(xml2::xml_attrs(x))
  tibble::data_frame(
    num_format_id = attr_integer(at$numFmtId),
    format_code = attr_character(at$formatCode))
}

## Below here is bits that may move around a bit; code for processing
## things out into values that R can understand, mostly for colours.
## We need to do the number formatting thing soon too.

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
