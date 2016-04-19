## [x] 18.7.1 author (Author) -- xlsx_ct_authors
## [x] 18.7.2 authors (Authors) -- (in xlsx_ct_author)
## [x] 18.7.3 comment (Comment) -- xlsx_ct_comment
## [x] 18.7.4 commentList (List of Comments) -- xlsx_ct_comment_list
## [ ] 18.7.5 commentPr (Comment Properties)
## [x] 18.7.6 comments (Comments) -- xlsx_ct_comments
## [x] 18.7.7 text (Comment Text) -- xlsx_ct_rst
xlsx_read_comments <- function(path, target) {
  xml <- xlsx_read_file(path, target)
  ns <- xml2::xml_ns(xml)
  xlsx_ct_comments(xml, ns)
}

xlsx_ct_comments <- function(xml, ns) {
  authors <- xlsx_ct_authors(xml, ns)
  comments <- xlsx_ct_comment_list(xml, ns, authors)
}

xlsx_ct_authors <- function(xml, ns) {
  vcapply(xml2::xml_children(xml2::xml_find_one(xml, "d1:authors", ns)),
          xml2::xml_text)
}

xlsx_ct_comment_list <- function(xml, ns, authors) {
  process_container(xml, "d1:commentList", ns, xlsx_ct_comment, authors)
}

xlsx_ct_comment <- function(x, ns, authors) {
  at <- as.list(xml2::xml_attrs(x))
  text <- xlsx_ct_rst(xml2::xml_find_one(x, "d1:text", ns), ns)
  tibble::data_frame(
    ref = attr_character(at$ref),
    author = authors[attr_integer(at$authorId) + 1L],
    shape_id = attr_integer(at$shapeId),
    text = text)
}
