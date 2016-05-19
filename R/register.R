rexcel_workbook <- function(path) {
  ## TO DO:
  ## if path is actually a workbook
  ## Recall(path$path)
  ## i.e. refresh registration of the workbook
  ## to be used when you are concerned the xlsx has changed
  is_xlsx(path)
  manifest <- xlsx_list_files(path)

  ct <- xlsx_read_file(path, "[Content_Types].xml") %>%
    xml2::xml_contents() %>%
    xml2::xml_attrs() %>%
    purrr::map(as.list) %>%
    dplyr::bind_rows()

  ## this appears to be always boring? omit it
  #rels <- xlsx_read_file(path, "_rels/.rels")

  xl_workbook <- xlsx_read_file(path, "xl/workbook.xml")
  sheets <- xl_workbook %>%
    xml2::xml_find_one("//d1:sheets", xml2::xml_ns(.)) %>%
    xml2::xml_contents() %>%
    xml2::xml_attrs() %>%
    purrr::map(as.list) %>%
    dplyr::bind_rows()

  lst(xlsx_path = path,
      reg_time = Sys.time(),
      manifest,
      content_types = ct,
      sheets
  )
  ## TODO: obviously this should return a workbook object!
}
