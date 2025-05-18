source("utils.R")

get_wml_node <- function(x) {
  xml_ <- to_wml(x)
  read_xml(wml_str(xml_))
}


test_that("run_autonum wml", {
  z <- expand.grid(
    pre_label = c("Table ", "Tableau "),
    post_label = c(": ", " : "),
    stringsAsFactors = FALSE
  )
  z$bkm <- c("bkm1", "bkm2", "bkm3", "null")
  z$bkm_all <- rep(c(TRUE, FALSE), 2)
  z$seq_id <- c("tab", "tbl", "fig", "plt")
  runs <- mapply(
    .pre = z$pre_label,
    .post = z$post_label,
    .bkm = z$bkm,
    .bkm_all = z$bkm_all,
    .seq_id = z$seq_id,
    FUN = function(.pre, .post, .bkm, .bkm_all, .seq_id) {
      z <- run_autonum(
        pre_label = .pre,
        seq_id = .seq_id,
        post_label = .post,
        bkm = if (.bkm %in% "null") NULL else .bkm,
        bkm_all = .bkm_all
      )
      to_wml(z)
    },
    SIMPLIFY = FALSE,
    USE.NAMES = FALSE
  )

  bkm_pattern <- "<w\\:bookmarkStart w\\:id=\"[0-9a-z\\-]+\" w\\:name=\"%s\"/>"
  expect_match(object = runs[[1]], regexp = sprintf(bkm_pattern, z$bkm[1]))
  expect_match(object = runs[[2]], regexp = sprintf(bkm_pattern, z$bkm[2]))
  expect_match(object = runs[[3]], regexp = sprintf(bkm_pattern, z$bkm[3]))
  expect_no_match(object = runs[[4]], regexp = sprintf(bkm_pattern, z$bkm[4]))

  expect_match(
    object = runs[[3]],
    regexp = sprintf(paste0("^", bkm_pattern), z$bkm[3])
  )
  expect_no_match(
    object = runs[[4]],
    regexp = sprintf(paste0("^", bkm_pattern), z$bkm[4])
  )

  node <- read_xml(wml_str(runs[[1]]))
  expect_equal(xml_text(xml_child(node, "w:r")), "Table ")
  expect_equal(xml_text(xml_child(node, xml_length(node))), ": ")
  node <- read_xml(wml_str(runs[[4]]))
  expect_equal(xml_text(xml_child(node, "w:r")), "Tableau ")
  expect_equal(xml_text(xml_child(node, xml_length(node))), " : ")

  seq_pattern <- "<w\\:instrText xml\\:space=\"preserve\" w\\:dirty=\"true\">SEQ %s"
  expect_match(object = runs[[1]], regexp = sprintf(seq_pattern, z$seq_id[1]))
  expect_match(object = runs[[2]], regexp = sprintf(seq_pattern, z$seq_id[2]))
  expect_match(object = runs[[3]], regexp = sprintf(seq_pattern, z$seq_id[3]))
  expect_match(object = runs[[4]], regexp = sprintf(seq_pattern, z$seq_id[4]))

  z <- run_autonum(seq_id = "tbl", start_at = 2, tnd = 2, tns = ".")
  node <- read_xml(wml_str(to_wml(z)))
  expect_equal(
    xml_text(node),
    "Table STYLEREF 2 \\r.SEQ tbl \\* Arabic \\r 2: "
  )
  z <- run_autonum(seq_id = "tbl", tnd = 1, tns = "-")
  node <- read_xml(wml_str(to_wml(z)))
  expect_equal(xml_text(node), "Table STYLEREF 1 \\r-SEQ tbl \\* Arabic: ")
})
