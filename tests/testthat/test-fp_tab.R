test_that("fp_tab", {
  tab <- fp_tab(pos = 0.4, style = "decimal")
  expect_equal(to_wml(tab), "<w:tab w:val=\"decimal\" w:leader=\"none\" w:pos=\"576\"/>")
  expect_equal(rtf_fp_tab(tab), "\\tx576\\tqdec")
  tab <- fp_tab(pos = 0.4, style = "left")
  expect_equal(to_wml(tab), "<w:tab w:val=\"left\" w:leader=\"none\" w:pos=\"576\"/>")
  expect_equal(rtf_fp_tab(tab), "\\tx576")
  tab <- fp_tab(pos = 0.4, style = "right")
  expect_equal(to_wml(tab), "<w:tab w:val=\"right\" w:leader=\"none\" w:pos=\"576\"/>")
  expect_equal(rtf_fp_tab(tab), "\\tx576\\tqr")
  tab <- fp_tab(pos = 0.4, style = "center")
  expect_equal(to_wml(tab), "<w:tab w:val=\"center\" w:leader=\"none\" w:pos=\"576\"/>")
  expect_equal(rtf_fp_tab(tab), "\\tx576\\tqc")
  expect_error(fp_tab(pos = 0.4, style = "blah"))
})


test_that("fp_tabs", {
  z <- fp_tabs(
    fp_tab(pos = 0.4, style = "decimal"),
    fp_tab(pos = 1, style = "decimal")
  )
  expect_equal(to_wml(z), "<w:tabs><w:tab w:val=\"decimal\" w:leader=\"none\" w:pos=\"576\"/><w:tab w:val=\"decimal\" w:leader=\"none\" w:pos=\"1440\"/></w:tabs>")
})


