# Package index

## Working with Word documents

- [`read_docx()`](https://davidgohel.github.io/officer/dev/reference/read_docx.md)
  : Create a 'Word' document object
- [`print(`*`<rdocx>`*`)`](https://davidgohel.github.io/officer/dev/reference/print.rdocx.md)
  : Write a 'Word' File
- [`set_doc_properties()`](https://davidgohel.github.io/officer/dev/reference/set_doc_properties.md)
  : Set document properties
- [`docx_set_paragraph_style()`](https://davidgohel.github.io/officer/dev/reference/docx_set_paragraph_style.md)
  : Add or replace paragraph style in a Word document
- [`docx_set_character_style()`](https://davidgohel.github.io/officer/dev/reference/docx_set_character_style.md)
  : Add character style in a Word document
- [`docx_set_settings()`](https://davidgohel.github.io/officer/dev/reference/docx_set_settings.md)
  : Set 'Microsoft Word' Document Settings
- [`change_styles()`](https://davidgohel.github.io/officer/dev/reference/change_styles.md)
  : Replace styles in a 'Word' Document

## Word documents information

- [`docx_summary()`](https://davidgohel.github.io/officer/dev/reference/docx_summary.md)
  : Get Word content in a data.frame
- [`docx_comments()`](https://davidgohel.github.io/officer/dev/reference/docx_comments.md)
  : Get comments in a Word document as a data.frame
- [`styles_info()`](https://davidgohel.github.io/officer/dev/reference/styles_info.md)
  : Read 'Word' styles
- [`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md)
  : Read document properties
- [`docx_dim()`](https://davidgohel.github.io/officer/dev/reference/docx_dim.md)
  : 'Word' page layout
- [`docx_show_chunk()`](https://davidgohel.github.io/officer/dev/reference/docx_show_chunk.md)
  : Show underlying text tag structure
- [`length(`*`<rdocx>`*`)`](https://davidgohel.github.io/officer/dev/reference/length.rdocx.md)
  : Number of blocks inside an rdocx object

## Navigate into Word documents

- [`cursor_begin()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  [`cursor_bookmark()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  [`cursor_end()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  [`cursor_reach()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  [`cursor_reach_test()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  [`cursor_forward()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  [`cursor_backward()`](https://davidgohel.github.io/officer/dev/reference/cursor.md)
  : Set cursor in a 'Word' document
- [`docx_bookmarks()`](https://davidgohel.github.io/officer/dev/reference/docx_bookmarks.md)
  : List Word bookmarks

## Add content to Word documents

- [`body_add_blocks()`](https://davidgohel.github.io/officer/dev/reference/body_add_blocks.md)
  : Add a list of blocks into a 'Word' document
- [`body_add_caption()`](https://davidgohel.github.io/officer/dev/reference/body_add_caption.md)
  : Add Word caption in a 'Word' document
- [`body_add_img()`](https://davidgohel.github.io/officer/dev/reference/body_add_img.md)
  : Add an image in a 'Word' document
- [`body_add_par()`](https://davidgohel.github.io/officer/dev/reference/body_add_par.md)
  : Add paragraphs of text in a 'Word' document
- [`body_add_table()`](https://davidgohel.github.io/officer/dev/reference/body_add_table.md)
  : Add table in a 'Word' document
- [`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md)
  : Add fpar in a 'Word' document
- [`body_add_toc()`](https://davidgohel.github.io/officer/dev/reference/body_add_toc.md)
  : Add table of content in a 'Word' document
- [`body_add_break()`](https://davidgohel.github.io/officer/dev/reference/body_add_break.md)
  : Add a page break in a 'Word' document
- [`body_add_docx()`](https://davidgohel.github.io/officer/dev/reference/body_add_docx.md)
  : Add an external docx in a 'Word' document
- [`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)
  : Import an external docx in a 'Word' document
- [`body_bookmark()`](https://davidgohel.github.io/officer/dev/reference/body_bookmark.md)
  : Add bookmark in a 'Word' document
- [`body_add_gg()`](https://davidgohel.github.io/officer/dev/reference/body_add_gg.md)
  : Add a 'ggplot' in a 'Word' document
- [`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md)
  : Add plot in a 'Word' document
- [`body_end_block_section()`](https://davidgohel.github.io/officer/dev/reference/body_end_block_section.md)
  : Add any section
- [`body_set_default_section()`](https://davidgohel.github.io/officer/dev/reference/body_set_default_section.md)
  : Define Default Section
- [`body_end_section_columns()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_columns.md)
  : Add multi columns section
- [`body_end_section_columns_landscape()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_columns_landscape.md)
  : Add a landscape multi columns section
- [`body_end_section_continuous()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_continuous.md)
  : Add continuous section
- [`body_end_section_landscape()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_landscape.md)
  : Add landscape section
- [`body_end_section_portrait()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_portrait.md)
  : Add portrait section
- [`body_comment()`](https://davidgohel.github.io/officer/dev/reference/body_comment.md)
  : Add comment in a 'Word' document
- [`body_append_start_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md)
  [`write_elements_to_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md)
  [`body_append_stop_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md)
  : Fast Append context to a Word document

## Word sections

- [`prop_section()`](https://davidgohel.github.io/officer/dev/reference/prop_section.md)
  : Section properties
- [`page_mar()`](https://davidgohel.github.io/officer/dev/reference/page_mar.md)
  : Page margins object
- [`page_size()`](https://davidgohel.github.io/officer/dev/reference/page_size.md)
  : Page size object
- [`section_columns()`](https://davidgohel.github.io/officer/dev/reference/section_columns.md)
  : Section columns

## Word simple tables

- [`prop_table()`](https://davidgohel.github.io/officer/dev/reference/prop_table.md)
  : Table properties
- [`table_colwidths()`](https://davidgohel.github.io/officer/dev/reference/table_colwidths.md)
  : Column widths of a table
- [`table_conditional_formatting()`](https://davidgohel.github.io/officer/dev/reference/table_conditional_formatting.md)
  : Table conditional formatting
- [`table_layout()`](https://davidgohel.github.io/officer/dev/reference/table_layout.md)
  : Algorithm for table layout
- [`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md)
  : Paragraph styles for columns
- [`table_width()`](https://davidgohel.github.io/officer/dev/reference/table_width.md)
  : Preferred width for a table

## Replace content in Word documents

- [`body_replace_gg_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_gg_at_bkm.md)
  [`body_replace_plot_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_gg_at_bkm.md)
  : Add plots at bookmark location in a 'Word' document
- [`body_replace_text_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_text_at_bkm.md)
  [`body_replace_img_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_text_at_bkm.md)
  [`headers_replace_text_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_text_at_bkm.md)
  [`headers_replace_img_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_text_at_bkm.md)
  [`footers_replace_text_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_text_at_bkm.md)
  [`footers_replace_img_at_bkm()`](https://davidgohel.github.io/officer/dev/reference/body_replace_text_at_bkm.md)
  : Replace text at a bookmark location
- [`body_replace_all_text()`](https://davidgohel.github.io/officer/dev/reference/body_replace_all_text.md)
  [`headers_replace_all_text()`](https://davidgohel.github.io/officer/dev/reference/body_replace_all_text.md)
  [`footers_replace_all_text()`](https://davidgohel.github.io/officer/dev/reference/body_replace_all_text.md)
  : Replace text anywhere in the document

## Remove content from Word documents

- [`body_remove()`](https://davidgohel.github.io/officer/dev/reference/body_remove.md)
  : Remove an element in a 'Word' document

## Reading/writing PowerPoint documents

- [`read_pptx()`](https://davidgohel.github.io/officer/dev/reference/read_pptx.md)
  : Create a 'PowerPoint' document object
- [`print(`*`<rpptx>`*`)`](https://davidgohel.github.io/officer/dev/reference/print.rpptx.md)
  : Write a 'PowerPoint' file.

## Manipulate slides

- [`add_slide()`](https://davidgohel.github.io/officer/dev/reference/add_slide.md)
  : Add a slide
- [`layout_default()`](https://davidgohel.github.io/officer/dev/reference/layout_default.md)
  : Default layout for new slides
- [`on_slide()`](https://davidgohel.github.io/officer/dev/reference/on_slide.md)
  : Change current slide
- [`move_slide()`](https://davidgohel.github.io/officer/dev/reference/move_slide.md)
  : Move a slide
- [`remove_slide()`](https://davidgohel.github.io/officer/dev/reference/remove_slide.md)
  : Remove a slide
- [`` `slide_visible<-`() ``](https://davidgohel.github.io/officer/dev/reference/slide-visible.md)
  [`slide_visible()`](https://davidgohel.github.io/officer/dev/reference/slide-visible.md)
  : Get or set slide visibility

## Slide content

- [`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.character()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.numeric()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.factor()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.logical()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.Date()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.block_list()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.unordered_list()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.data.frame()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.gg()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.plot_instr()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.external_img()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.fpar()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.empty_content()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  [`ph_with.xml_document()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
  : Add objects on the current slide
- [`phs_with()`](https://davidgohel.github.io/officer/dev/reference/phs_with.md)
  : Fill multiple placeholders using key value syntax
- [`ph_location()`](https://davidgohel.github.io/officer/dev/reference/ph_location.md)
  : Location for a placeholder from scratch
- [`ph_location_fullsize()`](https://davidgohel.github.io/officer/dev/reference/ph_location_fullsize.md)
  : Location of a full size element
- [`ph_location_id()`](https://davidgohel.github.io/officer/dev/reference/ph_location_id.md)
  : Location of a placeholder based on its id
- [`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md)
  : Location of a named placeholder
- [`ph_location_left()`](https://davidgohel.github.io/officer/dev/reference/ph_location_left.md)
  : Location of a left body element
- [`ph_location_right()`](https://davidgohel.github.io/officer/dev/reference/ph_location_right.md)
  : Location of a right body element
- [`ph_location_template()`](https://davidgohel.github.io/officer/dev/reference/ph_location_template.md)
  : Location for a placeholder based on a template
- [`ph_location_type()`](https://davidgohel.github.io/officer/dev/reference/ph_location_type.md)
  : Location of a placeholder based on a type
- [`ph_slidelink()`](https://davidgohel.github.io/officer/dev/reference/ph_slidelink.md)
  : Slide link to a placeholder
- [`ph_hyperlink()`](https://davidgohel.github.io/officer/dev/reference/ph_hyperlink.md)
  : Hyperlink a placeholder
- [`ph_remove()`](https://davidgohel.github.io/officer/dev/reference/ph_remove.md)
  : Remove a shape
- [`set_notes()`](https://davidgohel.github.io/officer/dev/reference/set_notes.md)
  : Set notes for current slide
- [`notes_location_label()`](https://davidgohel.github.io/officer/dev/reference/notes_location_label.md)
  : Location of a named placeholder for notes
- [`notes_location_type()`](https://davidgohel.github.io/officer/dev/reference/notes_location_type.md)
  : Location of a placeholder for notes

## Read PowerPoint information

- [`pptx_summary()`](https://davidgohel.github.io/officer/dev/reference/pptx_summary.md)
  : PowerPoint content in a data.frame
- [`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md)
  : Slides width and height
- [`length(`*`<rpptx>`*`)`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md)
  : Number of slides
- [`annotate_base()`](https://davidgohel.github.io/officer/dev/reference/annotate_base.md)
  : Placeholder parameters annotation
- [`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md)
  : Presentation layouts summary
- [`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md)
  : Slide layout properties
- [`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md)
  : Slide layout properties plot
- [`layout_dedupe_ph_labels()`](https://davidgohel.github.io/officer/dev/reference/layout_dedupe_ph_labels.md)
  : Detect and handle duplicate placeholder labels
- [`layout_rename_ph_labels()`](https://davidgohel.github.io/officer/dev/reference/layout_rename_ph_labels.md)
  [`` `layout_rename_ph_labels<-`() ``](https://davidgohel.github.io/officer/dev/reference/layout_rename_ph_labels.md)
  : Change ph labels in a layout
- [`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)
  : Slide content in a data.frame
- [`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md)
  : Color scheme of a PowerPoint file
- [`media_extract()`](https://davidgohel.github.io/officer/dev/reference/media_extract.md)
  : Extract media from a document object
- [`as.matrix(`*`<rpptx>`*`)`](https://davidgohel.github.io/officer/dev/reference/as.matrix.rpptx.md)
  : PowerPoint table to matrix

## Content formatting

Define formatted paragraphs or parts of paragraphs, such as text,
calculated Word fields, sets of paragraphs, etc.

- [`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)
  [`update(`*`<fpar>`*`)`](https://davidgohel.github.io/officer/dev/reference/fpar.md)
  : Formatted paragraph
- [`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md)
  : Caption block
- [`block_gg()`](https://davidgohel.github.io/officer/dev/reference/block_gg.md)
  : 'ggplot' block
- [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)
  : List of blocks
- [`block_pour_docx()`](https://davidgohel.github.io/officer/dev/reference/block_pour_docx.md)
  : External Word document placeholder
- [`block_section()`](https://davidgohel.github.io/officer/dev/reference/block_section.md)
  : Section for 'Word'
- [`block_table()`](https://davidgohel.github.io/officer/dev/reference/block_table.md)
  : Table block
- [`block_toc()`](https://davidgohel.github.io/officer/dev/reference/block_toc.md)
  : Table of content for 'Word'
- [`unordered_list()`](https://davidgohel.github.io/officer/dev/reference/unordered_list.md)
  : Unordered list
- [`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md)
  : Wrap plot instructions for png plotting in Powerpoint or Word
- [`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md)
  : Formatted chunk of text
- [`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md)
  : Auto number
- [`run_bookmark()`](https://davidgohel.github.io/officer/dev/reference/run_bookmark.md)
  : Bookmark for 'Word'
- [`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md)
  : Column break for 'Word'
- [`run_comment()`](https://davidgohel.github.io/officer/dev/reference/run_comment.md)
  : Comment for 'Word'
- [`run_footnote()`](https://davidgohel.github.io/officer/dev/reference/run_footnote.md)
  : Footnote for 'Word'
- [`run_footnoteref()`](https://davidgohel.github.io/officer/dev/reference/run_footnoteref.md)
  : Word footnote reference
- [`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md)
  : Page break for 'Word'
- [`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md)
  : Page break for 'Word'
- [`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md)
  : Cross reference
- [`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md)
  : Tab for 'Word'
- [`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md)
  [`run_seqfield()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md)
  : 'Word' computed field
- [`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)
  : Word chunk of text with a style
- [`set_autonum_bookmark()`](https://davidgohel.github.io/officer/dev/reference/set_autonum_bookmark.md)
  : Update bookmark of an autonumber run
- [`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md)
  : External image
- [`floating_external_img()`](https://davidgohel.github.io/officer/dev/reference/floating_external_img.md)
  : Floating external image
- [`hyperlink_ftext()`](https://davidgohel.github.io/officer/dev/reference/hyperlink_ftext.md)
  : Formatted chunk of text with hyperlink
- [`empty_content()`](https://davidgohel.github.io/officer/dev/reference/empty_content.md)
  : Empty block for 'PowerPoint'
- [`sp_line()`](https://davidgohel.github.io/officer/dev/reference/sp_line.md)
  [`print(`*`<sp_line>`*`)`](https://davidgohel.github.io/officer/dev/reference/sp_line.md)
  [`update(`*`<sp_line>`*`)`](https://davidgohel.github.io/officer/dev/reference/sp_line.md)
  : Line properties
- [`sp_lineend()`](https://davidgohel.github.io/officer/dev/reference/sp_lineend.md)
  [`print(`*`<sp_lineend>`*`)`](https://davidgohel.github.io/officer/dev/reference/sp_lineend.md)
  [`update(`*`<sp_lineend>`*`)`](https://davidgohel.github.io/officer/dev/reference/sp_lineend.md)
  : Line end properties

## Formatting properties

- [`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  [`fp_text_lite()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  [`format(`*`<fp_text>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  [`print(`*`<fp_text>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  [`update(`*`<fp_text>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  : Text formatting properties
- [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
  [`fp_par_lite()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
  [`print(`*`<fp_par>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
  [`update(`*`<fp_par>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
  : Paragraph formatting properties
- [`fp_border()`](https://davidgohel.github.io/officer/dev/reference/fp_border.md)
  [`update(`*`<fp_border>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_border.md)
  : Border properties object
- [`fp_cell()`](https://davidgohel.github.io/officer/dev/reference/fp_cell.md)
  [`format(`*`<fp_cell>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_cell.md)
  [`print(`*`<fp_cell>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_cell.md)
  [`update(`*`<fp_cell>`*`)`](https://davidgohel.github.io/officer/dev/reference/fp_cell.md)
  : Cell formatting properties
- [`fp_tabs()`](https://davidgohel.github.io/officer/dev/reference/fp_tabs.md)
  : Tabs properties object
- [`fp_tab()`](https://davidgohel.github.io/officer/dev/reference/fp_tab.md)
  : Tabulation mark properties object

## RTF

- [`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md)
  : Create an RTF document object
- [`rtf_add()`](https://davidgohel.github.io/officer/dev/reference/rtf_add.md)
  : Add content into an RTF document
- [`print(`*`<rtf>`*`)`](https://davidgohel.github.io/officer/dev/reference/print.rtf.md)
  : Write an 'RTF' File

## Excel

- [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
  [`length(`*`<rxlsx>`*`)`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
  [`print(`*`<rxlsx>`*`)`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
  : Create an 'Excel' document object
- [`add_sheet()`](https://davidgohel.github.io/officer/dev/reference/add_sheet.md)
  : Add a sheet
- [`sheet_select()`](https://davidgohel.github.io/officer/dev/reference/sheet_select.md)
  : Select sheet

## Misc

- [`officer-package`](https://davidgohel.github.io/officer/dev/reference/officer.md)
  [`officer`](https://davidgohel.github.io/officer/dev/reference/officer.md)
  : Manipulate Microsoft Word and PowerPoint Documents with 'officer'
- [`open_file()`](https://davidgohel.github.io/officer/dev/reference/open_file.md)
  : Opens a file locally
- [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
  [`length(`*`<rxlsx>`*`)`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
  [`print(`*`<rxlsx>`*`)`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
  : Create an 'Excel' document object
- [`add_sheet()`](https://davidgohel.github.io/officer/dev/reference/add_sheet.md)
  : Add a sheet
- [`sheet_select()`](https://davidgohel.github.io/officer/dev/reference/sheet_select.md)
  : Select sheet
- [`shortcuts`](https://davidgohel.github.io/officer/dev/reference/shortcuts.md)
  : shortcuts for formatting properties
