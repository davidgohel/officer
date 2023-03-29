# officer 0.6.3

## Issues

- Internal function `is_doc_open()` replaces `is_office_doc_edited()` to check if a document is open on Windows. This solves the issue where RStudio crashed while trying to write to an open Office document.


# officer 0.6.2

## Issues

- encode char to rtf from 128 instead of 256
- force 'firstLineChars' to 0 so that Chinese users don't get 
unwanted first indentation.
- avoid segfaults with SVG images by using the function `image_read_svg()`
when source is SVG in `external_img()`.

# officer 0.6.1

## Features

- add support for flextable in `block_list()`.
- add RTF support for `run_autonum()` and `run_reference()`, 
this enables captions support for flextable.

## Changes

- 'ragg' is used instead of base png because it manages 
any fonts.
- defunct `slip_in_footnote()`.

## Issues

- fix for `to_wml.block_caption()`


# officer 0.6.0

## Features

- add RTF support, see `rtf_doc()` and `rtf_add()`.

## Internals

- provides `image_to_base64()` and `uuid_generate()` as a tool 
for other 'officeverse' packages.
- styles are now injected as is and updated with `process_stylenames()`,
this allow to depend on a reference doc.

# officer 0.5.2

## Features

- if magick is available and argument `guess_size=TRUE`, image 
size is read and do not need to be provided.

## Issues

- fix images and links in sections headers and footer.
- fix `body_add()` content order.
- fix issue with properties in 'Office' documents.
- check arguments of `run_autonum()` and add tests.
- update only officer sections

## Internals

- remove `$get_at_cursor()`

# officer 0.5.1

## Issues

- manage also even/first/default sections not defined with officer (indicate 
there is even/first/default parts).
- fix behavior regression with headers and footers replace_text_at_bkm. 
The previous behavior "don't trigger an error if a bookmark is missing" 
is back. 
- force Word default continuous section if no type is specified and first 
page footer/header if it is seen.
- drop "Compatibility Mode" from Word documents.

## Changes

- Change license to MIT.
- drop internal `get_at_cursor()`
- use images sha1 value for filename with package 'openssl'. This should fix 
issues with duplicated images basenames, or non utf names, etc.

## Features

- export `shape_properties_tags()` for pptx extentions.

# officer 0.5.0

## New features

- new function `docx_set_character_style()` to add or 
replace a Word character style.
- new function `docx_set_paragraph_style()` to add or 
replace a Word paragraph style.
- new function `run_wordtext()` to add a chunk of text 
associated with a Word character style.
- `potx` and `dotx` files are now supported.
- SVG images are now supported
- Word sections can now have headers and footers. See `prop_section()`.
- function `cursor_reach_test()` to test if an expression has 
a match in the document.
- add function `docx_current_block_xml()` to let developpers access 
the xml content where the cursor is.  
- `ph_with.xml_document` now manages images paths in slide and 
treats them in the relevant documents.

## Issues

- close 'Slide Master View' automatically.
- more careful reading of document properties
- fix cursor behavior, when `pos='before'`, cursor 
is now set on element added 'before'. Internals about 
cursors have been refactored.

## Changes

- function `slip_in_footnote()` is deprecated. Use `run_footnote()` instead.

# officer 0.4.4

## New features

- `fp_par()` now have argument `word_style`
- add support for `ln` for `external_img` provided by Angus Moore
- add word_title and word_description to table properties (`prop_table()`). 
These values can be used as alternative text for Word tables. These values 
can also be set as "knitr" chunk options.

## Issues

- fix for `is_office_doc_edited()` provided by Andrew Tungate.
- `tab.lp` is no more set to null with usual rmarkdown outputs
- autofit was never used when outputting prop_table to xml

## Changes

- Function `opts_current_table()` read some Quarto values related to 
captions in order to reuse them later.

# officer 0.4.3

## New features

* `set_doc_properties()` now supports any character property. 
It provides an easy way to insert arbitrary fields. Given the challenges 
that can be encountered with find-and-replace in word with officer, the 
use of document fields and quick text fields provides a much more robust 
approach to automatic document generation from R.
* Adding support for fields (e.g. auto slide number) in Powerpoint (#429), 
see `empty_content()`.
* add functionality to set shape geometry and outline, 
see `ph_location(geam=...)`.

## Changes

- drop shortcuts$slip_in_tableref() and shortcuts$slip_in_plotref()
- defunct slip_in_column_break() and slip_in_xml()
- defunct slip_in_text()
- remove defunct functions `slip_in_img()`, `ph_add_fpar()`, 
`ph_add_par()` and `ph_add_text()`


# officer 0.4.2

## New features

* new as.matrix method for pptx to automatically extract one or all matrices
from the file.
* simple ([character()] and [block_list()]) speaker notes can 
now be added to a pptx presentation

## Changes

* deprecate slip_in_column_break and slip_in_text
* doc: use 'title case' for the titles of function manuals
* closing issues tab, opening discussion 
# officer 0.4.1

## Issues

* fix a bug in `ph_with.external_img()` that could be seen when `alt_text` was null
* change default value for `tab.cap.sep` from ":" to ": "
* fix an issue with `body_end_section_columns()` that is expected as 'continuous'.

## New features

* new parameter `scale` added to `ph_with.gg`, `body_add_gg` and `body_add.gg` 
to set the scale of ggplot outputs (like in ggsave).
* new function `set_autonum_bookmark()` to recycle an object 
made by `run_autonum()` by changing the bookmark value.
* add `tnd` argument to prefix a autonumber with the 
title number (i.e. 4.3-2 for figure 2 of chapter 4.3).
* `unordered_list()` now supports `level_list` < 1 which 
will be interpreted as 'no bullet'.
* add support to knitr table options `tab.cap.fp_text` to let format caption prefix 
in function `opts_current_table()`.

## Deprecation 

* deprecate almost all `slip_in*()` and `ph_add*()` functions. Functions `ftext()` and related 
used with `fpar()` are to be used as replacement.

# officer 0.4.0

## New features

* new function `fp_text_lite()` that do not force to provide a value 
for each properties - if a value is not provided, its attribute will 
not be written and as a result, the default properties will be applied. 
Function `fp_text()` has also been adapted, it now supports NA meaning 
to not write the attributes as in `fp_text_lite()`.
* new function `run_footnote` to add footnotes in a Word document (it 
also makes possible to deprecate totally slip_in* functions).

## Issues

* fix a bug when creating the XML of table properties for Word documents

# officer 0.3.19

## New features

* On Windows, the pptx file will not be overwritten if it is edited.
* `run_autonum` gained new argument start_at.
* `tab.topcaption` is now supported by `opts_current_table()`

## Issues

* fix issue with document properties that are null
* angle was not always preserved in PowerPoint
* fix pptx scrapping for grouped object containing a table
* fix invalid hyperlinks in docx and pptx
* fix issue with duplicated images basenames
* internal `get_reference_value` - fix - if reference_data is not an existing file, 
it is appened to `opts_knit$get("output.dir")`.

# officer 0.3.18

## Change

The sections are now corrected as follow, each section will be completed with 
the values of the default section if the value is missing. This should solve 
issue that lot of users have with page breaks when using sections. Now page 
breaks should disappear. 

# officer 0.3.17

## New features

* alt-text for images
* On Windows, the docx file will not be overwritten if it is edited.
* `fp_text` gained arguments to specify different fonts when mixing 
  CJK and latin characters: `cs.family`, `eastasia.family`, `hansi.family`.

## Issues

* fix issue with document properties where values have to be html-escaped


# officer 0.3.16

## New features

* new chunk function `hyperlink_ftext()`.

## Change

* drop cairo usage and let user settings do the job

## Issues

* fix annotate_base that was not presenting the correct informatons (thanks to John Harrold).


# officer 0.3.15

## New features

* new function `body_add_plot()` and `body_add_caption()`
* new function `run_bookmark()` to create a run with a bookmark.
* new function `body_set_default_section()` that changes default section properties
of a Word document.

## Issues

* fix #333: issue with &, <, > in `to_wml.block_table`.
* fix https://github.com/davidgohel/officedown/issues/41: when a space was in the path, pandoc send the short path format
instead of the real path name, it needed to be transformed with `normalizePath`.

## Change

* The documentation has been rewritten and can now be found
at the following URL: https://ardata-fr.github.io/officeverse/.
* `run_word_field` will supersed `run_seqfield`.
* argument `prop` of `ftext()` now default to NULL

# officer 0.3.14

## Issues

* revert PR#319 that introduced a major issue with repeated images
in Word document.

## Changes

* remove defunct functions (ph_with_*at) and set deprecated as defunct (other ph_with_*at)

# officer 0.3.13

## Issues

* fix sections issue with page margin (it effected also officedown).
* fix issue #320 (with URL encoding in .rel files).
* fix encoding issue with ftext
* explicit coding of bold, italic and uunderline attributes
* ftext with a name for chunk style is fixed

## New features

* change_styles - support for run/character, paragraph and table styles.
* new function `table_stylenames()` to define columns stylenames to be used in
tables, it benefits to `block_table` and `body_add_table`.
* line spacing is now a feature of `fp_par`
* run_autonum, run_seqfield and run_reference can now be formatted with
an object of class `fp_text`.

# officer 0.3.12

## New features

* added `body_end_block_section` and `body_add.block_section` so that users are free to add any
section they want

## Issues

* Use pandoc.exe when platform is windows

# officer 0.3.11

## Issues

* pandoc availability is now checked (for solaris and CRAN policy)

# officer 0.3.10

## New features

* new argument `level_list` in function `ph_with.block_list` ; you can now format block_list
as lists in PowerPoint.
* functions `body_add_table`, `ph_with.data.frame`, `body_add.data.frame` get
new argument `alignment` for column alignments.

## Issues

* run_reference now bookmarks only the number.

# officer 0.3.9

## Issues

* fix issues with invisible images when using ph_location_fullsize and a *blank* layout
* Embedded files in initial Word document were lost when they were read and printed
* fix run_reference issue with characters '-', '_'.

## New features

* enrich table blocks (`block_table()`) with table parameters such as width, layout. See `prop_table()`.
* new function `plot_layout_properties()` to help identifying placeholders on layouts
# officer 0.3.8

## Issues

* fix border issues with word paragraphs
* reverse changes to `body_add_blocks` and `body_add_gg` as it generated issues with cursor
* let pptx template have an empty master slide

## Changes

* internals ; drop digest from dependencies

# officer 0.3.7

## New features

* new function `get_reference_value` to read reference template used
by R Markdown.
* `fp_par` now have an argument "keep_with_next" that specifies that the paragraph (or at least
part of it) should be rendered on the same page as the next paragraph when possible.
* new experimental function `body_add` and associated methods

## Changes

* internals ; drop Rcpp and htmltools dependencies

## Issues

* fix ph_with.external_img #265 (placeholder properties were not used)
* fix ph_add_text when paragraph was empty and no content was rendered

# officer 0.3.6

## Enhancement

* support now for template generated from google docs thanks to Adam Lyon
* ph_location results can now be assigned and used as normal objetcs (and not only
as quosures).

## Changes

* `id_chr` is now depreacted in favor of `id` in function `ph_remove`, `ph_slidelink`, `ph_hyperlink`,
  `ph_add_text`, `ph_add_par`, `ph_add_fpar`.

## Issues

* fix ph_with.xml_document so that placeholders' labels are not forgotten (for rvg, flextable, etc.)
* fix underline text issue when used with powerpoint (#229).
* fix slip_in_text issue by escaping HTML entities (#234).
* fix issue with move_slide (#223).


# officer 0.3.5

## Enhancement

* new method `ph_with.xml_document` that will replace `ph_with` and `ph_with_at`.

## Issues

* fix properties inheritance with `ph_with` function.

# officer 0.3.4

## Enhancement

* new function `sanitize_images` to avoid file size inflation when replacing images
* svg support (will require rsvg package)

## Issues

* fix `external_img` size issue with method `ph_with`.
* fix bg inheritance when using `ph_with functions.

# officer 0.3.3

## Enhancement

* new generic function `ph_with()` and function `ph_location*()` to ease insertion
  of elements.
* new function `slide_size()` provide size of slides.

## Issues

* fix issue with fonts when east asian characters are used in Word.

# officer 0.3.2

## Enhancement

* new function `change_styles()` to change paragraph styles in
  a Word document.
* new function `move_slide()`: move a slide in a presentation.
* fix body_add_docx examples

## Issues

* fix issue with text underlined  and justified paragraphs in Word.
* skip errored test on macOS that can be read on CRAN check page.
* all examples are now generated in `tempdir()`.

# officer 0.3.1

## Issues

* fix function `body_add_fpar()` when argument `style` was used.
* `slide_summary` was using a bad xpath query.
* fixed character encoding issue for filename whith windows OS

# officer 0.3.0

## Enhancement

* function cursor_bookmark now let set the cursor in a text box thanks to
  Noam Ross. cursor_forward and cursor_backward can now fail if cursor
  is on a textbox but an error message will explain it to the user.
* Word documents support now footnotes.
* Word section functions have been refactored.
* New functions for replacement in headers and footers in Word documents.
  See functions `headers_replace*` and `footers_replace*`
* PowerPoint processing optimisation when generating more than few slides.

## Issues

* fix an issue with `layout_properties` when master layout is empty.

# officer 0.2.2

## Enhancement

* rdocx objects support external docx insertion
* Word margins can be modified now (thanks to Jonathan Cooper)
* New function `ph_fpars_at()` to add several formated paragraphs
  in a new shape.
* Function annotate_base will generate a slide to identify the
  placeholder indexes, master names and indexes.

## Issues

* fix issue with duplicated lines in layout_properties(#103)
* new argument par_default in ph_add_fpar so fpar par. properties
  can be kept as is.
* fix issue with images when duplicated `basename()`s

# officer 0.2.1

## Issues

* fix issue #97 with function `pptx_summary()`


# officer 0.2.0

## Enhancement

* new function `body_replace_all_text()` to replace
  any text in a Word document
* new functions for xlsx files (experimental).
* new functions `ph_with_gg()` and `ph_with_gg_at()` to make easier
  production of ggplot objects in PowerPoint
* new functions `ph_with_ul()` to make easier
  production of unordered lists of text in PowerPoint

## Issues

* an error is raised when adding an image with blank(s) in
  its basename (i.e. /home/user/bla bla.png).

# officer 0.1.8

## Issues

* decrease execution time necessary to add elements into big slide deck
* fix encoding issue in function "*_add_table"
* fix an issue with complex slide layouts (there is still an issue left but
  don't know how to manage it for now)

## Changes

* Functions slide_summary and layout_properties now return inches.

# officer 0.1.7

## Enhancement

* new function `body_replace_at` to replace text inside bookmark
* argument header for `body_add_table` and `ph_with_table`.
* `layout_properties` now returns placeholder id when available.

## Issues

* an error is now occurring when an incorrect index is used with ph_with_* functions.

# officer 0.1.6

## Enhancement

* function `ph_empty_at` can now make new shapes inherit
  properties from template

## Changes

* drop gdtools dependency

# officer 0.1.5

## Enhancement

* new function `body_default_section`
* fp_border supports width in double precision

## Issues

* characters <, > and & are now html encoded
* on_slide index is now the correct slide number id.

## Changes

* drop dplyr deprecated verbs from code
* rename `break_column` to `break_column_before`.

# officer 0.1.4

## Issues

* `body_end_section` is supposed to only work with cursor on a paragraph, an error is raised now if ending a section on something else than a paragraph.

## Enhancement

* read_pptx run faster than in previous version thanks to some code refactoring


# officer 0.1.3

## new feature

* new function media_extract to extract a media file from a document object. This function can be used to access images stored in a PowerPoint file.

## Issues

* drop magick dependence

# officer 0.1.2

## new features

* new functions `docx_summary` and `pptx_summary` to import content of an Office document into a tidy data.frame.
* new function `docx_dim()` is returning current page dimensions.
* new functions `set_doc_properties` and `doc_properties` to let you modify/access metadata of Word and PowerPoint documents.
* cursor can now reach paragraphs with a bookmark (functions `body_bookmark` and `cursor_bookmark`).
* Content can be inserted at any arbitrary location in PowerPoint (functions `ph_empty_at`, `ph_with_img_at` and `ph_with_table_at`).

## Issues

* cast all columns of data.frame as character when using ph_with_table and body_add_table
* fix pptx when more than 9 slides

# officer 0.1.1

## Enhancement

* argument `style` of functions `body_add*` and `slip_in*` now will use docx default style if not specified
* new function body_add_gg to add ggplots to Word documents
* new function test_zip for diagnostic purpose

## API changes

* classes `docx` and `pptx` have been renamed `rdocx` and `pptx` to avoid conflict with package ReporteRs.

