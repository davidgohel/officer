x <- read_pptx()

# INFO -------------

# Returns layout's ph_labels by default in same order as layout_properties()
layout_rename_ph_labels(x, "Comparison")
layout_properties(x, "Comparison")$ph_label


# BASICS -----------
#
# HINT: run `plot_layout_properties(x, "Comparison")` to see how labels change

# rename using key-value pairs: 'old label' = 'new label' or 'id' = 'new label'
layout_rename_ph_labels(x, "Comparison", "Title 1" = "LABEL MATCHED") # label matching
layout_rename_ph_labels(x, "Comparison", "3" = "ID MATCHED") # id matching
layout_rename_ph_labels(x, "Comparison", "Date Placeholder 6" = "DATE", "8" = "FOOTER") # label, id

# rename using a named list and the .dots arg
renames <- list("Content Placeholder 3" = "CONTENT_1", "6" = "CONTENT_2")
layout_rename_ph_labels(x, "Comparison", .dots = renames)

# rename via rhs assignment and optional index (not id!)
layout_rename_ph_labels(x, "Comparison") <- LETTERS[1:8]
layout_rename_ph_labels(x, "Comparison")[1:3] <- paste("CHANGED", 1:3)

# rename via rhs assignment and ph id (not index)
layout_rename_ph_labels(x, "Comparison", id = c(2, 4)) <- paste("ID =", c(2, 4))


# MORE ------------

# make all labels lower case
labels <- layout_rename_ph_labels(x, "Comparison")
layout_rename_ph_labels(x, "Comparison") <- tolower(labels)

# rename all labels to type [type_idx]
lp <- layout_properties(x, "Comparison")
layout_rename_ph_labels(x, "Comparison") <- paste0(lp$type, " [", lp$type_idx, "]")

# rename duplicated placeholders (see also `layout_dedupe_ph_labels()`)
file <- system.file("doc_examples", "ph_dupes.pptx", package = "officer")
x <- read_pptx(file)
lp <- layout_properties(x, "2-dupes")
idx <- which(lp$ph_label == "Content 7") # exists twice
layout_rename_ph_labels(x, "2-dupes")[idx] <- paste("DUPLICATE", seq_along(idx))

# warning: in case of duped labels only the first occurrence is renamed
x <- read_pptx(file)
layout_rename_ph_labels(x, "2-dupes", "Content 7" = "new label")
