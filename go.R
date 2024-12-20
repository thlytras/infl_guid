# Load all the data in /input
library(readxl)
dat.raw <- lapply(list.files("input", full=TRUE, pattern="\\.xlsx$"), function(x) as.data.frame(read_excel(x)))
names(dat.raw) <- gsub("(\\@)|( table)*\\.xlsx$", "", list.files("input", pattern="\\.xlsx$"))

library(meta)

