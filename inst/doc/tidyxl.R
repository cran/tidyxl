## ---- include = FALSE----------------------------------------------------
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>",
  fig.path = "README-"
)

## ---- echo = TRUE, eval = FALSE------------------------------------------
#  devtools::install_github("nacnudus/tidyxl")

## ---- echo = TRUE--------------------------------------------------------
ftable(Titanic, row.vars = 1:2)

## ---- echo = TRUE--------------------------------------------------------
readxl::read_excel("../inst/extdata/titanic.xlsx")

## ---- echo = TRUE--------------------------------------------------------
library(tidyxl)
x <- tidy_xlsx("../inst/extdata/titanic.xlsx")$data$Sheet1
# Specific sheets can be requested using `tidy_xlsx(file, sheet)`
str(x)

## ---- echo = TRUE--------------------------------------------------------
x[x$data_type == "character", c("address", "character")]
x[x$row == 4, c("address", "character", "numeric")]

## ---- echo = TRUE--------------------------------------------------------
# Bold
formats <- tidy_xlsx("../inst/extdata/titanic.xlsx")$formats
formats$local$font$bold
x[x$local_format_id %in% which(formats$local$font$bold),
  c("address", "character")]

# Yellow fill
formats$local$fill$patternFill$fgColor$rgb
x[x$local_format_id %in%
  which(formats$local$fill$patternFill$fgColor$rgb == "FFFFFF00"),
  c("address", "numeric")]

# Styles by name
formats$style$font$name["Normal"]
head(x[x$style_format == "Normal", c("address", "character")])

## ---- echo = TRUE--------------------------------------------------------
x[!is.na(x$comment), c("address", "comment")]

## ---- echo = TRUE-----------------------------------------------------------------------------------------------------
options(width = 120)
y <- tidy_xlsx("../inst/extdata/examples.xlsx", "Sheet1")$data[[1]]
y[!is.na(y$formula),
  c("address", "formula", "formula_type", "formula_ref", "formula_group",
    "error", "logical", "numeric", "date", "character")]

