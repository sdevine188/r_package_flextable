## ---- include=FALSE------------------------------------------------------
library(flextable)
library(magrittr)

dir.create("assets/docx", recursive = TRUE, showWarnings = FALSE)
dir.create("assets/pptx", recursive = TRUE, showWarnings = FALSE)
office_doc_link <- function(url){
  stopifnot(requireNamespace("htmltools", quietly = TRUE))
  htmltools::tags$p(  htmltools::tags$span("Download file "),
    htmltools::tags$a(basename(url), href = url), 
    htmltools::tags$span(" - view with"),
    htmltools::tags$a("office web viewer", target="_blank", 
      href = paste0("https://view.officeapps.live.com/op/view.aspx?src=", url)
      ), 
    style="text-align:center;font-style:italic;color:gray;"
    )
}

knitr::opts_chunk$set(
  message = FALSE,
  collapse = TRUE,
  comment = "#>", 
  eval = !is.null(knitr::opts_knit$get("rmarkdown.pandoc.to"))
)

## ------------------------------------------------------------------------
data <- iris[c(1:3, 51:53, 101:104),]

## ----warning=FALSE, echo=FALSE-------------------------------------------
library(flextable)
library(officer)

typology <- data.frame(
  col_keys = c( "Sepal.Length", "Sepal.Width", "Petal.Length",
                "Petal.Width", "Species" ),
  what = c("Sepal", "Sepal", "Petal", "Petal", "Species"),
  measure = c("Length", "Width", "Length", "Width", "Species"),
  stringsAsFactors = FALSE )

ft <- flextable(
  data, 
  col_keys = c("Species", "sep_1", "Sepal.Length", "Sepal.Width", 
               "sep_2",  "Petal.Length", "Petal.Width" ) )

ft <- set_header_df(ft, mapping = typology, key = "col_keys" )
ft <- merge_h(ft, part = "header")
ft <- merge_v(ft, j = "Species", part = "body")
ft <- merge_v(ft, j = "Species", part = "header")
ft <- theme_booktabs(ft)
ft <- empty_blanks(ft)
ft <- fix_border_issues(ft)
autofit(ft) 

## ----eval=FALSE----------------------------------------------------------
#  library(magrittr)
#  mytable <- cars %>%
#    head() %>%
#    flextable() %>%
#    autofit()
#  
#  mytable

## ------------------------------------------------------------------------
library(officer)

## ----results='hide'------------------------------------------------------
ft <- flextable(head(mtcars))
ft <- theme_booktabs(ft)
ft <- autofit(ft)

ppt <- read_pptx()
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with_flextable(ppt, value = ft, type = "body") 

print(ppt, target = "assets/pptx/example.pptx")

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/flextable/articles/", "assets/pptx/example.pptx" ) )

## ----results='hide'------------------------------------------------------
doc <- read_docx()
doc <- body_add_flextable(doc, value = ft)
print(doc, target = "assets/docx/example.docx")

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/flextable/articles/", "assets/docx/example.docx" ) )

## ----echo=FALSE----------------------------------------------------------
dmodel <- data.frame( 
  header = rep("body", 4 ),
  header_1 = c("data[1,1]", "data[2,1]", "...", "data[n,1]" ),
  header_2 = c("data[1,2]", "data[2,2]", "...", "data[n,2]" ),
  header_3 = c("data[1,3]", "data[2,3]", "...", "data[n,3]" ),
  header_4 = c("data[1,4]", "data[2,4]", "...", "data[n,4]" ),
  header_5 = c("data[1,5]", "data[2,5]", "...", "data[n,5]" ), 
  stringsAsFactors = FALSE
)
header_model <- data.frame(
  title1 = c( "header", rep("row 1 - cell from 1 to 5", 5) ),
  title2 = c( "header", rep("row 2 - cell from 1 to 2", 2), rep("row 2 cell from 3 to 5", 3) ),
  title3 = c( "header", sprintf("row 3 cell %.0f", 1:5) ), 
  key = c( names(dmodel)), 
  stringsAsFactors = FALSE )

data_model <- flextable( data = dmodel, 
                            col_keys = c("header", "blank", "header_1", "header_2", "header_3", "header_4", "header_5" ) )
data_model <- set_header_df(data_model, mapping = header_model, key = "key")
data_model <- merge_h(data_model, part = "header")

data_model <- add_footer(data_model, 
                         header = "footer",
                         header_1 = "Grouped footer 1.1", 
                         header_2 = "Grouped footer 1.1", 
                         header_3 = "Grouped footer 1.2", 
                         header_4 = "Grouped footer 1.2", 
                         header_5 = "Grouped footer 1.2" )
data_model <- add_footer(data_model, top = FALSE,
                         header = "footer",
                         header_1 = "Grouped footer 2.1", 
                         header_2 = "Grouped footer 2.1", 
                         header_3 = "Grouped footer 2.1", 
                         header_4 = "Grouped footer 2.1", 
                         header_5 = "Grouped footer 2.1" )
data_model <- merge_h(data_model, part = "footer")
data_model <- merge_v(data_model, part = "footer", j = 1)
data_model <- merge_v(data_model, part = "body", j = 1)
data_model <- merge_v(data_model, part = "header", j = 1)
data_model <- theme_box(data_model)
data_model <- flextable::align(data_model, align = "center", part = "all")
data_model <- bg(data_model, bg = "#eb5555", part = "header")
data_model <- bg(data_model, bg = "#475F77", part = "body")
data_model <- bg(data_model, bg = "#2BBBD8", part = "footer")
data_model <- color(data_model, color = "white", part = "footer")
# data_model <- bg(data_model, j = 1, bg = "#475F77", part = "all")
data_model <- width(data_model, j = 1, width = 1.2)
data_model <- width(data_model, j = 2, width = .2)
data_model <- width(data_model, j = 3:7, width = 1.2)

std_border <- fp_border(width = 1.5, color = "white")

data_model <- border_outer(data_model, part="all", border = std_border )
data_model <- border_inner_h(data_model, border = std_border, part="all")
data_model <- border_inner_v(data_model, border = std_border, part="all")

data_model <- empty_blanks(data_model)
data_model <- padding(data_model, padding = 4, part = "all")
data_model <- fontsize(data_model, size = 12, part = "all")
data_model <- italic(data_model, italic = FALSE, part = "footer")
data_model <- bold(data_model, bold = TRUE, part = "all")
data_model <- color(data_model, color = "white", part = "all")
data_model

## ----warning=FALSE-------------------------------------------------------
library(flextable)
library(officer)

myft <- flextable(
  head(mtcars), 
  col_keys = c("am", "carb", "gear", "mpg", "drat" ))
myft

## ----warning=FALSE-------------------------------------------------------
myft <- theme_vanilla(myft)
myft

## ----warning=FALSE-------------------------------------------------------
myft <- merge_v(myft, j = c("am", "carb") )
myft <- set_header_labels( myft, carb = "# carb." )
myft <- autofit(myft)
myft

## ------------------------------------------------------------------------
myft <- italic(myft, j = 1)
myft <- bg(myft, bg = "#C90000", part = "header")
myft <- color(myft, color = "white", part = "header")
myft <- color(myft, ~ drat > 3.5, ~ drat, color = "red")
myft <- bold(myft, ~ drat > 3.5, ~ drat, bold = TRUE)
myft <- autofit(myft)
myft

## ------------------------------------------------------------------------
ft <- flextable(head(iris, n = 10))
ft <- color(ft, i = ~ Sepal.Length < 5, j = ~ Sepal.Length + Sepal.Width, color = "orange")
ft

## ------------------------------------------------------------------------
ft <- flextable(head(iris, n = 10))
ft <- color(ft, i = 1:3, j = 1:3, color = "orange")
ft

## ------------------------------------------------------------------------
dat <- head(iris, n = 10)
ft <- flextable(dat)

ft <- color(ft, i = dat$Sepal.Length < 5, color = "orange")
ft

## ------------------------------------------------------------------------
dat <- head(iris, n = 10)
ft <- flextable(dat)
ft <- color(ft, j = "Sepal.Length", color = "orange", part = "all")
ft

