---
title: "How to read formate out of Excel?"
author: "Dr.lumine"
output: 
  html_document:
    keep_md: true
editor_options: 
  chunk_output_type: inline
---



## Acknowledgement

The tutorial thanks to: <https://www.r-bloggers.com/2018/05/tidying-messy-excel-data-tidyxl/>

Introduce: 

  - `library(tidyxl)` 
  - `library(upivotr)`

## installation


```r
setwd("~/Dropbox/Projects/Learning - R general/01 - Read Excel")
#devtools::install_github('nacnudus/unpivotr')
library(tidyxl)
library(unpivotr)
library(tidyverse)
```

```
## ── Attaching packages ─────────────────────────────────────── tidyverse 1.3.1 ──
```

```
## ✔ ggplot2 3.3.6     ✔ purrr   0.3.4
## ✔ tibble  3.1.7     ✔ dplyr   1.0.9
## ✔ tidyr   1.2.0     ✔ stringr 1.4.0
## ✔ readr   2.1.2     ✔ forcats 0.5.1
```

```
## ── Conflicts ────────────────────────────────────────── tidyverse_conflicts() ──
## ✖ dplyr::filter() masks stats::filter()
## ✖ dplyr::lag()    masks stats::lag()
## ✖ tidyr::pack()   masks unpivotr::pack()
## ✖ tidyr::unpack() masks unpivotr::unpack()
```


```r
list.files("data", "Excel", full.names = T)
```

```
## [1] "data/Colored Excel.xlsx"
```

```r
.libPaths()
```

```
## [1] "/Library/Frameworks/R.framework/Versions/4.2/Resources/library"
```



```r
cx <- paste0(getwd(), "/data/Colored Excel.xlsx") |> xlsx_cells()
cx.colors <- cx |>
  dplyr::filter(sheet == "a.colors")

names(cx.colors)
```

```
##  [1] "sheet"               "address"             "row"                
##  [4] "col"                 "is_blank"            "data_type"          
##  [7] "error"               "logical"             "numeric"            
## [10] "date"                "character"           "character_formatted"
## [13] "formula"             "is_array"            "formula_ref"        
## [16] "formula_group"       "comment"             "height"             
## [19] "width"               "style_format"        "local_format_id"
```


## Formating

```r
cx.colors$local_format_id
```

```
##  [1]  1  1  2  1  3  1  4  1  5  1  6  1  8  1  9  1  7  1 10  1
```
So two things dictating the formate of a cx object `syle_format` and `local_format_id`, the later is an id that stores the style imformation. Later will show you how this is extracted in a completely different file

> To look up the local formatting of a given cell, take the cell's `local_format_id` value `(my_cells$Sheet1[1, "local_format_id"])`, and use it as an index into the format structure. E.g. to look up the font size, `my_formats$local$font$size[local_format_id]`. To see all available formats, type str(my_formats$local).

### Background Color
Strangely background color is stored in `fgColor$rgb` not anywhere else

```r
fx <- paste0(getwd(), "/data/Colored Excel.xlsx") |> xlsx_formats()
```
The best way to do this is to through a function

```r
id_to_format <- function(x){
  fx$local$fill$patternFill$fgColor$rgb[x]
}
cx.colors |>
  behead("up", "headler", character) |>
  mutate(color = id_to_format(local_format_id)) |>
  pivot_wider(id_cols = c(row),names_from = "headler", values_from = "color")
```

```
## # A tibble: 9 × 3
##     row colors   index
##   <int> <chr>    <chr>
## 1     2 FF70AD47 <NA> 
## 2     3 FF70AD47 <NA> 
## 3     4 FFED7D31 <NA> 
## 4     5 FFFFC000 <NA> 
## 5     6 FFFFC000 <NA> 
## 6     7 FF70AD47 <NA> 
## 7     8 FF5B9BD5 <NA> 
## 8     9 FF4472C4 <NA> 
## 9    10 FFED7D31 <NA>
```

### deal with tables floating out of nowhere

```r
cx.b <- filter(cx, sheet == "b.floating tables")
cx.b |> 
  rectify()
```

```
## # A tibble: 6 × 5
##   `row/col` `2(B)`     `3(C)` `4(D)`    `5(E)`     
##       <int> <chr>      <chr>  <chr>     <chr>      
## 1         4 time       number text      categorical
## 2         5 2022-07-07 1      Anger     -1         
## 3         6 2022-07-08 2      Deniel    -1         
## 4         7 2022-07-09 3      Depressed -1         
## 5         8 2022-07-10 4      Bargain   1          
## 6         9 2022-07-11 5      Hope      1
```

```r
cx.b |>
  behead("up", "header", character) |>
  distinct(header)
```

```
## # A tibble: 4 × 1
##   header     
##   <chr>      
## 1 time       
## 2 number     
## 3 text       
## 4 categorical
```
### Merged

```r
cx.c <- dplyr::filter(cx, sheet == "c.floating merged cells")
cx.c |>
  rectify()
```

```
## # A tibble: 18 × 3
##    `row/col` `3(C)`                          `4(D)`
##        <int> <chr>                            <dbl>
##  1         7 This is a perfectly merged cell      1
##  2         8 <NA>                                 2
##  3         9 <NA>                                 3
##  4        10 <NA>                                 4
##  5        11 <NA>                                 5
##  6        12 <NA>                                 6
##  7        13 <NA>                                 7
##  8        14 <NA>                                 8
##  9        15 <NA>                                 9
## 10        16 <NA>                                10
## 11        17 <NA>                                11
## 12        18 <NA>                                12
## 13        19 <NA>                                NA
## 14        20 <NA>                                NA
## 15        21 Hello                                1
## 16        22 <NA>                                 2
## 17        23 <NA>                                 3
## 18        24 <NA>                                 4
```

```r
cx.c |>
  behead("left-up", "merged", character) |>
  unpivotr::pack() |> # pack second so you can creat headlers
  select(row, merged, value) |>
  unpivotr::unpack(value) # unpack later so you can use that
```

```
## # A tibble: 16 × 4
##      row merged                          data_type numeric
##    <int> <chr>                           <chr>       <dbl>
##  1     7 This is a perfectly merged cell numeric         1
##  2     8 This is a perfectly merged cell numeric         2
##  3     9 This is a perfectly merged cell numeric         3
##  4    10 This is a perfectly merged cell numeric         4
##  5    11 This is a perfectly merged cell numeric         5
##  6    12 This is a perfectly merged cell numeric         6
##  7    13 This is a perfectly merged cell numeric         7
##  8    14 This is a perfectly merged cell numeric         8
##  9    15 This is a perfectly merged cell numeric         9
## 10    16 This is a perfectly merged cell numeric        10
## 11    17 This is a perfectly merged cell numeric        11
## 12    18 This is a perfectly merged cell numeric        12
## 13    21 Hello                           numeric         1
## 14    22 Hello                           numeric         2
## 15    23 Hello                           numeric         3
## 16    24 Hello                           numeric         4
```
- row 19-20 is gone 
- it is gone when you `behead` everything
- the title we created from "up-left" because it has always been an up left, made sure everything is beheaded.

So in short it would be impossible to umerge cells if there is no other value besides it.

> Do a merged cell always have value placed on top left? 

The answer is yes try place a value in excel, merge outside, unmerge it again, 
the location turn.


```r
cx.d <- cx |> filter(sheet == "d.paralle merged cells")

## Value layer
cx.d |>
  unpivotr::pack() |> 
  rectify(value = value)
```

```
## # A tibble: 17 × 14
##    `row/col` `2(B)`   `3(C)` `4(D)` `5(E)`   `6(F)` `7(G)` `8(H)` `9(I)` `10(J)`
##        <int> <chr>     <dbl> <list> <chr>     <dbl> <list> <chr>  <list> <chr>  
##  1         3 <NA>         NA <NULL> <NA>         NA <NULL> <NA>   <NULL> <NA>   
##  2         4 <NA>         NA <NULL> <NA>         NA <NULL> <NA>   <NULL> <NA>   
##  3         5 <NA>         NA <NULL> <NA>         NA <NULL> <NA>   <NULL> block 3
##  4         6 <NA>         NA <NULL> <NA>         NA <NULL> Borde… <NULL> <NA>   
##  5         7 <NA>         NA <NULL> <NA>         NA <NULL> <NA>   <NULL> <NA>   
##  6         8 merged a      1 <NULL> merged …      1 <NULL> <NA>   <NULL> <NA>   
##  7         9 <NA>          2 <NULL> <NA>          2 <NULL> <NA>   <NULL> <NA>   
##  8        10 <NA>          3 <NULL> <NA>          3 <NULL> <NA>   <NULL> <NA>   
##  9        11 <NA>          4 <NULL> <NA>          4 <NULL> <NA>   <NULL> <NA>   
## 10        12 <NA>          5 <NULL> <NA>         NA <NULL> <NA>   <NULL> <NA>   
## 11        13 <NA>          6 <NULL> <NA>         NA <NULL> <NA>   <NULL> <NA>   
## 12        14 <NA>          7 <NULL> merged …      1 <NULL> <NA>   <NULL> <NA>   
## 13        15 <NA>          8 <NULL> <NA>          2 <NULL> <NA>   <NULL> <NA>   
## 14        16 <NA>          9 <NULL> <NA>          3 <NULL> <NA>   <NULL> <NA>   
## 15        17 <NA>         10 <NULL> <NA>          4 <NULL> <NA>   <NULL> <NA>   
## 16        18 <NA>         11 <NULL> <NA>          5 <NULL> <NA>   <NULL> <NA>   
## 17        19 <NA>         NA <NULL> <NA>         NA <NULL> <NA>   <NULL> <NA>   
## # … with 4 more variables: `11(K)` <list>, `12(L)` <chr>, `13(M)` <list>,
## #   `14(N)` <chr>
```

```r
## Formating layer
cx.d |>
  rectify(value = local_format_id)
```

```
## # A tibble: 17 × 14
##    `row/col` `2(B)` `3(C)` `4(D)` `5(E)` `6(F)` `7(G)` `8(H)` `9(I)` `10(J)`
##        <int>  <int>  <int>  <int>  <int>  <int>  <int>  <int>  <int>   <int>
##  1         3     NA     NA     NA     NA     NA     NA     NA     NA      NA
##  2         4     NA     NA     NA     NA     NA     NA     NA     NA      NA
##  3         5     NA     NA     NA     NA     NA     NA     NA     NA      16
##  4         6     NA     NA     NA     NA     NA     NA     25     NA      17
##  5         7     NA     NA     NA     NA     NA     NA     25     NA      17
##  6         8     13     12     NA     19      1     NA     25     NA      17
##  7         9     14     12     NA     20      1     NA     25     NA      17
##  8        10     14     12     NA     20      1     NA     25     NA      17
##  9        11     14     12     NA     21      1     NA     25     NA      17
## 10        12     14     12     NA     NA     NA     NA     25     NA      17
## 11        13     14     12     NA     NA     NA     NA     25     NA      17
## 12        14     14     12     NA     22      1     NA     25     NA      17
## 13        15     14     12     NA     23      1     NA     25     NA      17
## 14        16     14     12     NA     23      1     NA     25     NA      18
## 15        17     14     12     NA     23      1     NA     25     NA      NA
## 16        18     15     12     NA     24      1     NA     NA     NA      NA
## 17        19     NA     12     NA     NA     NA     NA     NA     NA      NA
## # … with 4 more variables: `11(K)` <int>, `12(L)` <int>, `13(M)` <int>,
## #   `14(N)` <int>
```

```r
fx <- paste0(getwd(), "/data/Colored Excel.xlsx") |>
  xlsx_formats()
```
Just saying that even though the merged cells looks as if there were nothing but still underneath is formatting. 

It is not because it can recognise merged cells, but because your merged cells has borders..... 

Boarder is identified formate

# Update: use data tree to visualise which data
### Explore formate list 

```r
#install.packages("data.tree")
#library("data.tree")
fx |> map(names)
```

```
## $local
## [1] "numFmt"     "font"       "fill"       "border"     "alignment" 
## [6] "protection"
## 
## $style
## [1] "numFmt"     "font"       "fill"       "border"     "alignment" 
## [6] "protection"
```

```r
fx$local$border
```

```
## $diagonalDown
##  [1] FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE
## [13] FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE
## [25] FALSE
## 
## $diagonalUp
##  [1] FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE
## [13] FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE
## [25] FALSE
## 
## $outline
##  [1] FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE
## [13] FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE FALSE
## [25] FALSE
## 
## $left
## $left$style
##  [1] NA       NA       NA       NA       NA       NA       NA       NA      
##  [9] NA       NA       NA       NA       "medium" "medium" "medium" "medium"
## [17] "medium" "medium" "medium" "medium" "medium" "medium" "medium" "medium"
## [25] NA      
## 
## $left$color
## $left$color$rgb
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [19] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [25] NA        
## 
## $left$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $left$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA 65 65 65 65 65 65 65 65 65 65 65 65 NA
## 
## $left$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $right
## $right$style
##  [1] NA       NA       NA       NA       NA       NA       NA       NA      
##  [9] NA       NA       NA       NA       "medium" "medium" "medium" "medium"
## [17] "medium" "medium" "medium" "medium" "medium" "medium" "medium" "medium"
## [25] NA      
## 
## $right$color
## $right$color$rgb
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [19] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [25] NA        
## 
## $right$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $right$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA 65 65 65 65 65 65 65 65 65 65 65 65 NA
## 
## $right$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $start
## $start$style
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $start$color
## $start$color$rgb
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $start$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $start$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $start$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $end
## $end$style
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $end$color
## $end$color$rgb
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $end$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $end$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $end$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $top
## $top$style
##  [1] NA       NA       NA       NA       NA       NA       NA       NA      
##  [9] NA       NA       NA       NA       "medium" NA       NA       "medium"
## [17] NA       NA       "medium" NA       NA       "medium" NA       NA      
## [25] NA      
## 
## $top$color
## $top$color$rgb
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] "FFFFFFFF" NA         NA         "FFFFFFFF" NA         NA        
## [19] "FFFFFFFF" NA         NA         "FFFFFFFF" NA         NA        
## [25] NA        
## 
## $top$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $top$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA 65 NA NA 65 NA NA 65 NA NA 65 NA NA NA
## 
## $top$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $bottom
## $bottom$style
##  [1] NA       NA       NA       NA       NA       NA       NA       NA      
##  [9] NA       NA       NA       NA       NA       NA       "medium" NA      
## [17] NA       "medium" NA       NA       "medium" NA       NA       "medium"
## [25] NA      
## 
## $bottom$color
## $bottom$color$rgb
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] NA         NA         "FFFFFFFF" NA         NA         "FFFFFFFF"
## [19] NA         NA         "FFFFFFFF" NA         NA         "FFFFFFFF"
## [25] NA        
## 
## $bottom$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $bottom$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA 65 NA NA 65 NA NA 65 NA NA 65 NA
## 
## $bottom$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $diagonal
## $diagonal$style
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $diagonal$color
## $diagonal$color$rgb
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $diagonal$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $diagonal$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $diagonal$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $vertical
## $vertical$style
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $vertical$color
## $vertical$color$rgb
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $vertical$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $vertical$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $vertical$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## 
## 
## $horizontal
## $horizontal$style
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $horizontal$color
## $horizontal$color$rgb
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $horizontal$color$theme
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $horizontal$color$indexed
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
## 
## $horizontal$color$tint
##  [1] NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA NA
```

```r
fx$local$border$right$color$rgb
```

```
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [19] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [25] NA
```

```r
fx$local$border$top$color$rgb
```

```
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] "FFFFFFFF" NA         NA         "FFFFFFFF" NA         NA        
## [19] "FFFFFFFF" NA         NA         "FFFFFFFF" NA         NA        
## [25] NA
```

```r
fx$local$border$left$color$rgb
```

```
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [19] "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF" "FFFFFFFF"
## [25] NA
```

```r
fx$local$border$bottom$color$rgb
```

```
##  [1] NA         NA         NA         NA         NA         NA        
##  [7] NA         NA         NA         NA         NA         NA        
## [13] NA         NA         "FFFFFFFF" NA         NA         "FFFFFFFF"
## [19] NA         NA         "FFFFFFFF" NA         NA         "FFFFFFFF"
## [25] NA
```


## How to Visualise a Nested List? 
[stack Visualise object as tree](https://stackoverflow.com/questions/51608378/visualise-object-in-r-as-tree)


```r
x <- list(
  id = 1,
  status = "active",
  coord = list(phi=0, theta=1, r=1),
  mt = NULL
)
depth <- function(x) ifelse(is.list(x), 1 + max(sapply(x, depth)), 0)

toTree <- function(x) {
  d <- depth(x)
  if(d > 1) {
    lapply(x, toTree)
  } else {
    children = lapply(names(x), function(nm) list(name=nm))
  }
}
```




```r
test_reg <- function(re){
  str_view_all(
    "style.border.diagonal.color.theme.Normal
    local.font.strike
    local.font.color.indexed", re)
}
test_reg("(?<=\\.)\\w+")
```

```{=html}
<div id="htmlwidget-aa6b719cf2a5780c8ffb" style="width:960px;height:100%;" class="str_view html-widget"></div>
<script type="application/json" data-for="htmlwidget-aa6b719cf2a5780c8ffb">{"x":{"html":"<ul>\n  <li>style.<span class='match'>border<\/span>.<span class='match'>diagonal<\/span>.<span class='match'>color<\/span>.<span class='match'>theme<\/span>.<span class='match'>Normal<\/span>\n    local.<span class='match'>font<\/span>.<span class='match'>strike<\/span>\n    local.<span class='match'>font<\/span>.<span class='match'>color<\/span>.<span class='match'>indexed<\/span><\/li>\n<\/ul>"},"evals":[],"jsHooks":[]}</script>
```

```r
test_reg("\\w+\\.\\w+")
```

```{=html}
<div id="htmlwidget-1de3987961bf0bd97edb" style="width:960px;height:100%;" class="str_view html-widget"></div>
<script type="application/json" data-for="htmlwidget-1de3987961bf0bd97edb">{"x":{"html":"<ul>\n  <li><span class='match'>style.border<\/span>.<span class='match'>diagonal.color<\/span>.<span class='match'>theme.Normal<\/span>\n    <span class='match'>local.font<\/span>.strike\n    <span class='match'>local.font<\/span>.<span class='match'>color.indexed<\/span><\/li>\n<\/ul>"},"evals":[],"jsHooks":[]}</script>
```

```r
test_reg("(?<=\\.)\\w+\\.\\w+(?=\\.)")
```

```{=html}
<div id="htmlwidget-bb072e80029137c3a91b" style="width:960px;height:100%;" class="str_view html-widget"></div>
<script type="application/json" data-for="htmlwidget-bb072e80029137c3a91b">{"x":{"html":"<ul>\n  <li>style.<span class='match'>border.diagonal<\/span>.<span class='match'>color.theme<\/span>.Normal\n    local.font.strike\n    local.<span class='match'>font.color<\/span>.indexed<\/li>\n<\/ul>"},"evals":[],"jsHooks":[]}</script>
```
# Give in to `openxl`

```r
#install.packages("openxlsx")
library(openxlsx)

wb<- list.files("data", "Colored", full.names = T) |>
  loadWorkbook()

getStyles(wb) |> class()
```

```
## [1] "list"
```

```r
wbStyle <- getStyles(wb)
#wbStyle |> map(names)
wbStyle |> length()
```

```
## [1] 27
```

```r
wbStyle |> depth()
```

```
## [1] 1
```


```r
# NOT RUN {
## Create a new workbook
wb <- createWorkbook()

## Add a worksheet
addWorksheet(wb, "Sheet 1")
addWorksheet(wb, "Sheet 2")

## Merge cells: Row 2 column C to F (3:6)
mergeCells(wb, "Sheet 1", cols = 2, rows = 3:6)

## Merge cells:Rows 10 to 20 columns A to J (1:10)
mergeCells(wb, 1, cols = 1:10, rows = 10:20)

## Intersecting merges
mergeCells(wb, 2, cols = 1:10, rows = 1)
mergeCells(wb, 2, cols = 5:10, rows = 2)
mergeCells(wb, 2, cols = c(1, 10), rows = 12) ## equivalent to 1:10 as only min/max are used
# mergeCells(wb, 2, cols = 1, rows = c(1,10)) # Throws error because intersects existing merge

## remove merged cells
removeCellMerge(wb, 2, cols = 1, rows = 1) # removes any intersecting merges
mergeCells(wb, 2, cols = 1, rows = 1:10) # Now this works

## Save workbook
# }
# NOT RUN {
saveWorkbook(wb, "data/mergeCellsExample.xlsx", overwrite = TRUE)
# }
```

# Ideas for production
Okayy...this seems fun... just I can not possible find a "perfect" solution. 
I guess there were two ways to experiment with: 
  a. use the neighboring values to determine if a 
  b. format 
    b.a general format
    b.b boarder forma - which is most likely to be unreliable. 
  
A little CBT extraction idea: 
find where splitter is, then find row, in the same row, find maximum column that are not NULL. Then you locate where our CBT is!!

To find CBT port value... just look around neighbor and `length()` not Blank. or literally just `col - 1`


Extracting FDPs links should be much easier??? As done `behead(. "right-up", "outgoing_fdp)` OMG this is OP!! 
We just need a bit analytic around where to cut, it is li~terally just `filter()`

My data looks like this:
```
___________________
ENCLOSURE fdp
-------------------
Subs       |   |  |
xxxx       |   |  |
-------------------
(a lot  
    of 
blank space)
--------
fdp|   |------------X----------------------------------------> | X |~ cbt
   |   |------------X------------------------> | X |~ cbt      |   |
   |   |------c------> |fdp|   | #             |   |           |   |~ cbt 
   |   |------c------> |   |   | #             |   |.          |   |
   |   |.                                      |   |.
   |   |.
   |   |.
   |   |.
   |   |.
--------
```
If we consider spine node and edge all together is totally fun ~ we visualize things as it is... it is just going to have much-root. 

## Do merged cells have special formate? 
`t()` and `flatten()` really helps a lot to identify which part of the code is different. 

```r
fx |> flatten() |> data.frame() -> ftb
```

```
## Warning in (function (..., row.names = NULL, check.rows = FALSE, check.names =
## TRUE, : row names were found from a short variable and have been discarded
```

```r
  cx.c |> rectify(value = local_format_id)
```

```
## # A tibble: 18 × 3
##    `row/col` `3(C)` `4(D)`
##        <int>  <int>  <int>
##  1         7     13      1
##  2         8     14      1
##  3         9     14      1
##  4        10     14      1
##  5        11     14      1
##  6        12     14      1
##  7        13     14      1
##  8        14     14      1
##  9        15     14      1
## 10        16     14      1
## 11        17     14      1
## 12        18     15      1
## 13        19     NA     NA
## 14        20     NA     NA
## 15        21     13      1
## 16        22     14      1
## 17        23     14      1
## 18        24     15      1
```

```r
cx.d |> rectify(value = local_format_id)
```

```
## # A tibble: 17 × 14
##    `row/col` `2(B)` `3(C)` `4(D)` `5(E)` `6(F)` `7(G)` `8(H)` `9(I)` `10(J)`
##        <int>  <int>  <int>  <int>  <int>  <int>  <int>  <int>  <int>   <int>
##  1         3     NA     NA     NA     NA     NA     NA     NA     NA      NA
##  2         4     NA     NA     NA     NA     NA     NA     NA     NA      NA
##  3         5     NA     NA     NA     NA     NA     NA     NA     NA      16
##  4         6     NA     NA     NA     NA     NA     NA     25     NA      17
##  5         7     NA     NA     NA     NA     NA     NA     25     NA      17
##  6         8     13     12     NA     19      1     NA     25     NA      17
##  7         9     14     12     NA     20      1     NA     25     NA      17
##  8        10     14     12     NA     20      1     NA     25     NA      17
##  9        11     14     12     NA     21      1     NA     25     NA      17
## 10        12     14     12     NA     NA     NA     NA     25     NA      17
## 11        13     14     12     NA     NA     NA     NA     25     NA      17
## 12        14     14     12     NA     22      1     NA     25     NA      17
## 13        15     14     12     NA     23      1     NA     25     NA      17
## 14        16     14     12     NA     23      1     NA     25     NA      18
## 15        17     14     12     NA     23      1     NA     25     NA      NA
## 16        18     15     12     NA     24      1     NA     NA     NA      NA
## 17        19     NA     12     NA     NA     NA     NA     NA     NA      NA
## # … with 4 more variables: `11(K)` <int>, `12(L)` <int>, `13(M)` <int>,
## #   `14(N)` <int>
```

```r
c(12, 13, 14) %in% 1:nrow(ftb)
```

```
## [1] TRUE TRUE TRUE
```

```r
ftb[1:nrow(ftb) %in% c(12, 13, 14, 25), ] |> 
  t() |>
  data.frame()
```

```
##                                              X12      X13      X14      X25
## numFmt                                   General  General  General  General
## font.bold                                  FALSE    FALSE    FALSE    FALSE
## font.italic                                FALSE    FALSE    FALSE    FALSE
## font.underline                              <NA>     <NA>     <NA>     <NA>
## font.strike                                FALSE    FALSE    FALSE    FALSE
## font.vertAlign                              <NA>     <NA>     <NA>     <NA>
## font.size                                     12       12       12       12
## font.color.rgb                          FF000000 FF000000 FF000000 FF000000
## font.color.theme                           text1    text1    text1    text1
## font.color.indexed                          <NA>     <NA>     <NA>     <NA>
## font.color.tint                             <NA>     <NA>     <NA>     <NA>
## font.name                                Calibri  Calibri  Calibri  Calibri
## font.family                                    2        2        2        2
## font.scheme                                minor    minor    minor    minor
## fill.patternFill.fgColor.rgb                <NA>     <NA>     <NA>     <NA>
## fill.patternFill.fgColor.theme              <NA>     <NA>     <NA>     <NA>
## fill.patternFill.fgColor.indexed            <NA>     <NA>     <NA>     <NA>
## fill.patternFill.fgColor.tint               <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.rgb                <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.theme              <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.indexed            <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.tint               <NA>     <NA>     <NA>     <NA>
## fill.patternFill.patternType                <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.type                      <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.degree                    <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.left                      <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.right                     <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.top                       <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.bottom                    <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.position            <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.rgb           <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.theme         <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.indexed       <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.tint          <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.position            <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.rgb           <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.theme         <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.indexed       <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.tint          <NA>     <NA>     <NA>     <NA>
## border.diagonalDown                        FALSE    FALSE    FALSE    FALSE
## border.diagonalUp                          FALSE    FALSE    FALSE    FALSE
## border.outline                             FALSE    FALSE    FALSE    FALSE
## border.left.style                           <NA>   medium   medium     <NA>
## border.left.color.rgb                       <NA> FFFFFFFF FFFFFFFF     <NA>
## border.left.color.theme                     <NA>     <NA>     <NA>     <NA>
## border.left.color.indexed                   <NA>       65       65     <NA>
## border.left.color.tint                      <NA>     <NA>     <NA>     <NA>
## border.right.style                          <NA>   medium   medium     <NA>
## border.right.color.rgb                      <NA> FFFFFFFF FFFFFFFF     <NA>
## border.right.color.theme                    <NA>     <NA>     <NA>     <NA>
## border.right.color.indexed                  <NA>       65       65     <NA>
## border.right.color.tint                     <NA>     <NA>     <NA>     <NA>
## border.start.style                          <NA>     <NA>     <NA>     <NA>
## border.start.color.rgb                      <NA>     <NA>     <NA>     <NA>
## border.start.color.theme                    <NA>     <NA>     <NA>     <NA>
## border.start.color.indexed                  <NA>     <NA>     <NA>     <NA>
## border.start.color.tint                     <NA>     <NA>     <NA>     <NA>
## border.end.style                            <NA>     <NA>     <NA>     <NA>
## border.end.color.rgb                        <NA>     <NA>     <NA>     <NA>
## border.end.color.theme                      <NA>     <NA>     <NA>     <NA>
## border.end.color.indexed                    <NA>     <NA>     <NA>     <NA>
## border.end.color.tint                       <NA>     <NA>     <NA>     <NA>
## border.top.style                            <NA>   medium     <NA>     <NA>
## border.top.color.rgb                        <NA> FFFFFFFF     <NA>     <NA>
## border.top.color.theme                      <NA>     <NA>     <NA>     <NA>
## border.top.color.indexed                    <NA>       65     <NA>     <NA>
## border.top.color.tint                       <NA>     <NA>     <NA>     <NA>
## border.bottom.style                         <NA>     <NA>     <NA>     <NA>
## border.bottom.color.rgb                     <NA>     <NA>     <NA>     <NA>
## border.bottom.color.theme                   <NA>     <NA>     <NA>     <NA>
## border.bottom.color.indexed                 <NA>     <NA>     <NA>     <NA>
## border.bottom.color.tint                    <NA>     <NA>     <NA>     <NA>
## border.diagonal.style                       <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.rgb                   <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.theme                 <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.indexed               <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.tint                  <NA>     <NA>     <NA>     <NA>
## border.vertical.style                       <NA>     <NA>     <NA>     <NA>
## border.vertical.color.rgb                   <NA>     <NA>     <NA>     <NA>
## border.vertical.color.theme                 <NA>     <NA>     <NA>     <NA>
## border.vertical.color.indexed               <NA>     <NA>     <NA>     <NA>
## border.vertical.color.tint                  <NA>     <NA>     <NA>     <NA>
## border.horizontal.style                     <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.rgb                 <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.theme               <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.indexed             <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.tint                <NA>     <NA>     <NA>     <NA>
## alignment.horizontal                     general   center   center   center
## alignment.vertical                        bottom   center   center   center
## alignment.wrapText                         FALSE    FALSE    FALSE    FALSE
## alignment.readingOrder                   context  context  context  context
## alignment.indent                               0        0        0        0
## alignment.justifyLastLine                  FALSE    FALSE    FALSE    FALSE
## alignment.shrinkToFit                      FALSE    FALSE    FALSE    FALSE
## alignment.textRotation                         0       90       90       90
## protection.locked                           TRUE     TRUE     TRUE     TRUE
## protection.hidden                          FALSE    FALSE    FALSE    FALSE
## numFmt.1                                 General  General  General  General
## font.bold.1                                FALSE    FALSE    FALSE    FALSE
## font.italic.1                              FALSE    FALSE    FALSE    FALSE
## font.underline.1                            <NA>     <NA>     <NA>     <NA>
## font.strike.1                              FALSE    FALSE    FALSE    FALSE
## font.vertAlign.1                            <NA>     <NA>     <NA>     <NA>
## font.size.1                                   12       12       12       12
## font.color.rgb.1                        FF000000 FF000000 FF000000 FF000000
## font.color.theme.1                         text1    text1    text1    text1
## font.color.indexed.1                        <NA>     <NA>     <NA>     <NA>
## font.color.tint.1                           <NA>     <NA>     <NA>     <NA>
## font.name.1                              Calibri  Calibri  Calibri  Calibri
## font.family.1                                  2        2        2        2
## font.scheme.1                              minor    minor    minor    minor
## fill.patternFill.fgColor.rgb.1              <NA>     <NA>     <NA>     <NA>
## fill.patternFill.fgColor.theme.1            <NA>     <NA>     <NA>     <NA>
## fill.patternFill.fgColor.indexed.1          <NA>     <NA>     <NA>     <NA>
## fill.patternFill.fgColor.tint.1             <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.rgb.1              <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.theme.1            <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.indexed.1          <NA>     <NA>     <NA>     <NA>
## fill.patternFill.bgColor.tint.1             <NA>     <NA>     <NA>     <NA>
## fill.patternFill.patternType.1              <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.type.1                    <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.degree.1                  <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.left.1                    <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.right.1                   <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.top.1                     <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.bottom.1                  <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.position.1          <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.rgb.1         <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.theme.1       <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.indexed.1     <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop1.color.tint.1        <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.position.1          <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.rgb.1         <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.theme.1       <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.indexed.1     <NA>     <NA>     <NA>     <NA>
## fill.gradientFill.stop2.color.tint.1        <NA>     <NA>     <NA>     <NA>
## border.diagonalDown.1                      FALSE    FALSE    FALSE    FALSE
## border.diagonalUp.1                        FALSE    FALSE    FALSE    FALSE
## border.outline.1                           FALSE    FALSE    FALSE    FALSE
## border.left.style.1                         <NA>     <NA>     <NA>     <NA>
## border.left.color.rgb.1                     <NA>     <NA>     <NA>     <NA>
## border.left.color.theme.1                   <NA>     <NA>     <NA>     <NA>
## border.left.color.indexed.1                 <NA>     <NA>     <NA>     <NA>
## border.left.color.tint.1                    <NA>     <NA>     <NA>     <NA>
## border.right.style.1                        <NA>     <NA>     <NA>     <NA>
## border.right.color.rgb.1                    <NA>     <NA>     <NA>     <NA>
## border.right.color.theme.1                  <NA>     <NA>     <NA>     <NA>
## border.right.color.indexed.1                <NA>     <NA>     <NA>     <NA>
## border.right.color.tint.1                   <NA>     <NA>     <NA>     <NA>
## border.start.style.1                        <NA>     <NA>     <NA>     <NA>
## border.start.color.rgb.1                    <NA>     <NA>     <NA>     <NA>
## border.start.color.theme.1                  <NA>     <NA>     <NA>     <NA>
## border.start.color.indexed.1                <NA>     <NA>     <NA>     <NA>
## border.start.color.tint.1                   <NA>     <NA>     <NA>     <NA>
## border.end.style.1                          <NA>     <NA>     <NA>     <NA>
## border.end.color.rgb.1                      <NA>     <NA>     <NA>     <NA>
## border.end.color.theme.1                    <NA>     <NA>     <NA>     <NA>
## border.end.color.indexed.1                  <NA>     <NA>     <NA>     <NA>
## border.end.color.tint.1                     <NA>     <NA>     <NA>     <NA>
## border.top.style.1                          <NA>     <NA>     <NA>     <NA>
## border.top.color.rgb.1                      <NA>     <NA>     <NA>     <NA>
## border.top.color.theme.1                    <NA>     <NA>     <NA>     <NA>
## border.top.color.indexed.1                  <NA>     <NA>     <NA>     <NA>
## border.top.color.tint.1                     <NA>     <NA>     <NA>     <NA>
## border.bottom.style.1                       <NA>     <NA>     <NA>     <NA>
## border.bottom.color.rgb.1                   <NA>     <NA>     <NA>     <NA>
## border.bottom.color.theme.1                 <NA>     <NA>     <NA>     <NA>
## border.bottom.color.indexed.1               <NA>     <NA>     <NA>     <NA>
## border.bottom.color.tint.1                  <NA>     <NA>     <NA>     <NA>
## border.diagonal.style.1                     <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.rgb.1                 <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.theme.1               <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.indexed.1             <NA>     <NA>     <NA>     <NA>
## border.diagonal.color.tint.1                <NA>     <NA>     <NA>     <NA>
## border.vertical.style.1                     <NA>     <NA>     <NA>     <NA>
## border.vertical.color.rgb.1                 <NA>     <NA>     <NA>     <NA>
## border.vertical.color.theme.1               <NA>     <NA>     <NA>     <NA>
## border.vertical.color.indexed.1             <NA>     <NA>     <NA>     <NA>
## border.vertical.color.tint.1                <NA>     <NA>     <NA>     <NA>
## border.horizontal.style.1                   <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.rgb.1               <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.theme.1             <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.indexed.1           <NA>     <NA>     <NA>     <NA>
## border.horizontal.color.tint.1              <NA>     <NA>     <NA>     <NA>
## alignment.horizontal.1                   general  general  general  general
## alignment.vertical.1                      bottom   bottom   bottom   bottom
## alignment.wrapText.1                       FALSE    FALSE    FALSE    FALSE
## alignment.readingOrder.1                 context  context  context  context
## alignment.indent.1                             0        0        0        0
## alignment.justifyLastLine.1                FALSE    FALSE    FALSE    FALSE
## alignment.shrinkToFit.1                    FALSE    FALSE    FALSE    FALSE
## alignment.textRotation.1                       0        0        0        0
## protection.locked.1                         TRUE     TRUE     TRUE     TRUE
## protection.hidden.1                        FALSE    FALSE    FALSE    FALSE
```
So merged cells will share some same formate such as font, background color (exceptions are boarders)... however cells share teh same formate does not means they are mereged... make this unreliable way of identifying a mreged cell.

