# Read 'Word' styles

Read Word styles and get results in a data.frame.

## Usage

``` r
styles_info(
  x,
  type = c("paragraph", "character", "table", "numbering"),
  is_default = c(TRUE, FALSE)
)
```

## Arguments

- x:

  an rdocx object

- type, is_default:

  subsets for types (i.e. paragraph) and default style (when
  `is_default` is TRUE or FALSE)

## See also

Other functions for Word document informations:
[`doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.md),
[`docx_bookmarks()`](https://davidgohel.github.io/officer/reference/docx_bookmarks.md),
[`docx_dim()`](https://davidgohel.github.io/officer/reference/docx_dim.md),
[`length.rdocx()`](https://davidgohel.github.io/officer/reference/length.rdocx.md),
[`set_doc_properties()`](https://davidgohel.github.io/officer/reference/set_doc_properties.md)

## Examples

``` r
x <- read_docx()
styles_info(x)
#>    style_type             style_id             style_name        base_on
#> 1   paragraph               Normal                 Normal           <NA>
#> 2   paragraph               Titre1              heading 1         Normal
#> 3   paragraph               Titre2              heading 2         Normal
#> 4   paragraph               Titre3              heading 3         Normal
#> 5   character       Policepardfaut Default Paragraph Font           <NA>
#> 6       table        TableauNormal           Normal Table           <NA>
#> 7   numbering          Aucuneliste                No List           <NA>
#> 8   character               strong                 strong Policepardfaut
#> 9   paragraph             centered               centered         Normal
#> 10      table        tabletemplate         table_template  TableauNormal
#> 11      table  Listeclaire-Accent2    Light List Accent 2  TableauNormal
#> 12  character            Titre1Car            Titre 1 Car Policepardfaut
#> 13  character            Titre2Car            Titre 2 Car Policepardfaut
#> 14  character            Titre3Car            Titre 3 Car Policepardfaut
#> 15  paragraph         ImageCaption          Image Caption         Normal
#> 16  paragraph         TableCaption          Table Caption   ImageCaption
#> 17      table Tableauprofessionnel     Table Professional  TableauNormal
#> 18  paragraph                  TM1                  toc 1         Normal
#> 19  paragraph                  TM2                  toc 2         Normal
#> 20  paragraph        Textedebulles           Balloon Text         Normal
#> 21  character     TextedebullesCar    Texte de bulles Car Policepardfaut
#> 22  character          referenceid           reference_id Policepardfaut
#> 23  paragraph         graphictitle          graphic title   ImageCaption
#> 24  paragraph           tabletitle            table title   TableCaption
#>    is_custom is_default  align keep_next line_spacing padding.bottom
#> 1      FALSE       TRUE   <NA>     FALSE           NA           <NA>
#> 2      FALSE      FALSE   <NA>      TRUE           NA           <NA>
#> 3      FALSE      FALSE   <NA>      TRUE           NA           <NA>
#> 4      FALSE      FALSE   <NA>      TRUE           NA           <NA>
#> 5      FALSE       TRUE   <NA>     FALSE           NA           <NA>
#> 6      FALSE       TRUE   <NA>     FALSE           NA           <NA>
#> 7      FALSE       TRUE   <NA>     FALSE           NA           <NA>
#> 8       TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 9       TRUE      FALSE center     FALSE           NA           <NA>
#> 10      TRUE      FALSE  right     FALSE           NA           <NA>
#> 11     FALSE      FALSE   <NA>     FALSE           NA           <NA>
#> 12      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 13      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 14      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 15      TRUE      FALSE center     FALSE           NA           <NA>
#> 16      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 17     FALSE      FALSE   <NA>     FALSE           NA           <NA>
#> 18     FALSE      FALSE   <NA>     FALSE           NA            100
#> 19     FALSE      FALSE   <NA>     FALSE           NA            100
#> 20     FALSE      FALSE   <NA>     FALSE           NA           <NA>
#> 21      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 22      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 23      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#> 24      TRUE      FALSE   <NA>     FALSE           NA           <NA>
#>    padding.top padding.left padding.right shading.color.par border.bottom.width
#> 1         <NA>         <NA>          <NA>              <NA>                  NA
#> 2          480         <NA>          <NA>              <NA>                 0.5
#> 3          200         <NA>          <NA>              <NA>                  NA
#> 4          200         <NA>          <NA>              <NA>                  NA
#> 5         <NA>         <NA>          <NA>              <NA>                  NA
#> 6         <NA>         <NA>          <NA>              <NA>                  NA
#> 7         <NA>         <NA>          <NA>              <NA>                  NA
#> 8         <NA>         <NA>          <NA>              <NA>                  NA
#> 9         <NA>         <NA>          <NA>              <NA>                  NA
#> 10        <NA>         <NA>          <NA>              <NA>                  NA
#> 11        <NA>         <NA>          <NA>              <NA>                  NA
#> 12        <NA>         <NA>          <NA>              <NA>                  NA
#> 13        <NA>         <NA>          <NA>              <NA>                  NA
#> 14        <NA>         <NA>          <NA>              <NA>                  NA
#> 15        <NA>         <NA>          <NA>              <NA>                  NA
#> 16        <NA>         <NA>          <NA>              <NA>                  NA
#> 17        <NA>         <NA>          <NA>              <NA>                  NA
#> 18        <NA>         <NA>          <NA>              <NA>                  NA
#> 19        <NA>          240          <NA>              <NA>                  NA
#> 20        <NA>         <NA>          <NA>              <NA>                  NA
#> 21        <NA>         <NA>          <NA>              <NA>                  NA
#> 22        <NA>         <NA>          <NA>              <NA>                  NA
#> 23        <NA>         <NA>          <NA>              <NA>                  NA
#> 24        <NA>         <NA>          <NA>              <NA>                  NA
#>    border.bottom.color border.bottom.style border.top.width border.top.color
#> 1                 <NA>                <NA>               NA             <NA>
#> 2                 auto              single               NA             <NA>
#> 3                 <NA>                <NA>               NA             <NA>
#> 4                 <NA>                <NA>               NA             <NA>
#> 5                 <NA>                <NA>               NA             <NA>
#> 6                 <NA>                <NA>               NA             <NA>
#> 7                 <NA>                <NA>               NA             <NA>
#> 8                 <NA>                <NA>               NA             <NA>
#> 9                 <NA>                <NA>               NA             <NA>
#> 10                <NA>                <NA>               NA             <NA>
#> 11                <NA>                <NA>               NA             <NA>
#> 12                <NA>                <NA>               NA             <NA>
#> 13                <NA>                <NA>               NA             <NA>
#> 14                <NA>                <NA>               NA             <NA>
#> 15                <NA>                <NA>               NA             <NA>
#> 16                <NA>                <NA>               NA             <NA>
#> 17                <NA>                <NA>               NA             <NA>
#> 18                <NA>                <NA>               NA             <NA>
#> 19                <NA>                <NA>               NA             <NA>
#> 20                <NA>                <NA>               NA             <NA>
#> 21                <NA>                <NA>               NA             <NA>
#> 22                <NA>                <NA>               NA             <NA>
#> 23                <NA>                <NA>               NA             <NA>
#> 24                <NA>                <NA>               NA             <NA>
#>    border.top.style border.left.width border.left.color border.left.style
#> 1              <NA>                NA              <NA>              <NA>
#> 2              <NA>                NA              <NA>              <NA>
#> 3              <NA>                NA              <NA>              <NA>
#> 4              <NA>                NA              <NA>              <NA>
#> 5              <NA>                NA              <NA>              <NA>
#> 6              <NA>                NA              <NA>              <NA>
#> 7              <NA>                NA              <NA>              <NA>
#> 8              <NA>                NA              <NA>              <NA>
#> 9              <NA>                NA              <NA>              <NA>
#> 10             <NA>                NA              <NA>              <NA>
#> 11             <NA>                NA              <NA>              <NA>
#> 12             <NA>                NA              <NA>              <NA>
#> 13             <NA>                NA              <NA>              <NA>
#> 14             <NA>                NA              <NA>              <NA>
#> 15             <NA>                NA              <NA>              <NA>
#> 16             <NA>                NA              <NA>              <NA>
#> 17             <NA>                NA              <NA>              <NA>
#> 18             <NA>                NA              <NA>              <NA>
#> 19             <NA>                NA              <NA>              <NA>
#> 20             <NA>                NA              <NA>              <NA>
#> 21             <NA>                NA              <NA>              <NA>
#> 22             <NA>                NA              <NA>              <NA>
#> 23             <NA>                NA              <NA>              <NA>
#> 24             <NA>                NA              <NA>              <NA>
#>    border.right.width border.right.color border.right.style font.size bold
#> 1                  NA               <NA>               <NA>      <NA> <NA>
#> 2                  NA               <NA>               <NA>        32 <NA>
#> 3                  NA               <NA>               <NA>        26 <NA>
#> 4                  NA               <NA>               <NA>      <NA> <NA>
#> 5                  NA               <NA>               <NA>      <NA> <NA>
#> 6                  NA               <NA>               <NA>      <NA> <NA>
#> 7                  NA               <NA>               <NA>      <NA> <NA>
#> 8                  NA               <NA>               <NA>      <NA> <NA>
#> 9                  NA               <NA>               <NA>      <NA> <NA>
#> 10                 NA               <NA>               <NA>      <NA> <NA>
#> 11                 NA               <NA>               <NA>      <NA> <NA>
#> 12                 NA               <NA>               <NA>        32 <NA>
#> 13                 NA               <NA>               <NA>        26 <NA>
#> 14                 NA               <NA>               <NA>      <NA> <NA>
#> 15                 NA               <NA>               <NA>      <NA> <NA>
#> 16                 NA               <NA>               <NA>      <NA> <NA>
#> 17                 NA               <NA>               <NA>      <NA> <NA>
#> 18                 NA               <NA>               <NA>      <NA> <NA>
#> 19                 NA               <NA>               <NA>      <NA> <NA>
#> 20                 NA               <NA>               <NA>        18 <NA>
#> 21                 NA               <NA>               <NA>        18 <NA>
#> 22                 NA               <NA>               <NA>      <NA> <NA>
#> 23                 NA               <NA>               <NA>      <NA> <NA>
#> 24                 NA               <NA>               <NA>      <NA> <NA>
#>    italic underlined color   font.family vertical.align shading.color
#> 1    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 2    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 3    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 4    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 5    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 6    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 7    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 8    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 9    <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 10   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 11   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 12   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 13   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 14   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 15   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 16   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 17   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 18   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 19   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 20   <NA>       <NA>  <NA> Lucida Grande           <NA>          <NA>
#> 21   <NA>       <NA>  <NA> Lucida Grande           <NA>          <NA>
#> 22   <NA>       <NA>  <NA>          <NA>    superscript          <NA>
#> 23   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#> 24   <NA>       <NA>  <NA>          <NA>           <NA>          <NA>
#>     hansi.family eastasia.family cs.family bold.cs font.size.cs lang.val
#> 1           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 2           <NA>            <NA>      <NA>    <NA>           32     <NA>
#> 3           <NA>            <NA>      <NA>    <NA>           26     <NA>
#> 4           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 5           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 6           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 7           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 8           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 9           <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 10          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 11          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 12          <NA>            <NA>      <NA>    <NA>           32     <NA>
#> 13          <NA>            <NA>      <NA>    <NA>           26     <NA>
#> 14          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 15          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 16          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 17          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 18          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 19          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 20 Lucida Grande            <NA>      <NA>    <NA>           18     <NA>
#> 21 Lucida Grande            <NA>      <NA>    <NA>           18     <NA>
#> 22          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 23          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#> 24          <NA>            <NA>      <NA>    <NA>         <NA>     <NA>
#>    lang.eastasia lang.bidi
#> 1           <NA>      <NA>
#> 2           <NA>      <NA>
#> 3           <NA>      <NA>
#> 4           <NA>      <NA>
#> 5           <NA>      <NA>
#> 6           <NA>      <NA>
#> 7           <NA>      <NA>
#> 8           <NA>      <NA>
#> 9           <NA>      <NA>
#> 10          <NA>      <NA>
#> 11          <NA>      <NA>
#> 12          <NA>      <NA>
#> 13          <NA>      <NA>
#> 14          <NA>      <NA>
#> 15          <NA>      <NA>
#> 16          <NA>      <NA>
#> 17          <NA>      <NA>
#> 18          <NA>      <NA>
#> 19          <NA>      <NA>
#> 20          <NA>      <NA>
#> 21          <NA>      <NA>
#> 22          <NA>      <NA>
#> 23          <NA>      <NA>
#> 24          <NA>      <NA>
styles_info(x, type = "paragraph", is_default = TRUE)
#>   style_type style_id style_name base_on is_custom is_default align keep_next
#> 1  paragraph   Normal     Normal    <NA>     FALSE       TRUE  <NA>     FALSE
#>   line_spacing padding.bottom padding.top padding.left padding.right
#> 1           NA           <NA>        <NA>         <NA>          <NA>
#>   shading.color.par border.bottom.width border.bottom.color border.bottom.style
#> 1              <NA>                  NA                <NA>                <NA>
#>   border.top.width border.top.color border.top.style border.left.width
#> 1               NA             <NA>             <NA>                NA
#>   border.left.color border.left.style border.right.width border.right.color
#> 1              <NA>              <NA>                 NA               <NA>
#>   border.right.style font.size bold italic underlined color font.family
#> 1               <NA>      <NA> <NA>   <NA>       <NA>  <NA>        <NA>
#>   vertical.align shading.color hansi.family eastasia.family cs.family bold.cs
#> 1           <NA>          <NA>         <NA>            <NA>      <NA>    <NA>
#>   font.size.cs lang.val lang.eastasia lang.bidi
#> 1         <NA>     <NA>          <NA>      <NA>
```
