
<!-- README.md is generated from README.Rmd. Please edit that file -->

# outlookcalendarR

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://www.tidyverse.org/lifecycle/#experimental)
<!-- badges: end -->

outlookcalendarR is a Shiny App that presents information about a users
meetings in their Microsoft Outlook Calendar.

This will only work on a Windows machine with Microsoft Outlook.

## Installation

You can install outlookcalendarR from [GitHub](https://github.com/)
with:

``` r
# install.packages("devtools")
devtools::install_github("quickbendelat/outlookcalendarR")

# or

remotes::install_github("quickbendelat/outlookcalendarR")
```

## Running

outlookcalendarR is run by:

``` r
outlookcalendarR::run_app()
```

## Connecting to MS Outlook

The package
[excel.link](https://cran.r-project.org/web/packages/excel.link/index.html)
provides the ability to connect to MS Outlook.

[This Microsoft
link](https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders)
lists all the enumeration values to connect to the different aspects of
Outlook.

To convert the COMDate returned by the date fields from Outlook, I found
the package [extrospectr](https://github.com/aecoleman/extrospectr/),
but I could not build the package. So I copied the function
`.COMDate_to_POSIX()` from extrospectr to use in outlookcalendarR.