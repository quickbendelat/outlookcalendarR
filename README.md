
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
# install.packages("rempotes")
remotes::install_github("quickbendelat/outlookcalendarR")
```

You will also need to install RDCOMClient.

For R \< 4.0:

``` r
remotes::install_github("omegahat/RDCOMClient")
```

As of R \>= 4.0, the app crashes due to an error with a function from
RDCOMClient. See
[issue](https://github.com/omegahat/RDCOMClient/issues/24) on the github
repo.

The current solution for R \> 4.0 is to download and install a specially
built binary:

``` r
dir <- tempdir()
zip <- file.path(dir, "RDCOMClient.zip")
url <- "https://github.com/dkyleward/RDCOMClient/releases/download/v0.94/RDCOMClient_binary.zip"
download.file(url, zip)
install.packages(zip, repos = NULL, type = "win.binary")
```

## Running

outlookcalendarR is run by:

``` r
outlookcalendarR::run_app()
```

## Connecting to MS Outlook

The package [RDCOMClient](https://github.com/omegahat/RDCOMClient)
provides the ability to connect to MS Outlook.

[This Microsoft
link](https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders)
lists all the enumeration values to connect to the different aspects of
Outlook.

To convert the COMDate returned by the date fields from Outlook, I found
the package [extrospectr](https://github.com/aecoleman/extrospectr/),
but I could not build the package. So I copied the function
`.COMDate_to_POSIX()` from extrospectr to use in outlookcalendarR.
