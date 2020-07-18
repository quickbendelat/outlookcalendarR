#' COMDate to POSIX
#' from https://github.com/aecoleman/extrospectr/
#'
#' @param x COMDate object
#'
#' @return POSIXct
#'
.COMDate_to_POSIX <- function(x) {
  
  stopifnot('COMDate' %in% class(x))
  
  x %>% purrr::map_dbl( ~ .x) %>% openxlsx::convertToDateTime()
  
}