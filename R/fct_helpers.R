#' COMDate to POSIX
#' from https://github.com/aecoleman/extrospectr/
#'
#' @param x COMDate object
#'
#' @return POSIXct
#'
.COMDate_to_POSIX <- function(x) {
  
  stopifnot('COMDate' %in% class(x))
  
  x %>% purrr::map_dbl( ~ .x) %>% convertToDateTime()
  
}


#' find the user name
#'
#' @return me
#'
find_user <- function() {
  
  OutApp <- COMCreate("Outlook.Application")
  outlookNameSpace = OutApp$GetNameSpace("MAPI")
  
  sent_fld <- outlookNameSpace$GetDefaultFolder(5) # 5 is sent folder
  
  sent_emails <- sent_fld$items
  
  me <- sent_emails(1)[['SenderName']]
  
}