# Module UI

#' @title   mod_outlookcalendar_ui and mod_outlookcalendar_server
#' @description  A shiny Module.
#'
#' @param id shiny id
#' @param input internal
#' @param output internal
#' @param session internal
#'
#' @rdname mod_my_first_module
#'
#' @keywords internal
#' @export 
#' @importFrom shiny NS tagList 
mod_outlookcalendar_ui <- function(id){
  ns <- NS(id)
  tagList(
    verbatimTextOutput(ns("calendar_meetings"))
  )
}

# Module Server

#' @rdname mod_my_first_module
#' @export
#' @keywords internal
#' @import RDCOMClient
#' @import dplyr
#' @import tidyr
#' @import purrr
#' @import tibble
#' @import stringr
#' @import openxlsx

mod_outlookcalendar_server <- function(input, output, session){
  ns <- session$ns
  
  ## COMDate to POSIX
  ## https://github.com/aecoleman/extrospectr/
  .COMDate_to_POSIX <- function(x) {
    
    stopifnot('COMDate' %in% class(x))
    
    x %>% purrr::map_dbl( ~ .x) %>% convertToDateTime()
    
  }
  
  OutApp <- COMCreate("Outlook.Application")
  outlookNameSpace = OutApp$GetNameSpace("MAPI")
  sent_fld <- outlookNameSpace$GetDefaultFolder(5) # 5 is sent folder
  sent_emails <- sent_fld$items
  me <- sent_emails(1)[['SenderName']]
  
  calendar <- outlookNameSpace$GetDefaultFolder(9) #9 is calendar
  Cnt = calendar$Items()$Count()
  meetings <- calendar$items
  df <- seq(1:Cnt) %>% 
    tibble::enframe(name = NULL, value = "record")
  
  
  calendar_meetings <- df %>% 
    rowwise() %>% 
    mutate(title = meetings(record)[['subject']]) %>% 
    filter(title != "") %>% 
    mutate(sender = meetings(record)[['Organizer']],
           start = .COMDate_to_POSIX(meetings(record)[['start']]),
           end = .COMDate_to_POSIX(meetings(record)[['end']])) %>% 
    ungroup %>% 
    # remove where sender includes 4 hyphens which is in some hash thingy for loaded things like public holidays
    filter(!str_detect(sender, '^([a-zA-Z0-9]*-[a-zA-Z0-9]*){4}$')) %>% 
    mutate(duration = difftime(end, start, units="hours")) %>% 
    filter(duration < 24) %>% 
    mutate(org_by_other = case_when(sender == me ~ FALSE,
                                    TRUE ~ TRUE))
    
  
  
  output$calendar_meetings <- renderPrint({
    calendar_meetings
    
  })
  
  
  
  
  
  
  
}