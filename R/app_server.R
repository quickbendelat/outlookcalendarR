#' The application server-side
#' 
#' @param input,output,session Internal parameters for {shiny}. 
#'     DO NOT REMOVE.
#' @import shiny
#' @import dplyr
#' @import tidyr
#' @import RDCOMClient
#' @import purrr
#' @import tibble
#' @import stringr
#' @import openxlsx
#' @noRd
app_server <- function( input, output, session ) {
  
  
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
  

  calendar_meetings <- reactive({
    
    df %>% 
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
    
  })


  
  
  
  output$debug <- renderPrint({
    debug = calendar_meetings()
    debug
  })
  
}
