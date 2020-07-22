# Module UI

#' @title   mod_outlookcalendar_ui and mod_outlookcalendar_server
#' @description  A shiny Module.
#'
#' @param id shiny id
#' @param input internal
#' @param output internal
#' @param session internal
#'
#' @rdname mod_outlookcalendar
#'
#' @keywords internal
#' @export 
#' @importFrom shiny NS tagList 
mod_outlookcalendar_ui <- function(id){
  ns <- NS(id)
  tagList(
    dateRangeInput(ns("dateRange"), label = 'Date range input: yyyy-mm-dd',
                   start = ymd(Sys.Date()) - ddays(5), end = ymd(Sys.Date()) + ddays(2)),
    plotOutput(ns("top5_plot")),
    br(),
    br(),
    plotOutput(ns("meeting_num_plot")),
    br(),
    br(),
    plotOutput(ns("daily_time_plot"))
  )
}

# Module Server

#' @rdname mod_outlookcalendar
#' @export
#' @keywords internal
#' @import RDCOMClient
#' @import dplyr
#' @import tidyr
#' @import purrr
#' @import tibble
#' @import stringr
#' @import openxlsx
#' @import lubridate
#' @import ggplot2
#' @importFrom scales percent
#' @importFrom forcats fct_reorder

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
  
  
  all_calendar_meetings <- df %>% 
    rowwise() %>% 
    mutate(title = meetings(record)[['subject']]) %>% 
    filter(title != "") %>% 
    mutate(sender = meetings(record)[['Organizer']],
           start = .COMDate_to_POSIX(meetings(record)[['start']]),
           end = .COMDate_to_POSIX(meetings(record)[['end']])) %>% 
    ungroup %>% 
    # remove where sender includes 4 hyphens which is in some hash thingy for loaded things like public holidays
    filter(!str_detect(sender, '^([a-zA-Z0-9]*-[a-zA-Z0-9]*){4}$')) %>% 
    mutate(duration = as.numeric(difftime(end, start, units="hours"))) %>% 
    filter(duration < 24) %>% # remove appointments that are more than 24 hours
    mutate(org_by_other = case_when(sender == me ~ FALSE,
                                    TRUE ~ TRUE)) 
  
  calendar_meetings <- reactive({
    all_calendar_meetings %>% 
      filter(start >= input$dateRange[1],
             start <= input$dateRange[2])
  })
    

  ## top 5 by organiser but not me
  top_5_by_org <- reactive({
    calendar_meetings() %>%
      filter(org_by_other) %>%
      group_by(sender) %>% 
      summarise(num_meetings = n(),
                duration = sum(duration)) %>% 
      # count(sender, name = "num_meetings") %>% 
      arrange(desc(num_meetings)) %>% 
      head(5)
  })
  
  output$top5_plot <- renderPlot({
    top_5_by_org() %>% 
      ggplot(aes(x = num_meetings, y = reorder(sender, duration))) +
      # geom_bar(stat="identity", position = position_dodge2(preserve = "single"), fill = "darkgreen", alpha = 0.4) +
      geom_bar(stat="identity", position = position_dodge(width  =1), fill = "darkgreen", alpha = 0.4) +
      geom_text(aes(label = num_meetings), position = position_dodge(width = 1), hjust = 3) +
      labs(title = "Top 5 meeting Organisers showing number of meetings") +
      theme_minimal() +
      theme(
        axis.title.x = element_blank(),
        axis.title.y = element_blank(),
        axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        panel.grid.major.x = element_blank(),
        panel.grid.minor.x = element_blank(),
        plot.title.position = "plot"
      )
  })
  
  ## number of meetings organised by me or others
  calendar_meetings_summary <- reactive({
    calendar_meetings() %>% 
      rename(organiser = org_by_other) %>% 
      mutate(organiser = case_when(organiser == FALSE ~ "Me",
                                   TRUE ~ "Other")) %>% 
      ungroup() %>% 
      group_by(organiser) %>% 
      summarise(duration = sum(duration),
                num_meetings = n()) %>% 
      mutate(num_ratio = num_meetings/sum(num_meetings),
             duration_ratio = round(duration/sum(duration), digits = 2))
      
  })
  
  output$meeting_num_plot <- renderPlot({
    calendar_meetings_summary() %>%
      ggplot(aes(y = num_meetings, fill = organiser, x = 1)) +
      geom_bar(stat="identity", position = "stack", alpha = 0.4) +
      geom_text(aes(label=paste0(num_meetings, "    (", percent(num_ratio), ")")), hjust=1.2, position = position_stack(reverse = FALSE, vjust = 0.5)) +
      ylab("num_meetings") +
      theme_void() +
      ggtitle("Number of meetings")
  })
  
  ## time spent in meetings
  daily_time <- reactive({
    date_seq <- seq.Date(from = input$dateRange[1], to = input$dateRange[2], by = "day")
    
    meetings_durations_1 <- calendar_meetings() %>% 
      mutate(start = lubridate::as_date(start)) %>% 
      select(start, duration) %>% 
      group_by(start) %>% 
      summarise(dur_meetings = sum(duration)) %>% 
      mutate(dur_working = 8 - dur_meetings) %>% 
      pivot_longer(dur_meetings:dur_working, names_to = "type", names_prefix = "dur_", values_to = "hours")
    
    date_seq %>% 
      enframe(name = NULL, value = "start") %>% 
      left_join(meetings_durations_1, by = "start") %>% 
      mutate(day = lubridate::wday(start, label = TRUE)) %>% 
      filter(!day %in% c("Sat", "Sun")) %>% 
      arrange(start) %>% 
      mutate(type = ifelse(is.na(type), "working", type),
             hours = ifelse(is.na(hours), 8, hours),
             date = paste(day, lubridate::day(start), lubridate::month(start, label = TRUE)),
             date = fct_reorder(date, start))
  })
  
  output$daily_time_plot <- renderPlot({
    daily_time() %>%
      ggplot(aes(fill=type, x=hours, y=date)) + 
      geom_bar(position="stack", stat="identity", alpha = 0.4) +
      ggtitle("Hours in meetings vs working by day") +
      xlab("") + 
      theme_minimal() +
      theme(
        axis.title.x = element_blank(),
        axis.title.y = element_blank(),
        axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        panel.grid.major.x = element_blank(),
        panel.grid.minor.x = element_blank(),
        plot.title.position = "plot"
      ) +
      geom_text(aes(label=hours), position = position_stack(reverse = FALSE, 0.9), hjust = 1.1, size = 3.5)
  })

  
  
}