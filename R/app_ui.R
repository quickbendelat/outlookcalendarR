#' The application User-Interface
#' 
#' @param request Internal parameter for `{shiny}`. 
#'     DO NOT REMOVE.
#' @import shiny
#' @import shinydashboard
#' @noRd
app_ui <- function(request) {
  tagList(
    # Leave this function for adding external resources
    # golem_add_external_resources(),
    # List the first level UI elements here 
    dashboardPage(
      skin = "black",
      dashboardHeader(title = "Outlook Calendar Dash"),
      dashboardSidebar(
        sidebarMenu(
          menuItem("Dashboard", tabName = "dashboard", icon = icon("dashboard"))
        )
      ),
      dashboardBody(
        tabItems(
          # First tab content
          tabItem(tabName = "dashboard",
                  fluidRow(
                    column(width = 12,
                           # dateRangeInput('dateRange',
                           #                label = 'Date range input: yyyy-mm-dd',
                           #                start = Sys.Date() - 2, end = Sys.Date() + 2
                           # ),
                           mod_outlookcalendar_ui("outlookcalendar_module_ui")
                    )
                  )
          )
        )
      )
    )
  )
}

#' #' Add external Resources to the Application
#' #' 
#' #' This function is internally used to add external 
#' #' resources inside the Shiny application. 
#' #' 
#' #' @import shiny
#' #' @importFrom golem add_resource_path activate_js favicon bundle_resources
#' #' @noRd
#' golem_add_external_resources <- function(){
#'   
#'   add_resource_path(
#'     'www', app_sys('app/www')
#'   )
#'  
#'   tags$head(
#'     favicon(),
#'     bundle_resources(
#'       path = app_sys('app/www'),
#'       app_title = 'outlookcalendarR'
#'     )
#'     # Add here other external resources
#'     # for example, you can add shinyalert::useShinyalert() 
#'   )
#' }

