library(shiny)
library(rhandsontable)
library(dplyr)
library(tidyr)

server <- function(input, output, session) {
  # Reactive expression to read the uploaded file
  data <- reactive({
    req(input$file1)
    inFile <- input$file1
    read.csv(inFile$datapath)
  })
  
  # Update species selection based on uploaded data
  output$speciesSelect <- renderUI({
    req(data())
    selectInput("species", "Filter by Species:", 
                choices = c("All", unique(data()$Species)), 
                selected = "All")
  })
  
  # Filtered data based on species selection
  filteredData <- reactive({
    req(data())
    if (input$species == "All") {
      data()
    } else {
      data() %>% filter(Species == input$species)
    }
  })
  
  # Render the editable data table
  output$table <- renderRHandsontable({
    req(filteredData())
    rhandsontable(head(filteredData(), input$nRows), rowHeaders = NULL)
  })
  
  # Reactive expression to get the edited data
  editedData <- reactive({
    if (!is.null(input$table)) {
      hot_to_r(input$table)
    } else {
      head(filteredData(), input$nRows)
    }
  })
  
  # Calculate basic statistics for each species
  stats <- reactive({
    df <- editedData()
    df %>%
      group_by(Species) %>%
      summarise_if(is.numeric, list(
        mean = ~mean(.),
        min = ~min(.),
        max = ~max(.),
        count = ~n()
      )) %>%
      pivot_longer(cols = -Species, names_to = c("Variable", "Statistic"), names_sep = "_")
  })
  
  # Render the statistics table
  output$statsTable <- renderTable({
    stats()
  })
  
  # Downloadable csv of the edited dataset
  output$downloadData <- downloadHandler(
    filename = function() {
      paste("table-", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      write.csv(editedData(), file, row.names = FALSE)
    }
  )
}
