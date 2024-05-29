library(shiny)
library(rhandsontable)

ui <- fluidPage(
  titlePanel("CSV File Upload and Statistics Display"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file1", "Choose CSV File",
                accept = c(
                  "text/csv",
                  "text/comma-separated-values,text/plain",
                  ".csv")
      ),
      numericInput("nRows", "Number of rows to display:", 10, min = 1),
      uiOutput("speciesSelect"),
      downloadButton("downloadData", "Download Table as CSV")
    ),
    mainPanel(
      h3("Uploaded Data"),
      rHandsontableOutput("table"),
      h3("Statistics"),
      tableOutput("statsTable")
    )
  )
)
