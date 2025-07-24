library(shiny)
library(readxl)
library(tidyverse)
library(plotly)
library(DT)
library(webshot2)
library(htmlwidgets)

# webshot::install_phantomjs()  # Nur beim ersten Mal nötig

# Excel-Datei aufteilen in Kompetenzbereiche
split_excel_by_section <- function(file_path, sheet = 1) {
  df_raw <- read_excel(file_path, sheet = sheet, col_names = FALSE)
  col_names <- as.character(df_raw[1, ])
  df_data <- df_raw[-1, ]

  result <- list()
  current_section <- NULL
  current_block <- list()

  for (i in 1:nrow(df_data)) {
    row <- df_data[i, ]
    if (all(is.na(row))) {
      if (!is.null(current_section) && length(current_block) > 0) {
        block_df <- bind_rows(current_block)
        colnames(block_df) <- col_names
        result[[current_section]] <- block_df
      }
      current_section <- NULL
      current_block <- list()
    } else {
      if (is.null(current_section)) {
        current_section <- as.character(row[[1]])
      } else {
        current_block[[length(current_block) + 1]] <- row
      }
    }
  }

  if (!is.null(current_section) && length(current_block) > 0) {
    block_df <- bind_rows(current_block)
    colnames(block_df) <- col_names
    result[[current_section]] <- block_df
  }

  return(result)
}

# UI ----
ui <- fluidPage(
  titlePanel("Teamkompetenz-Analyse nach Bereichen"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Excel-Datei hochladen (.xlsx)", accept = ".xlsx"),
      uiOutput("bereich_ui"),
      uiOutput("person_ui"),
      checkboxInput("show_gap", "Kompetenzlücken anzeigen (Ziel-Level)", value = FALSE),
      sliderInput("zielwert", "Ziel-Level", min = 0, max = 3, step = 0.1, value = 2)
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Radar Chart", plotlyOutput("radar_plot"), downloadButton("download_radar", "Radar Chart als PNG")),
        tabPanel("Heatmap", plotOutput("heatmap_plot"), downloadButton("download_heatmap", "Heatmap als PNG")),
        tabPanel("Tabelle", dataTableOutput("skill_table")),
        tabPanel("Skillabdeckung", plotOutput("coverage_plot"), downloadButton("download_coverage", "Skillabdeckung als PNG"))
      )
    )
  )
)

# Server ----
server <- function(input, output, session) {
  daten <- reactiveVal(NULL)

  observeEvent(input$file, {
    req(input$file)
    daten(split_excel_by_section(input$file$datapath))
  })

  output$bereich_ui <- renderUI({
    req(daten())
    selectInput("bereich", "Kompetenzbereich auswählen", choices = names(daten()))
  })

  output$person_ui <- renderUI({
    req(daten(), input$bereich)
    df <- daten()[[input$bereich]]
    personen <- colnames(df)[-1]
    selectizeInput("selected_people", "Teammitglieder auswählen",
                   choices = personen, multiple = TRUE, selected = personen,
                   options = list(plugins = list('remove_button')))
  })

  selected_df <- reactive({
    req(daten(), input$bereich)
    df <- daten()[[input$bereich]]
    df %>%
      mutate(across(-1, ~ as.numeric(.))) %>%
      mutate(across(-1, ~ replace_na(., 0)))
  })

  radar_chart_plot <- reactive({
    req(selected_df(), input$selected_people)
    ziel_level <- input$zielwert
    df <- selected_df() %>%
      pivot_longer(-Wissen, names_to = "Person", values_to = "Level") %>%
      filter(Person %in% input$selected_people)
    kategorien <- unique(df$Wissen)
    df <- df %>% mutate(Wissen = factor(Wissen, levels = kategorien)) %>% arrange(Person, Wissen)
    p <- plot_ly(type = "scatterpolar", fill = "toself")
    for (person in input$selected_people) {
      person_data <- df %>% filter(Person == person)
      p <- add_trace(p, r = person_data$Level, theta = person_data$Wissen,
                     name = person, mode = "markers+lines", fill = "toself")
    }
    if (input$show_gap) {
      ziel <- df %>% distinct(Wissen) %>% mutate(Level = ziel_level) %>% arrange(Wissen)
      ziel <- bind_rows(ziel, ziel[1, ])
      p <- add_trace(p,
                     r = ziel$Level,
                     theta = ziel$Wissen,
                     name = paste0("Ziel-Level (", ziel_level, ")"),
                     mode = "lines",
                     fill = "toself",
                     fillcolor = "rgba(0, 128, 0, 0.1)",  # sehr dezente grüne Fläche
                     line = list(
                       dash = "dash",
                       width = 2,
                       color = "rgba(0, 128, 0, 0.4)"
                     ),
                     showlegend = TRUE
      )
      
    }
    p %>% layout(polar = list(radialaxis = list(visible = TRUE, range = c(0, 3))), showlegend = TRUE)
  })

  output$radar_plot <- renderPlotly({ radar_chart_plot() })

  output$download_radar <- downloadHandler(
    filename = function() { paste0("radar_chart_", Sys.Date(), ".png") },
    content = function(file) {
      html <- tempfile(fileext = ".html")
      saveWidget(radar_chart_plot(), html, selfcontained = TRUE)
      webshot2::webshot(html, file = file, vwidth = 800, vheight = 800)
    }
  )

  heatmap_plot_obj <- reactive({
    req(selected_df(), input$selected_people)
    df <- selected_df() %>%
      pivot_longer(-Wissen, names_to = "Person", values_to = "Level") %>%
      filter(Person %in% input$selected_people)
    ggplot(df, aes(x = Person, y = Wissen, fill = Level)) +
      geom_tile(color = "white") +
      scale_fill_gradient(low = "#ffffcc", high = "#ff4444", name = "Level") +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
  })

  output$heatmap_plot <- renderPlot({ heatmap_plot_obj() })

  output$download_heatmap <- downloadHandler(
    filename = function() { paste0("heatmap_", Sys.Date(), ".png") },
    content = function(file) {
      ggsave(file, plot = heatmap_plot_obj(), width = 10, height = 8, dpi = 300)
    }
  )

  coverage_plot_obj <- reactive({
    req(selected_df(), input$selected_people)
    ziel_level <- input$zielwert
    min_personen <- 2
    df <- selected_df() %>%
      select(Wissen, all_of(input$selected_people)) %>%
      pivot_longer(-Wissen, names_to = "Person", values_to = "Level") %>%
      mutate(Level = as.numeric(Level)) %>%
      group_by(Wissen) %>%
      summarise(Anzahl = sum(Level >= ziel_level, na.rm = TRUE), .groups = "drop") %>%
      mutate(Ziel = Anzahl >= min_personen)
    ggplot(df, aes(x = Wissen, y = Anzahl, fill = Ziel)) +
      geom_col() +
      geom_hline(yintercept = min_personen, linetype = "dashed", color = "red") +
      scale_fill_manual(values = c("TRUE" = "#66bb6a", "FALSE" = "#ef5350"),
                        labels = c("TRUE" = "Ziel erreicht", "FALSE" = "Ziel nicht erreicht"),
                        name = "") +
      labs(y = paste0("Anzahl ≥ ", ziel_level), x = "Kompetenz") +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1), legend.position = "top")
  })

  output$coverage_plot <- renderPlot({ coverage_plot_obj() })

  output$download_coverage <- downloadHandler(
    filename = function() { paste0("skillabdeckung_", Sys.Date(), ".png") },
    content = function(file) {
      ggsave(file, plot = coverage_plot_obj(), width = 10, height = 6, dpi = 300)
    }
  )

  output$skill_table <- renderDataTable({
    req(selected_df())
    datatable(selected_df(), options = list(pageLength = 10))
  })
}

shinyApp(ui, server)

