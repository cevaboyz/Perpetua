###AMAZON WEEKLY REPORT ANALYSIS###


# ░▒▓██████▓▒░░▒▓██████████████▓▒░ ░▒▓██████▓▒░░▒▓████████▓▒░░▒▓██████▓▒░░▒▓███████▓▒░
# ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░      ░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░
# ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░    ░▒▓██▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░
# ░▒▓████████▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓████████▓▒░  ░▒▓██▓▒░  ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░
# ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓██▓▒░    ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░
# ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░      ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░
# ░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░░▒▓█▓▒░▒▓█▓▒░░▒▓█▓▒░▒▓████████▓▒░░▒▓██████▓▒░░▒▓█▓▒░░▒▓█▓▒░
#




# Funzione per verificare e installare le librerie necessarie
check_and_install_packages <- function(packages) {
  new_packages <- packages[!(packages %in% installed.packages()[, "Package"])]
  if (length(new_packages)) {
    cat("Installazione dei pacchetti mancanti:\n")
    print(new_packages)
    install.packages(new_packages, dependencies = TRUE)
  }
  
  for (package in packages) {
    if (!require(package, character.only = TRUE)) {
      stop(paste("Errore nel caricamento del pacchetto:", package))
    }
  }
  cat("Tutti i pacchetti necessari sono stati installati e caricati con successo.\n")
}

# Lista dei pacchetti necessari
required_packages <- c("readr", "dplyr", "tidyr", "openxlsx")

# Verifica e installa i pacchetti
check_and_install_packages(required_packages)

# Funzione per selezionare la directory interattivamente
select_directory <- function() {
  if (.Platform$OS.type == "windows") {
    choose.dir(default = "", caption = "Seleziona la cartella contenente i file da analizzare")
  } else {
    file.choose(new = TRUE)
  }
}

# Seleziona la directory interattivamente
working_directory <- select_directory()

if (is.null(working_directory)) {
  stop("Nessuna cartella selezionata. L'esecuzione dello script è stata interrotta.")
}

# Imposta la directory di lavoro
setwd(working_directory)
print(paste("Directory di lavoro impostata:", working_directory))

# Funzione per importare i file
import_files <- function(pattern) {
  files <- list.files(pattern = pattern)
  if (length(files) == 0) {
    stop(paste("Nessun file trovato con il pattern:", pattern))
  }
  data_list <- lapply(files, function(file) {
    week <- as.numeric(sub(".*w(\\d+).*", "\\1", file))
    if (grepl("\\.csv$", file)) {
      data <- read_csv(file)
    } else if (grepl("\\.xlsx$", file)) {
      data <- read.xlsx(file)
    }
    data$Week <- week
    return(data)
  })
  return(list(data = do.call(rbind, data_list), files = files))
}

# Importa i dati aggregati
aggregated_result <- import_files("aggregato_w\\d+")
aggregated_data <- aggregated_result$data
aggregated_files <- aggregated_result$files

# Importa i dati per ASIN
asin_result <- import_files("asin_w\\d+")
asin_data <- asin_result$data
asin_files <- asin_result$files

# Funzione per filtrare i dati con Spend > 0
filter_non_zero_spend <- function(data) {
  data %>% filter(Spend > 0)
}

# Applica il filtro ai dati
aggregated_data <- filter_non_zero_spend(aggregated_data)
asin_data <- filter_non_zero_spend(asin_data)

# Estrai i numeri delle settimane
weeknumbers <- sort(unique(c(aggregated_data$Week, asin_data$Week)))

# Funzione per calcolare le variazioni percentuali
calculate_variations <- function(data, group_vars) {
  data %>%
    arrange(!!!syms(group_vars), Week) %>%
    group_by(!!!syms(group_vars)) %>%
    mutate(across(where(is.numeric) &
                    !Week, ~ (. - lag(.)) / lag(.) * 100, .names = "{.col}_Var%")) %>%
    ungroup() %>%
    select(!!!syms(group_vars), Week, ends_with("Var%"))
}

# Calcola le variazioni per i dati aggregati
aggregated_variations <- calculate_variations(aggregated_data, "CATEGORY")

# Calcola le variazioni per i dati ASIN
asin_variations <- calculate_variations(asin_data, c("Variant ASIN", "ASIN title", "CATEGORY"))

# Funzione per formattare i dati come percentuali
format_as_percentage <- function(data) {
  data %>%
    mutate(across(ends_with("Var%"), ~ round(., 2))) %>%
    mutate(across(ends_with("Var%"), ~ paste0(., "%")))
}

# Applica la formattazione percentuale
aggregated_variations_formatted <- format_as_percentage(aggregated_variations)
asin_variations_formatted <- format_as_percentage(asin_variations)

# Crea un nuovo workbook
wb <- createWorkbook()

# Funzione per aggiungere un foglio al workbook con opzioni di filtro e ordinamento
add_sheet <- function(wb,
                      data,
                      sheet_name,
                      filter_na = FALSE,
                      sort_columns = NULL) {
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, data)
  
  # Se il foglio contiene dati formattati come percentuali, applica lo stile percentuale
  if (any(grepl("Var%", names(data)))) {
    percent_cols <- which(grepl("Var%", names(data)))
    for (col in percent_cols) {
      addStyle(
        wb,
        sheet_name,
        style = createStyle(numFmt = "0.00%"),
        rows = 2:(nrow(data) + 1),
        cols = col,
        gridExpand = TRUE
      )
    }
  }
  
  # Applica il filtro per nascondere le righe NA se richiesto
  if (filter_na) {
    addFilter(wb,
              sheet_name,
              row = 1,
              cols = 1:ncol(data))
  }
  
  # Applica l'ordinamento se sono specificate le colonne
  if (!is.null(sort_columns)) {
    orderData <- data[do.call(order, data[sort_columns]), ]
    writeData(wb,
              sheet_name,
              orderData,
              startRow = 2,
              colNames = FALSE)
  }
}

# Prepara i dati ASIN per l'ordinamento
asin_variations_for_sorting <- asin_variations_formatted %>%
  mutate(Ordered_Units = as.numeric(gsub("%", "", `Ordered Units_Var%`)))

# Aggiungi i fogli al workbook
add_sheet(wb, aggregated_variations_formatted, "Aggregated Variations")
add_sheet(
  wb,
  asin_variations_for_sorting,
  "ASIN Variations",
  filter_na = TRUE,
  sort_columns = c("CATEGORY", "Ordered_Units")
)
add_sheet(wb, aggregated_data, "Aggregated Raw Data")
add_sheet(wb, asin_data, "ASIN Raw Data")


# Crea il nome del file dinamicamente
filename <- sprintf("performance_analysis_week%d_vs_week%d.xlsx",
                    max(weeknumbers),
                    min(weeknumbers))

# Salva il workbook
saveWorkbook(wb, file.path(working_directory, filename), overwrite = TRUE)

# Funzione per eliminare i file di partenza
delete_input_files <- function(files) {
  for (file in files) {
    if (file.exists(file)) {
      file.remove(file)
      print(paste("File eliminato:", file))
    } else {
      print(paste("File non trovato:", file))
    }
  }
}

# Elimina i file di partenza
delete_input_files(c(aggregated_files, asin_files))

print(paste(
  "Analisi completata. I risultati sono stati salvati in",
  file.path(working_directory, filename)
))
print("I file di partenza sono stati eliminati.")
print(
  paste(
    "Righe rimosse (Spend = 0):",
    nrow(aggregated_result$data) - nrow(aggregated_data),
    "(dati aggregati),",
    nrow(asin_result$data) - nrow(asin_data),
    "(dati ASIN)"
  )
)
