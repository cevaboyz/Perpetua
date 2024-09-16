# Amazon Weekly Report Analysis

Questo repository contiene uno script R progettato per analizzare i report settimanali di Amazon, confrontando le performance tra due settimane consecutive. Lo script automatizza il processo di importazione dei dati, calcolo delle variazioni percentuali, generazione di report e applicazione di formattazione condizionale per evidenziare le variazioni positive e negative.

## Funzionalità

- **Importazione automatica dei dati**: Lo script importa automaticamente i file con specifici pattern di nomenclatura (`aggregato_wXX` e `asin_wXX`) dalla cartella selezionata.
- **Calcolo delle variazioni percentuali**: Calcola le variazioni percentuali per tutti i KPI (Key Performance Indicators) tra due settimane consecutive.
- **Generazione di report Excel**: Crea un file Excel con fogli separati per i dati aggregati e per ASIN, inclusi i dati raw e le variazioni calcolate.
- **Formattazione condizionale**: Applica una formattazione condizionale per colorare in rosso le variazioni negative e in verde quelle positive.
- **Eliminazione automatica dei file di input**: Dopo l'elaborazione, i file di input vengono eliminati per mantenere la cartella pulita.

## Prerequisiti

Assicurati di avere installato sul tuo sistema:

- R (versione 3.6 o superiore)
- Pacchetti R richiesti:
  - readr
  - dplyr
  - tidyr
  - openxlsx

## Installazione

Lo script include una funzione che verifica e installa automaticamente i pacchetti necessari. Non è richiesta alcuna azione manuale per l'installazione dei pacchetti.

## Utilizzo

1. **Preparazione dei file di input**:
   - Posiziona nella stessa cartella i file che contengono i dati da analizzare.
   - I file devono rispettare la seguente nomenclatura:
     - Dati aggregati: `aggregato_wXX.csv` o `aggregato_wXX.xlsx`
     - Dati per ASIN: `asin_wXX.csv` o `asin_wXX.xlsx`
     - `XX` rappresenta il numero della settimana (ad esempio, `aggregato_w37.csv`).

2. **Esecuzione dello script**:
   - Avvia R o RStudio.
   - Carica lo script nel tuo ambiente di lavoro.
   - Esegui lo script. Ti verrà richiesto di selezionare la cartella contenente i file da analizzare.

3. **Selezione della cartella**:
   - Verrà aperta una finestra di dialogo per selezionare la cartella con i file di input.

4. **Generazione del report**:
   - Lo script elaborerà i dati, genererà il file Excel con i risultati e applicherà la formattazione condizionale.
   - Il file di output sarà salvato nella stessa cartella dei file di input, con il nome `performance_analysis_weekXX_vs_weekYY.xlsx`, dove `XX` e `YY` sono i numeri delle settimane confrontate.

5. **Pulizia automatica**:
   - Dopo aver generato il report, lo script eliminerà automaticamente i file di input.

## Struttura del file di output

Il file Excel generato conterrà i seguenti fogli:

1. **Aggregated Variations**: Variazioni percentuali dei KPI aggregati per categoria.
2. **ASIN Variations**: Variazioni percentuali dei KPI per ciascun ASIN.
3. **Aggregated Raw Data**: Dati originali aggregati.
4. **ASIN Raw Data**: Dati originali per ASIN.

## Personalizzazione

Se desideri modificare il comportamento dello script, puoi farlo modificando direttamente il codice nelle sezioni pertinenti. Alcune possibili personalizzazioni includono:

- Cambiare le condizioni di filtro
- Aggiungere o rimuovere KPI
- Modificare la formattazione del report

I commenti all'interno dello script ti aiuteranno a individuare le parti rilevanti per le modifiche.

## Risoluzione dei problemi

- **Pacchetti mancanti**: Verifica la connessione internet e i permessi per installare pacchetti R.
- **File non trovati**: Controlla che i file di input rispettino la nomenclatura richiesta e siano nella cartella selezionata.
- **Errori di lettura dei file**: Assicurati che i file siano in formato CSV o XLSX e che la struttura delle colonne sia coerente.

## Contributi

Contributi, suggerimenti e segnalazioni di bug sono benvenuti! Sentiti libero di aprire una issue o una pull request.

## Licenza

Questo progetto è distribuito sotto la licenza MIT. Consulta il file `LICENSE` per maggiori dettagli.

## Autore

[Il tuo Nome]
