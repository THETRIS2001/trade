// Configurazioni globali
const allowedTypes = ["Commercio", "Pagamento degli interessi", "Premio", "Redditi"];
const operationTypes = [
    "Buy trade",
    "Savings plan execution",
    "Sell trade",
    "Warrant Exercise"
];

// Variabili globali
let additionalValues = []; // Ora ogni elemento sarà un oggetto { value: number, rowId: string, description: string }
let outputData = [];
let finalData = {
    dfInteressiGiacenza: [],
    dfWarrant: [],
    dfWarrantExercise: [],
    dfSaveback: [],
    dfCommercioSell: [],
    dfCommercioBuy: []
};
let loadingOverlay;
let additionalValuesSection;
let resultsSection;

// Variabili globali per i grafici
let tradeChart = null;
let chartPeriod = 'monthly';
let chartYear = new Date().getFullYear();
let availableYears = [];

// Funzione globale per la gestione dei valori aggiuntivi
let setupAdditionalValuesHandling;

// Elementi DOM
document.addEventListener('DOMContentLoaded', () => {
    // Riferimenti agli elementi DOM
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('file-input');
    const selectFileBtn = document.getElementById('select-file-btn');
    const removeFileBtn = document.getElementById('remove-file');
    const processBtn = document.getElementById('process-btn');
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const fileInfo = document.getElementById('file-info');
    const additionalValuesSection = document.querySelector('.additional-values-section');
    const resultsSection = document.querySelector('.results-section');
    
    // Output e tabelle
    const outputTable = document.getElementById('output-table');
    const finalTable = document.getElementById('final-table');
    const outputSearchBtn = document.getElementById('output-search-btn');
    const finalSearchBtn = document.getElementById('final-search-btn');
    const outputSearch = document.getElementById('output-search');
    const finalSearch = document.getElementById('final-search');
    const downloadOutputBtn = document.getElementById('download-output-csv');
    const downloadFinalBtn = document.getElementById('download-final-csv');
    
    // Elementi per aggiungere valori
    const additionalValuesInput = document.getElementById('additional-values');
    const addValuesBtn = document.getElementById('add-values-btn');
    const associateRowSelect = document.getElementById('associate-row');
    const addedValuesContainer = document.getElementById('added-values-container');
    
    // Elementi riepilogo
    const totalTrade = document.getElementById('total-trade');
    const totalInterest = document.getElementById('total-interest');
    const totalSaveback = document.getElementById('total-saveback');
    const totalComprehensive = document.getElementById('total-comprehensive');
    
    // Elementi per guadagni e perdite
    const totalProfits = document.getElementById('total-profits');
    const totalLosses = document.getElementById('total-losses');
    
    // Loading overlay
    const loadingOverlay = document.querySelector('.loading-overlay');
    
    // Event listeners per il drag and drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    // Funzioni di supporto per il drag and drop
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    function highlight() {
        dropArea.classList.add('highlight');
    }
    
    function unhighlight() {
        dropArea.classList.remove('highlight');
    }
    
    function handleDrop(e) {
        console.log('File dropped:', e);
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length) {
            handleFiles(files);
        }
    }
    
    // Drop file handler
    dropArea.addEventListener('drop', handleDrop, false);
    
    // Seleziona file button handler
    selectFileBtn.addEventListener('click', () => {
        console.log('Seleziona file button clicked');
        fileInput.click();
    });
    
    // Input file change handler
    fileInput.addEventListener('change', (e) => {
        console.log('File input change:', e.target.files);
        handleFiles(e.target.files);
    });
    
    // Rimuovi file button handler
    removeFileBtn.addEventListener('click', () => {
        // Ripulisci e resetta
        fileName.textContent = '';
        fileSize.textContent = '';
        fileInfo.style.display = 'none';
        window.selectedFile = null;
        dropArea.style.borderColor = '';
        
        // Nascondi le sezioni
        additionalValuesSection.style.display = 'none';
        resultsSection.style.display = 'none';
    });
    
    // Event listener per i tab
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.getAttribute('data-tab');
            
            // Rimuovi la classe active da tutti i bottoni e pannelli
            tabBtns.forEach(b => b.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));
            
            // Aggiungi la classe active al bottone corrente e al pannello corrispondente
            btn.classList.add('active');
            document.getElementById(target).classList.add('active');
        });
    });
    
    // Event listeners per la ricerca
    outputSearchBtn.addEventListener('click', () => {
        const searchTerm = outputSearch.value.toLowerCase();
        searchTable(outputTable, searchTerm);
    });
    
    finalSearchBtn.addEventListener('click', () => {
        const searchTerm = finalSearch.value.toLowerCase();
        searchTable(finalTable, searchTerm);
    });
    
    // Modifica per ricerca in tempo reale
    outputSearch.addEventListener('input', (e) => {
        const searchTerm = outputSearch.value.toLowerCase();
        searchTable(outputTable, searchTerm);
    });
    
    finalSearch.addEventListener('input', (e) => {
        const searchTerm = finalSearch.value.toLowerCase();
        searchTable(finalTable, searchTerm);
    });
    
    // Manteniamo anche gli event listener per il tasto Enter
    outputSearch.addEventListener('keyup', (e) => {
        if (e.key === 'Enter') {
            const searchTerm = outputSearch.value.toLowerCase();
            searchTable(outputTable, searchTerm);
        }
    });
    
    finalSearch.addEventListener('keyup', (e) => {
        if (e.key === 'Enter') {
            const searchTerm = finalSearch.value.toLowerCase();
            searchTable(finalTable, searchTerm);
        }
    });
    
    // Event listeners per i download CSV
    downloadOutputBtn.addEventListener('click', () => {
        // Ottieni la data dal titolo del riepilogo
        const summaryTitle = document.querySelector('.summary-container h2');
        let dateText = '';
        
        if (summaryTitle) {
            // Estrai la data dal titolo che ha il formato "Riepilogo Ricavi al [DATA]"
            const titleText = summaryTitle.textContent;
            const dateMatch = titleText.match(/al\s+(.+)$/);
            if (dateMatch && dateMatch[1]) {
                dateText = dateMatch[1].trim();
            }
        }
        
        // Crea il nome del file con la data corrente
        const filename = dateText ? `Dettaglio Operazioni al ${dateText}.csv` : 'Dettaglio Operazioni.csv';
        downloadCSV(outputTable, filename);
    });
    
    downloadFinalBtn.addEventListener('click', () => {
        // Ottieni la data dal titolo del riepilogo
        const summaryTitle = document.querySelector('.summary-container h2');
        let dateText = '';
        
        if (summaryTitle) {
            // Estrai la data dal titolo che ha il formato "Riepilogo Ricavi al [DATA]"
            const titleText = summaryTitle.textContent;
            const dateMatch = titleText.match(/al\s+(.+)$/);
            if (dateMatch && dateMatch[1]) {
                dateText = dateMatch[1].trim();
            }
        }
        
        // Crea il nome del file con la data corrente
        const filename = dateText ? `Riepilogo Finale al ${dateText}.csv` : 'Riepilogo Finale.csv';
        downloadCSV(finalTable, filename);
    });
    
    // Event listeners per ordinare le colonne delle tabelle
    const outputHeaders = outputTable.querySelectorAll('th[data-sort]');
    const finalHeaders = finalTable.querySelectorAll('th[data-sort]');
    
    // Applica gli event handlers alle intestazioni di entrambe le tabelle
    handleHeaderClick(outputHeaders, outputTable);
    handleHeaderClick(finalHeaders, finalTable);
    
    // Implemento la funzione a livello globale
    setupAdditionalValuesHandling = function() {
        // Riferimenti agli elementi DOM
        const associateRowSelect = document.getElementById('associate-row');
        const rowSearchInput = document.getElementById('row-search');
        
        // Array per memorizzare tutte le opzioni originali
        let allOptions = [];
        
        // Funzione per aggiornare le opzioni del selettore di righe
        function updateRowSelectOptions() {
            console.log("Aggiornamento opzioni del selettore di righe");
            
            if (!associateRowSelect) {
                console.error("Elemento select 'associate-row' non trovato!");
                return;
            }
            
            // Salva gli elementi selezionati prima di cancellare
            const selectedValue = associateRowSelect.value;
            
            // Pulisci completamente il selettore
            associateRowSelect.innerHTML = '';
            allOptions = []; // Resetta anche l'array delle opzioni
            
            // Non aggiungiamo più l'opzione separatore per evitare la prima riga
            
            // Ottieni le righe dal riepilogo finale
            let rows;
            try {
                rows = createFinalResult();
                console.log("Righe trovate:", rows.length);
            } catch (e) {
                console.error("Errore nel recupero delle righe:", e);
                rows = [];
            }
            
            // Tiene traccia delle descrizioni già aggiunte per evitare duplicati
            const addedDescriptions = new Set();
            
            if (rows && rows.length > 0) {
                rows.forEach((row, index) => {
                    // Filtra per includere anche Interessi Giacenza e Saveback, oltre a Buy trade e Sell trade Parziale
                    if (row.TIPO_OPERAZIONE !== 'Buy trade' && 
                        row.TIPO_OPERAZIONE !== 'Sell trade Parziale' && 
                        row.TIPO_OPERAZIONE !== 'Interessi Giacenza' && 
                        row.TIPO_OPERAZIONE !== 'Saveback') {
                        return; // salta questa iterazione se non è uno dei tipi richiesti
                    }
                    
                    // Crea la descrizione della riga
                    let rowDescription;
                    
                    // Per Interessi Giacenza e Saveback, mostra solo il tipo senza aggiungere "- Senza descrizione"
                    if ((row.TIPO_OPERAZIONE === 'Interessi Giacenza' || row.TIPO_OPERAZIONE === 'Saveback') && !row.Descrizione) {
                        rowDescription = row.TIPO_OPERAZIONE;
                    } else {
                        rowDescription = `${row.TIPO_OPERAZIONE} - ${row.Descrizione || 'Senza descrizione'}`;
                    }
                    
                    // Controlla se questa descrizione è già stata aggiunta
                    if (addedDescriptions.has(rowDescription)) {
                        console.log(`Saltata riga duplicata: ${rowDescription}`);
                        return; // salta questa iterazione
                    }
                    
                    // Aggiungi la descrizione al set per controllare i duplicati
                    addedDescriptions.add(rowDescription);
                    
                    // Non tronchiamo più la descrizione 
                    const displayDescription = rowDescription;
                    
                    const option = document.createElement('option');
                    option.value = `row-${index}`;
                    option.textContent = displayDescription;
                    
                    // Aggiungi attributi per i dettagli completi della riga
                    option.setAttribute('data-tipo', row.TIPO_OPERAZIONE || '');
                    option.setAttribute('data-desc', row.Descrizione || '');
                    
                    // Memorizza l'opzione nell'array per il filtraggio successivo
                    allOptions.push({
                        element: option,
                        text: displayDescription.toLowerCase(),
                        tipo: (row.TIPO_OPERAZIONE || '').toLowerCase(),
                        desc: (row.Descrizione || '').toLowerCase()
                    });
                    
                    associateRowSelect.appendChild(option);
                });
                
                console.log(`Aggiunte ${addedDescriptions.size} opzioni uniche al selettore`);
                
                // Ripristina la selezione precedente se possibile
                if (selectedValue) {
                    const previousOption = Array.from(associateRowSelect.options).find(
                        opt => opt.value === selectedValue
                    );
                    if (previousOption) {
                        associateRowSelect.value = selectedValue;
                    }
                }
                
                // Resetta il campo di ricerca
                if (rowSearchInput) {
                    rowSearchInput.value = '';
                }
            }
        }
        
        // Funzione per filtrare le opzioni in base al testo di ricerca
        function filterOptions(searchText) {
            // Se non ci sono opzioni o elementi, esci
            if (!allOptions.length || !associateRowSelect) return;
            
            // Pulisci completamente il selettore
            associateRowSelect.innerHTML = '';
            
            // Converti il testo di ricerca in minuscolo per una ricerca case-insensitive
            const searchLower = searchText.toLowerCase();
            
            // Se la ricerca è vuota, mostra tutte le opzioni
            if (!searchLower) {
                allOptions.forEach(opt => {
                    associateRowSelect.appendChild(opt.element);
                });
                // Assicura che non ci sia nessun valore selezionato di default
                associateRowSelect.selectedIndex = -1;
                return;
            }
            
            // Filtra le opzioni che corrispondono alla ricerca
            const filteredOptions = allOptions.filter(opt => 
                opt.text.includes(searchLower) || 
                opt.tipo.includes(searchLower) || 
                opt.desc.includes(searchLower)
            );
            
            // Aggiungi le opzioni filtrate al selettore
            filteredOptions.forEach(opt => {
                associateRowSelect.appendChild(opt.element);
            });
            
            // Assicura che non ci sia nessun valore selezionato di default
            associateRowSelect.selectedIndex = -1;
            
            // Se non ci sono risultati, mostra un messaggio
            if (filteredOptions.length === 0) {
                const noResultOption = document.createElement('option');
                noResultOption.disabled = true;
                noResultOption.textContent = 'Nessun risultato trovato';
                associateRowSelect.appendChild(noResultOption);
            }
        }
        
        // Aggiungi l'event listener per la ricerca in tempo reale
        if (rowSearchInput) {
            // Mostra la tendina quando si clicca sulla barra di ricerca
            rowSearchInput.addEventListener('click', () => {
                // Mostra tutti i risultati prima di iniziare a digitare
                filterOptions('');
                associateRowSelect.style.display = 'block';
                associateRowSelect.size = 8;
                // Assicura che non ci sia nessun valore selezionato di default
                associateRowSelect.selectedIndex = -1;
            });
            
            rowSearchInput.addEventListener('input', (e) => {
                filterOptions(e.target.value);
                // Mostra il dropdown solo quando l'utente digita
                if (associateRowSelect) {
                    associateRowSelect.style.display = 'block';
                    associateRowSelect.size = 8; // Mostra più righe di opzioni
                    // Assicura che non ci sia nessun valore selezionato di default
                    associateRowSelect.selectedIndex = -1;
                }
            });
            
            // Aggiungi eventi per chiudere la tendina quando si perde il focus
            rowSearchInput.addEventListener('blur', (e) => {
                // Controlliamo se l'elemento che riceve il focus è il select
                // per evitare di chiuderlo immediatamente quando si clicca su di esso
                setTimeout(() => {
                    if (document.activeElement !== associateRowSelect) {
                        associateRowSelect.style.display = 'none';
                    }
                }, 200);
            });
            
            // Chiudi la tendina quando viene selezionata un'opzione
            associateRowSelect.addEventListener('change', () => {
                // Imposta il valore nell'input di ricerca
                if (associateRowSelect.selectedIndex >= 0) {
                    const selectedOption = associateRowSelect.options[associateRowSelect.selectedIndex];
                    rowSearchInput.value = selectedOption.textContent;
                }
                associateRowSelect.style.display = 'none';
            });
            
            // Chiudi la tendina quando si perde il focus dal select
            associateRowSelect.addEventListener('blur', () => {
                associateRowSelect.style.display = 'none';
            });
        }
        
        // Ritorna la funzione per aggiornare le opzioni del selettore
        return updateRowSelectOptions;
    };
    
    // Gestione del click sul pulsante Aggiungi
    addValuesBtn.addEventListener('click', () => {
        const valueStr = additionalValuesInput.value.trim();
        if (!valueStr) return;
        
        try {
            // Converti il valore in numero (supportando sia punto che virgola come separatore decimale)
            const cleanValue = valueStr.replace(',', '.');
            const newValue = parseFloat(cleanValue);
            
            if (isNaN(newValue)) {
                throw new Error('Valore non valido');
            }
            
            // Ottieni la riga selezionata
            const selectedRowId = associateRowSelect.value;
            
            // Verifica che sia stata selezionata una riga (non è più possibile associare a "global")
            if (!selectedRowId || selectedRowId === 'global') {
                alert('Seleziona una riga dal riepilogo a cui associare il valore.');
                return;
            }
            
            const selectedOption = associateRowSelect.options[associateRowSelect.selectedIndex];
            const tipo = selectedOption.getAttribute('data-tipo');
            const desc = selectedOption.getAttribute('data-desc');
            const rowDescription = `${tipo} - ${desc}`;
            
            console.log(`Associazione valore ${newValue} a: ${selectedRowId} (${rowDescription})`);
            
            // Aggiungi il nuovo valore all'array esistente
            additionalValues.push({
                value: newValue,
                rowId: selectedRowId,
                description: rowDescription
            });
            
            // Aggiorna la visualizzazione dei valori
            updateAddedValuesDisplay();
            
            // Pulisci l'input
            additionalValuesInput.value = '';
            
            // Pulisci anche il campo di ricerca
            const rowSearchInput = document.getElementById('row-search');
            if (rowSearchInput) {
                rowSearchInput.value = '';
            }
            
            // Se ci sono dati elaborati, ricalcola i totali
            if (Object.keys(finalData).length > 0) {
                // Ricrea il risultato finale con i nuovi valori
                populateFinalTable();
                calculateAndDisplayTotals();
                
                // Aggiungi la riga al dettaglio operazioni
                addValueToOutputTable(newValue, rowDescription);
                
                // Aggiorna la data del riepilogo
                updateSummaryDate();
                
                // Aggiorna il selettore delle righe immediatamente
                if (typeof window.updateRowSelector === 'function') {
                    window.updateRowSelector();
                }
            }
        } catch (e) {
            alert('Errore: Inserisci un valore numerico valido.');
        }
    });
    
    // Aggiungi anche l'evento sul press del tasto Invio
    additionalValuesInput.addEventListener('keyup', (e) => {
        if (e.key === 'Enter') {
            addValuesBtn.click();
        }
    });
    
    // Limitiamo l'input a soli numeri e punto/virgola
    additionalValuesInput.addEventListener('input', (e) => {
        // Rimuovi tutti i caratteri che non sono numeri, punto o virgola
        const validValue = e.target.value.replace(/[^0-9.,]/g, '');
        
        // Se il valore è cambiato, aggiornalo
        if (validValue !== e.target.value) {
            e.target.value = validValue;
        }
    });
    
    // Funzione per aggiornare la visualizzazione dei valori aggiunti
    window.updateAddedValuesDisplay = function() {
        addedValuesContainer.innerHTML = '';
        
        additionalValues.forEach((item, index) => {
            const tag = document.createElement('div');
            tag.className = `value-tag ${item.value >= 0 ? 'positive' : 'negative'}`;
            
            tag.innerHTML = `
                <div class="value-tag-header">
                    ${normalizeNumericFormat(item.value, true)} € 
                    <span class="remove-value" data-index="${index}">✕</span>
                </div>
                <div class="value-tag-association" title="${item.description}">
                    ${item.description}
                </div>
            `;
            
            addedValuesContainer.appendChild(tag);
        });
        
        // Aggiungi event listener per rimuovere i valori
        document.querySelectorAll('.remove-value').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const index = parseInt(e.target.getAttribute('data-index'));
                const removedValue = additionalValues[index];
                additionalValues.splice(index, 1);
                updateAddedValuesDisplay();
                
                // Se ci sono dati elaborati, ricalcola i totali
                if (Object.keys(finalData).length > 0) {
                    // Rimuovi la riga corrispondente dal dettaglio operazioni (se presente)
                    removeAddedValueRow(removedValue);
                    
                    // Ricrea il risultato finale con i nuovi valori
                    populateFinalTable();
                    calculateAndDisplayTotals();
                    
                    // Aggiorna la data del riepilogo
                    updateSummaryDate();
                    
                    // Aggiorna il selettore delle righe immediatamente
                    if (typeof window.updateRowSelector === 'function') {
                        window.updateRowSelector();
                    }
                }
            });
        });
    };
    
    // Funzione per rimuovere una riga di valore aggiunto dalla tabella dettaglio operazioni
    function removeAddedValueRow(removedValue) {
        // Cerca tutte le righe con classe added-value-row
        const rows = document.querySelectorAll('#output-table tbody tr.added-value-row');
        
        // Prepara la descrizione da cercare
        let description = removedValue.description;
        // Estrai la parte della descrizione dopo il primo trattino (se presente)
        const descriptionParts = description.split(' - ');
        if (descriptionParts.length > 1) {
            description = descriptionParts.slice(1).join(' - '); // Prendi tutto dopo il primo "Tipo - "
        }
        
        // Prepara il valore formattato da cercare
        const valueText = normalizeNumericFormat(removedValue.value, true) + " €";
        
        // Controlla ogni riga per trovare quella corrispondente
        rows.forEach(row => {
            const descCell = row.querySelector('td:nth-child(4)');
            const valueCell = row.querySelector('td:nth-child(7)');  // Ora il prezzo di vendita è nella 7a colonna
            
            // Verifica sia la descrizione che il valore (il prezzo di vendita)
            if (descCell && valueCell && 
                descCell.textContent.trim() === description && 
                valueCell.textContent.trim() === valueText) {
                // Trova corrispondenza, rimuovi la riga
                row.remove();
            }
        });
        
        // Ricalcola i totali dopo la rimozione del valore
        calculateAndDisplayTotals();
    }
    
    // Setup delle tabelle e dei selettori di righe
    setupTablesAndSelectors();
    
    // Setup dei grafici
    setupChartEventHandlers();
});

// Funzione per ordinare le tabelle
function sortTable(table, sortKey, ascending) {
    const colIndex = getColumnIndex(table, sortKey);
    if (colIndex === -1) return;
    
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    
    // Se non ci sono righe, ritorna
    if (rows.length === 0) return;
    
    // Ordina le righe in base al contenuto della colonna
    rows.sort((a, b) => {
        const cellA = a.cells[colIndex].textContent.trim();
        const cellB = b.cells[colIndex].textContent.trim();
        
        // Gestisci diversi tipi di dati in base alla colonna
        if (sortKey === 'data' || sortKey === 'data-inizio' || sortKey === 'data-fine') {
            return compareDates(cellA, cellB) * (ascending ? 1 : -1);
        } else if (['quantita', 'prezzo-vendita', 'prezzo-acquisto', 'prezzo-unitario', 'prezzo-unitario-acquisto', 'prezzo-unitario-vendita', 'ricavo', 'ricavo-percentuale'].includes(sortKey)) {
            // Estrai solo i valori numerici (ignorando simboli di valuta, ecc.)
            const numA = extractNumericValue(cellA);
            const numB = extractNumericValue(cellB);
            return (numA - numB) * (ascending ? 1 : -1);
        } else {
            // Ordinamento testuale per le altre colonne
            return cellA.localeCompare(cellB, 'it', { sensitivity: 'base' }) * (ascending ? 1 : -1);
        }
    });
    
    // Rimuovi le righe esistenti e inserisci quelle ordinate
    const tbody = table.querySelector('tbody');
    while (tbody.firstChild) {
        tbody.removeChild(tbody.firstChild);
    }
    
    // Inserisci le righe ordinate
    rows.forEach(row => tbody.appendChild(row));
}

// Funzione per estrarre valori numerici dal testo
function extractNumericValue(value) {
    // Rimuovi tutti i caratteri non numerici tranne i separatori decimali
    const numericStr = value.replace(/[^\d,.-]/g, '')
                            .replace(/,/g, '.'); // Converti le virgole in punti
    return parseFloat(numericStr) || 0;
}

// Funzione per confrontare date
function compareDates(a, b) {
    // Supporto per diversi formati di data
    let dateA, dateB;
    
    try {
        // Tenta di convertire stringhe di data in oggetti Date
        dateA = convertDate(a);
        dateB = convertDate(b);
        
        // Confronta le date
        return dateA - dateB;
    } catch (error) {
        // Se la conversione fallisce, torna all'ordinamento testuale
        console.warn('Errore confrontando le date:', error);
        return a.localeCompare(b);
    }
}

// Funzione per ottenere l'indice della colonna dal data-sort
function getColumnIndex(table, sortKey) {
    const headers = table.querySelectorAll('th');
    for (let i = 0; i < headers.length; i++) {
        if (headers[i].getAttribute('data-sort') === sortKey) {
            return i;
        }
    }
    return -1;
}

// Funzioni di utilità adattate dal codice Python originale
function convertDate(dateStr) {
    if (!dateStr) return null;
    
    // Se è già un oggetto Date, restituiscilo
    if (dateStr instanceof Date) return dateStr;
    
    // Se è un numero, assumiamo sia un formato Excel
    if (!isNaN(dateStr)) {
        return new Date((dateStr - 25569) * 86400 * 1000);
    }
    
    // Per date nel formato "12 apr 24"
    if (typeof dateStr === 'string') {
        const dateParts = dateStr.split(' ');
        if (dateParts.length === 3) {
            const day = parseInt(dateParts[0]);
            const monthNames = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
            const month = monthNames.indexOf(dateParts[1].toLowerCase());
            const year = parseInt('20' + dateParts[2]);
            if (month !== -1 && !isNaN(day) && !isNaN(year)) {
                return new Date(year, month, day);
            }
        }
    }
    
    // Prova a convertire usando il costruttore Date standard
    try {
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) { // Verifica se la data è valida
            return date;
        }
    } catch (e) {
        // Ignora errori di parsing
    }
    
    // Se non siamo riusciti a convertire, restituisci null
    return null;
}

function separateOperation(text) {
    for (const opType of operationTypes) {
        if (text.startsWith(opType)) {
            const restDescription = text.substring(opType.length).trim();
            return [opType, restDescription];
        }
    }
    
    // Se non corrisponde a nessun tipo specificato, ritorna il testo originale e nulla
    return [text, ""];
}

function formatQuantity(text) {
    if (!text || !text.includes(", quantity")) return "";
    
    const parts = text.split(", quantity", 2);
    let quantityRaw = parts.length > 1 ? parts[1] : "";
    
    // Formatta la quantità
    if (quantityRaw) {
        quantityRaw = quantityRaw.trim();
        if (quantityRaw.startsWith(":")) {
            quantityRaw = quantityRaw.substring(1).trim();
        }
        
        // Converte il punto in virgola per numeri
        try {
            const numMatch = quantityRaw.match(/[-+]?\d*\.\d+|\d+/);
            if (numMatch) {
                const numberStr = numMatch[0];
                // Aggiungi la parola "azioni" dopo il numero
                return numberStr.replace('.', ',') + " azioni";
            } else {
                // Se non è un numero, aggiungi comunque "azioni" se non è già presente
                return quantityRaw.includes("azioni") ? quantityRaw : quantityRaw + " azioni";
            }
        } catch (e) {
            return quantityRaw.includes("azioni") ? quantityRaw : quantityRaw + " azioni";
        }
    }
    
    return "";
}

function convertNumber(value) {
    if (!value) return 0.0;
    
    // Se è già un numero, lo ritorna
    if (typeof value === 'number') return parseFloat(value);
    
    // Altrimenti, prova a convertire la stringa
    try {
        // Rimuove eventuali simboli di valuta e spazi
        let valueStr = String(value).trim();
        valueStr = valueStr.replace('€', '').replace(/\s/g, '').trim();
        
        // Sostituisce la virgola con il punto per la conversione
        valueStr = valueStr.replace(',', '.');
        
        return parseFloat(valueStr);
    } catch (e) {
        return 0.0;
    }
}

function formatNumber(value, decimals = 2) {
    if (value === null || value === undefined) return "";
    
    try {
        return value.toFixed(decimals).replace('.', ',');
    } catch (e) {
        return "";
    }
}

// Funzione per normalizzare il formato numerico
function normalizeNumericFormat(value, round = true, decimals = 3) {
    if (typeof value !== 'number') {
        if (typeof value === 'string') {
            // Se è una stringa, prova a convertirla in numero
            value = value.replace(',', '.'); // Sostituisce la virgola con il punto per la conversione
            value = parseFloat(value);
            if (isNaN(value)) return value; // Se non è un numero valido, ritorna il valore originale
        } else {
            return value;
        }
    }
    
    // Se valore è NaN dopo la conversione, ritorna 0
    if (isNaN(value)) {
        return 0;
    }
    
    // Arrotonda se richiesto
    if (round) {
        const factor = Math.pow(10, decimals);
        value = Math.round(value * factor) / factor;
    }
    
    // Formatta con virgola come separatore decimale
    return value.toString().replace('.', ',');
}

// Funzione per convertire una data seriale Excel in formato leggibile
function formatExcelDate(serialDate) {
    if (!serialDate || isNaN(serialDate)) return '';
    
    // La data Excel è il numero di giorni a partire dal 1-gen-1900
    // (con un errore noto: Excel considera erroneamente il 1900 come anno bisestile)
    const date = new Date((serialDate - 25569) * 86400 * 1000);
    
    // Array dei mesi abbreviati in italiano
    const mesi = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
    
    // Formatta la data come "12 Apr 24"
    const giorno = date.getDate();
    const mese = mesi[date.getMonth()];
    const anno = date.getFullYear().toString().slice(2); // Prendi solo le ultime 2 cifre
    
    return `${giorno} ${mese} ${anno}`;
}

// Funzione per popolare la tabella di output
function populateOutputTable(rows) {
    const tbody = document.querySelector('#output-table tbody');
    tbody.innerHTML = '';
    
    // Crea una riga per ogni elemento di data
    rows.forEach(row => {
        const tr = document.createElement('tr');
        
        // Formatta la data se è un numero (formato Excel)
        let formattedDate = row.DATA;
        if (!isNaN(row.DATA)) {
            formattedDate = formatExcelDate(row.DATA);
        }
        
        // Formatta i campi monetari assicurandosi di usare la virgola come separatore decimale
        let formattedPrezzoVendita = '';
        if (row.PREZZO_DI_VENDITA) {
            // Converte in numero e poi formatta usando la funzione che garantisce l'uso della virgola
            const valoreNumerico = convertNumber(row.PREZZO_DI_VENDITA);
            formattedPrezzoVendita = normalizeNumericFormat(valoreNumerico, true);
            // Aggiungi simbolo € se non c'è già
            if (!formattedPrezzoVendita.includes('€')) {
                formattedPrezzoVendita += ' €';
            }
        }
        
        let formattedPrezzoAcquisto = '';
        if (row.PREZZO_DI_ACQUISTO) {
            // Converte in numero e poi formatta usando la funzione che garantisce l'uso della virgola
            const valoreNumerico = convertNumber(row.PREZZO_DI_ACQUISTO);
            formattedPrezzoAcquisto = normalizeNumericFormat(valoreNumerico, true);
            // Aggiungi simbolo € se non c'è già
            if (!formattedPrezzoAcquisto.includes('€')) {
                formattedPrezzoAcquisto += ' €';
            }
        }
        
        // Formatta il prezzo unitario assicurandosi di usare la virgola
        let formattedPrezzoUnitario = '';
        if (row.PREZZO_UNITARIO) {
            // Converte in numero e poi formatta usando la funzione che garantisce l'uso della virgola
            const valoreNumerico = convertNumber(row.PREZZO_UNITARIO);
            formattedPrezzoUnitario = normalizeNumericFormat(valoreNumerico, true);
            // Aggiungi €/azione se non c'è già
            if (!formattedPrezzoUnitario.includes('/azione')) {
                formattedPrezzoUnitario += ' €/azione';
            }
        }
        
        // Correggi tipo e tipo operazione in base alle nuove regole richieste
        let tipo = row.TIPO || '';
        let tipoOperazione = row.TIPO_OPERAZIONE || '';
        
        // Per pagamenti interessi
        if (tipoOperazione.toLowerCase().includes('interest') || tipo.toLowerCase().includes('interessi')) {
            tipo = 'Pagamento degli interessi';
            tipoOperazione = 'Your interest payment';
        }
        // Per Saveback
        else if (tipoOperazione.toLowerCase().includes('saveback') || tipo.toLowerCase().includes('premio')) {
            tipo = 'Premio';
            tipoOperazione = 'Your Saveback payment';  
        }
        // Per operazioni di vendita (Sell trade)
        else if (tipoOperazione.toLowerCase().includes('sell')) {
            tipo = 'Commercio';
            tipoOperazione = 'Sell trade';
        }
        // Per operazioni di acquisto (Buy trade)
        else if (tipoOperazione.toLowerCase().includes('buy')) {
            tipo = 'Commercio';
            tipoOperazione = 'Buy trade';
        }
        
        // Crea celle per ogni colonna con attributo title per mostrare il testo completo
        tr.innerHTML = `
            <td title="${formattedDate || ''}">${formattedDate || ''}</td>
            <td title="${tipo}">${tipo}</td>
            <td title="${tipoOperazione}">${tipoOperazione}</td>
            <td title="${row.Descrizione || ''}">${row.Descrizione || ''}</td>
            <td title="${row.Quantita || ''}">${row.Quantita || ''}</td>
            <td title="${formattedPrezzoAcquisto || ''}">${formattedPrezzoAcquisto || ''}</td>
            <td title="${formattedPrezzoVendita || ''}">${formattedPrezzoVendita || ''}</td>
            <td title="${formattedPrezzoUnitario || ''}">${formattedPrezzoUnitario || ''}</td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Ordina la tabella per data (dalla più recente alla più lontana)
    sortTable(document.getElementById('output-table'), 'data', false); // Imposta l'ordinamento decrescente
    
    // Aggiorna la data del riepilogo
    updateSummaryDate();
}

// Funzione per trovare la data più recente nel dataset
function getLatestDate(data) {
    if (!data || data.length === 0) {
        return null;
    }
    
    // Trova tutte le date valide
    const dates = [];
    
    data.forEach(row => {
        if (row.DATA) {
            // Se è un numero (formato Excel), convertilo in una data
            if (!isNaN(row.DATA)) {
                const excelEpoch = new Date(1899, 11, 30); // Excel epoch è 30/12/1899
                const dateObj = new Date(excelEpoch.getTime() + (parseInt(row.DATA) * 24 * 60 * 60 * 1000));
                dates.push(dateObj);
            } 
            // Se è già una data in formato stringa, prova a convertirla
            else if (typeof row.DATA === 'string') {
                // Per date nel formato "12 apr 24"
                const dateParts = row.DATA.split(' ');
                if (dateParts.length === 3) {
                    const day = parseInt(dateParts[0]);
                    const monthNames = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
                    const month = monthNames.indexOf(dateParts[1].toLowerCase());
                    const year = parseInt('20' + dateParts[2]);
                    if (month !== -1 && !isNaN(day) && !isNaN(year)) {
                        dates.push(new Date(year, month, day));
                    }
                }
            }
        }
    });
    
    // Se non ci sono date valide, restituisci null
    if (dates.length === 0) {
        return null;
    }
    
    // Trova la data più recente
    const latestDate = new Date(Math.max.apply(null, dates));
    
    // Formatta la data nel formato "12 apr 24"
    const day = latestDate.getDate();
    const month = latestDate.toLocaleString('it', { month: 'short' });
    const year = latestDate.getFullYear().toString().substr(-2);
    
    return `${day} ${month} ${year}`;
}

// Funzione per aggiornare la data del riepilogo in base ai valori nella tabella
function updateSummaryDate() {
    // Ottieni tutte le righe della tabella di output, escluse quelle con classe added-value-row
    // Cioè escludiamo le righe dei valori aggiuntivi
    const rows = Array.from(document.querySelectorAll('#output-table tbody tr')).filter(row => !row.classList.contains('added-value-row'));
    const dates = [];
    
    // Estrai tutte le date dalle righe originali (non aggiunte manualmente)
    rows.forEach(row => {
        const dateCell = row.querySelector('td:first-child');
        if (dateCell) {
            const dateText = dateCell.textContent.trim();
            // Per date nel formato "12 apr 24"
            const dateParts = dateText.split(' ');
            if (dateParts.length === 3) {
                const day = parseInt(dateParts[0]);
                const monthNames = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
                const month = monthNames.indexOf(dateParts[1]);
                const year = parseInt('20' + dateParts[2]);
                if (month !== -1 && !isNaN(day) && !isNaN(year)) {
                    dates.push(new Date(year, month, day));
                }
            }
        }
    });
    
    // Se ci sono valori aggiunti attivi, usiamo la data odierna
    const hasAddedValues = additionalValues && additionalValues.length > 0;
    
    // Se non ci sono date valide nelle righe originali e non ci sono valori aggiunti, non fare nulla
    if (dates.length === 0 && !hasAddedValues) {
        return;
    }
    
    let formattedDate;
    
    if (hasAddedValues) {
        // Se ci sono valori aggiunti, usa la data odierna
        const today = new Date();
        const day = today.getDate();
        const month = today.toLocaleString('it', { month: 'short' });
        const year = today.getFullYear().toString().substr(-2);
        formattedDate = `${day} ${month} ${year}`;
    } else {
        // Se non ci sono valori aggiunti, trova la data più recente
        const latestDate = new Date(Math.max.apply(null, dates));
        
        // Formatta la data nel formato "12 apr 24"
        const day = latestDate.getDate();
        const month = latestDate.toLocaleString('it', { month: 'short' });
        const year = latestDate.getFullYear().toString().substr(-2);
        formattedDate = `${day} ${month} ${year}`;
    }
    
    // Aggiorna il titolo del riepilogo
    const summaryTitle = document.querySelector('.summary-container h2');
    if (summaryTitle) {
        summaryTitle.textContent = `Riepilogo Ricavi al ${formattedDate}`;
    }
}

// Modifica la funzione processExcelFile per aggiornare il titolo del riepilogo
function processExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Prendi il primo foglio
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // Converti in JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            // Estrai le intestazioni
            const headers = jsonData[0].map(header => header ? String(header).trim() : "");
            
            // Trova gli indici delle colonne IN ENTRATA e IN USCITA
            let inEntrataIdx, inUscitaIdx;
            
            headers.forEach((header, idx) => {
                const headerUpper = header ? header.toUpperCase() : "";
                if (headerUpper.includes("IN ENTRATA")) {
                    inEntrataIdx = idx;
                } else if (headerUpper.includes("IN USCITA")) {
                    inUscitaIdx = idx;
                }
            });
            
            // Se gli indici non sono stati trovati, assumiamo che siano nelle posizioni 4 e 5
            if (inEntrataIdx === undefined || inUscitaIdx === undefined) {
                inEntrataIdx = 4;
                inUscitaIdx = 5;
            }
            
            // Raccogli i dati validi
            const validRows = [];
            
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row || row.length === 0) continue;
                
                // Controlla se il tipo è nei tipi ammessi
                const tipo = row[1];
                if (!tipo || !allowedTypes.includes(tipo)) continue;
                
                // La data è nella colonna A
                const dataStr = row[0] ? String(row[0]) : "";
                const data = convertDate(dataStr);
                
                // Aggiungi la riga ai dati da ordinare
                validRows.push([row, tipo, data]);
            }
            
            // Ordina le righe: prima per tipo, poi per data
            validRows.sort((a, b) => {
                // Ordina prima per tipo
                if (a[1] < b[1]) return -1;
                if (a[1] > b[1]) return 1;
                
                // Se i tipi sono uguali, ordina per data
                return a[2] - b[2];
            });
            
            // Processa le righe ordinate
            const rows = [];
            
            for (const [rowData, tipoVal, data] of validRows) {
                // Estraggo i valori dalle celle originali
                const dataVal = rowData[0]; // DATA (colonna A)
                
                const descrizioneOriginale = rowData[2] ? String(rowData[2]) : "";
                
                // Separo il tipo di operazione dal resto della descrizione
                const [tipoOp, restoDescrizione] = separateOperation(descrizioneOriginale);
                
                // Estraggo e formatto la quantità
                const quantita = formatQuantity(descrizioneOriginale);
                
                // Rimuovo la parte ", quantity:" dalla descrizione se presente
                let descrizioneFinale = restoDescrizione;
                if (restoDescrizione.includes(", quantity")) {
                    descrizioneFinale = restoDescrizione.split(", quantity")[0];
                }
                
                // Estraggo i valori di IN ENTRATA e IN USCITA usando gli indici corretti
                const prezzoVendita = rowData[inEntrataIdx] !== undefined ? rowData[inEntrataIdx] : "";
                const prezzoAcquisto = rowData[inUscitaIdx] !== undefined ? rowData[inUscitaIdx] : "";
                
                // Normalizzo i valori di entrata/uscita (rimuovo eventuali simboli € e formatto)
                let prezzoVenditaFormatted = "";
                let prezzoAcquistoFormatted = "";
                
                if (prezzoVendita) {
                    prezzoVenditaFormatted = String(prezzoVendita).replace('€', '').trim();
                }
                
                if (prezzoAcquisto) {
                    prezzoAcquistoFormatted = String(prezzoAcquisto).replace('€', '').trim();
                }
                
                // Calcolo del prezzo unitario
                let prezzoUnitario = "";
                try {
                    // Converto quantità in numero
                    const quantitaNum = convertNumber(quantita);
                    
                    if (quantitaNum > 0) {
                        // Determino quale importo usare (entrata o uscita)
                        if (prezzoAcquistoFormatted && String(prezzoAcquistoFormatted).trim()) {
                            const importo = convertNumber(prezzoAcquistoFormatted);
                            if (importo > 0) {
                                // Prezzo = Importo pagato/Quantità
                                prezzoUnitario = formatNumber(importo/quantitaNum);
                            }
                        } else if (prezzoVenditaFormatted && String(prezzoVenditaFormatted).trim()) {
                            const importo = convertNumber(prezzoVenditaFormatted);
                            if (importo > 0) {
                                // Se è un'entrata, il prezzo è ancora Importo/Quantità
                                prezzoUnitario = formatNumber(importo/quantitaNum);
                            }
                        }
                    }
                } catch (e) {
                    console.error("Errore nel calcolo del prezzo unitario:", e);
                }
                
                // Aggiungo la riga formattata
                rows.push({
                    DATA: dataVal,
                    TIPO: tipoVal,
                    TIPO_OPERAZIONE: tipoOp,
                    Descrizione: descrizioneFinale,
                    Quantita: quantita,
                    PREZZO_DI_VENDITA: prezzoVenditaFormatted,
                    PREZZO_DI_ACQUISTO: prezzoAcquistoFormatted,
                    PREZZO_UNITARIO: prezzoUnitario
                });
            }
            
            // Salva i dati di output
            outputData = rows;
            
            // Processa i dati per il file finale
            processFinalData(rows);
            
            // Assicuriamoci che l'overlay di caricamento sia nascosto
            const loadingOverlay = document.querySelector('.loading-overlay');
            if (loadingOverlay) {
                loadingOverlay.style.display = 'none';
            }
            
            // Mostra la sezione dei valori aggiuntivi e dei risultati
            const additionalValuesSection = document.querySelector('.additional-values-section');
            const resultsSection = document.querySelector('.results-section');
            
            if (additionalValuesSection) {
                additionalValuesSection.style.display = 'block';
            }
            
            if (resultsSection) {
                resultsSection.style.display = 'block';
            }
            
            // Popola le tabelle e i riepiloghi
            populateOutputTable(rows);
            populateFinalTable();
            
            // Calcola e mostra i totali
            calculateAndDisplayTotals();
            
            // Aggiorna il titolo del riepilogo con la data più recente
            const latestDate = getLatestDate(rows);
            const summaryTitle = document.querySelector('.summary-container h2');
            if (summaryTitle && latestDate) {
                summaryTitle.textContent = `Riepilogo Ricavi al ${latestDate}`;
            }
            
            // Aggiorna il selettore di righe per i valori aggiuntivi
            console.log("Tentativo di aggiornamento del selettore dopo il caricamento della tabella finale");
            
            // Aggiungiamo un piccolo ritardo per assicurarci che tutto sia caricato
            setTimeout(() => {
                if (window.updateRowSelector && typeof window.updateRowSelector === 'function') {
                    console.log("Aggiornamento del selettore di righe...");
                    window.updateRowSelector();
                } else {
                    console.error("updateRowSelector non disponibile o non è una funzione");
                    // Tentiamo di recuperarla
                    if (typeof setupAdditionalValuesHandling === 'function') {
                        window.updateRowSelector = setupAdditionalValuesHandling();
                        if (window.updateRowSelector) {
                            window.updateRowSelector();
                        }
                    }
                }
            }, 200);
        } catch (error) {
            console.error("Errore durante l'elaborazione del file:", error);
            alert("Si è verificato un errore durante l'elaborazione del file. Controlla la console per i dettagli.");
            
            // Assicuriamoci che l'overlay di caricamento sia nascosto anche in caso di errore
            const loadingOverlay = document.querySelector('.loading-overlay');
            if (loadingOverlay) {
                loadingOverlay.style.display = 'none';
            }
        }
    };
    
    reader.onerror = function() {
        console.error("Errore nella lettura del file");
        alert("Si è verificato un errore nella lettura del file.");
        
        // Assicuriamoci che l'overlay di caricamento sia nascosto anche in caso di errore
        const loadingOverlay = document.querySelector('.loading-overlay');
        if (loadingOverlay) {
            loadingOverlay.style.display = 'none';
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Funzione per processare i dati finali
function processFinalData(rows) {
    // Inizializza l'oggetto dati finale
    finalData = {
        dfCommercioBuy: [],
        dfCommercioSell: [],
        dfInteressi: [],
        dfPremio: [],
        dfRedditi: [],
        dfWarrant: []
    };
    
    // Filtra i dati per tipo
    const dfCommercio = rows.filter(row => row.TIPO === 'Commercio');
    const dfInteressi = rows.filter(row => row.TIPO === 'Pagamento degli interessi');
    const dfPremio = rows.filter(row => row.TIPO === 'Premio');
    const dfRedditi = rows.filter(row => row.TIPO === 'Redditi');
    
    // Separa i dati in entrata e in uscita
    const dfIn = dfCommercio.filter(row => row.PREZZO_DI_VENDITA && row.PREZZO_DI_VENDITA !== '');
    const dfOut = dfCommercio.filter(row => row.PREZZO_DI_ACQUISTO && row.PREZZO_DI_ACQUISTO !== '');
    
    // Aggrega i dati in entrata per tipo operazione e descrizione
    const dfInAgg = aggregateCommercio(dfIn, 'PREZZO_DI_VENDITA');
    const dfOutAgg = aggregateCommercio(dfOut, 'PREZZO_DI_ACQUISTO');
    
    // Aggiungi la colonna TIPO
    dfInAgg.forEach(row => {
        row.TIPO = 'Commercio';
        row.PREZZO_DI_ACQUISTO = '';
    });
    
    dfOutAgg.forEach(row => {
        row.TIPO = 'Commercio';
        row.PREZZO_DI_VENDITA = '';
    });
    
    // Unisci i dataframe
    const dfCommercioAgg = [...dfInAgg, ...dfOutAgg];
    
    // Gestione "Pagamento degli interessi"
    let dfInteressiAgg = [];
    if (dfInteressi.length > 0) {
        // Somma tutti i pagamenti di interessi
        const sommaInteressi = dfInteressi.reduce((total, row) => {
            return total + convertNumber(row.PREZZO_DI_VENDITA);
        }, 0);
        
        // Crea una singola riga
        dfInteressiAgg = [{
            TIPO: 'Pagamento degli interessi',
            TIPO_OPERAZIONE: 'Your interest payment',
            Descrizione: '',
            Quantita: '',
            PREZZO_DI_VENDITA: sommaInteressi.toString(),
            PREZZO_DI_ACQUISTO: '',
            PREZZO_UNITARIO: ''
        }];
    }
    
    // Gestione "Premio"
    let dfPremioAgg = [];
    if (dfPremio.length > 0) {
        // Somma tutti i premi
        const sommaPremio = dfPremio.reduce((total, row) => {
            return total + convertNumber(row.PREZZO_DI_VENDITA);
        }, 0);
        
        // Crea una singola riga
        dfPremioAgg = [{
            TIPO: 'Premio',
            TIPO_OPERAZIONE: 'Your Saveback payment',
            Descrizione: '',
            Quantita: '',
            PREZZO_DI_VENDITA: sommaPremio.toString(),
            PREZZO_DI_ACQUISTO: '',
            PREZZO_UNITARIO: ''
        }];
    }
    
    // Salva i dati nei rispettivi array
    const dfFinal = [...dfCommercioAgg, ...dfInteressiAgg, ...dfPremioAgg, ...dfRedditi];
    
    // Ora separiamo per le categorie specifiche per il risultato finale
    const dfBuy = dfFinal.filter(row => 
        row.TIPO === 'Commercio' && row.TIPO_OPERAZIONE === 'Buy trade'
    );
    
    const dfSell = dfFinal.filter(row => 
        row.TIPO === 'Commercio' && row.TIPO_OPERAZIONE === 'Sell trade'
    );
    
    const dfSavings = dfFinal.filter(row => 
        row.TIPO_OPERAZIONE === 'Savings plan execution'
    );
    
    // Copia dfSavings come se fosse un "Buy trade"
    const dfSavingsForBuy = dfSavings.map(row => {
        const newRow = {...row};
        newRow.TIPO_OPERAZIONE = 'Buy trade';
        return newRow;
    });
    
    // Unisce i PAC accumulo ai Buy trade
    const dfBuyWithSavings = [...dfBuy, ...dfSavingsForBuy];
    
    // Ottieni i warrant exercise
    const dfWarrant = dfFinal.filter(row => 
        row.TIPO_OPERAZIONE === 'Warrant Exercise'
    );
    
    // Salva i dati intermedi
    finalData.dfCommercioBuy = dfBuyWithSavings;
    finalData.dfCommercioSell = dfSell;
    finalData.dfInteressi = dfInteressiAgg;
    finalData.dfPremio = dfPremioAgg;
    finalData.dfRedditi = dfRedditi;
    finalData.dfWarrant = dfWarrant;
}

// Funzione per aggregare i dati commercio
function aggregateCommercio(data, entryCol) {
    // Oggetto per memorizzare i gruppi
    const groups = {};
    
    // Raggruppa per TIPO_OPERAZIONE e Descrizione
    data.forEach(row => {
        const key = `${row.TIPO_OPERAZIONE}|${row.Descrizione}`;
        
        if (!groups[key]) {
            groups[key] = {
                rows: [],
                totalQty: 0,
                totalValue: 0,
                weights: [],
                prices: []
            };
        }
        
        // Converti i valori numerici
        const qty = convertNumber(row.Quantita);
        const value = convertNumber(row[entryCol]);
        const unitPrice = convertNumber(row.PREZZO_UNITARIO);
        
        groups[key].rows.push(row);
        groups[key].totalQty += qty;
        groups[key].totalValue += value;
        
        // Solo se la quantità è valida, aggiungi ai pesi per la media ponderata
        if (qty > 0) {
            groups[key].weights.push(qty);
            groups[key].prices.push(unitPrice);
        }
    });
    
    // Calcola i valori aggregati per ogni gruppo
    const result = [];
    
    for (const key in groups) {
        const group = groups[key];
        const [tipoOp, desc] = key.split('|');
        
        // Calcola il prezzo unitario medio ponderato
        let avgPrice = 0;
        if (group.weights.length > 0) {
            let weightSum = 0;
            let weightedSum = 0;
            
            for (let i = 0; i < group.weights.length; i++) {
                weightSum += group.weights[i];
                weightedSum += group.weights[i] * group.prices[i];
            }
            
            if (weightSum > 0) {
                avgPrice = weightedSum / weightSum;
            }
        }
        
        result.push({
            TIPO_OPERAZIONE: tipoOp,
            Descrizione: desc,
            Quantita: normalizeNumericFormat(group.totalQty, true),
            [entryCol]: normalizeNumericFormat(group.totalValue, true),
            PREZZO_UNITARIO: normalizeNumericFormat(avgPrice, true)
        });
    }
    
    return result;
}

// Funzione per popolare la tabella finale
function populateFinalTable() {
    // Dati dal riepilogo finale
    const risultatoFinale = createFinalResult();
    
    // Popola la tabella finale
    const tbody = document.querySelector('#final-table tbody');
    tbody.innerHTML = '';
    
    risultatoFinale.forEach(row => {
        const tr = document.createElement('tr');
        
        // Aggiungi classe in base al tipo di operazione o al ricavo
        if (row.TIPO_OPERAZIONE === 'Interessi Giacenza') {
            tr.classList.add('type-interest');
        } else if (row.TIPO_OPERAZIONE === 'Saveback') {
            tr.classList.add('type-saveback');
        } else if (row.TIPO_OPERAZIONE === 'Buy trade') {
            tr.classList.add('type-buy-trade');
        } else if (row.TIPO_OPERAZIONE === 'Sell trade Parziale') {
            tr.classList.add('type-sell-trade-partial');
        } else if (row.RICAVO) {
            // Estrai il valore numerico del ricavo (rimuovendo il simbolo €)
            const ricavoStr = row.RICAVO.replace('€', '').trim();
            const ricavoVal = convertNumber(ricavoStr);
            
            // Aggiungi la classe appropriata in base al valore
            if (ricavoVal > 0) {
                tr.classList.add('profit-positive');
            } else if (ricavoVal < 0) {
                tr.classList.add('profit-negative');
            }
        }
        
        // Modifica il valore del ricavo per Buy trade e Sell trade Parziale
        if (row.TIPO_OPERAZIONE === 'Buy trade' || row.TIPO_OPERAZIONE === 'Sell trade Parziale') {
            // Aggiungi un punto interrogativo arancione in grassetto per indicare che il ricavo non è ancora calcolato
            row.RICAVO = '<span class="pending-profit">?</span>';
            // Lasciamo vuoto il ricavo percentuale
            row.RICAVO_PERCENTUALE = '';
        }
        
        // Crea celle per ogni colonna con attributo title e possibilità di andare a capo
        tr.innerHTML = `
            <td title="${row.DATA_INIZIO || ''}">${row.DATA_INIZIO || ''}</td>
            <td title="${row.DATA_FINE || ''}">${row.DATA_FINE || ''}</td>
            <td title="${row.TIPO_OPERAZIONE || ''}">${row.TIPO_OPERAZIONE || ''}</td>
            <td class="desc-cell" title="${row.Descrizione || ''}">${row.Descrizione || ''}</td>
            <td title="${row.Quantita || ''}">${row.Quantita || ''}</td>
            <td title="${row.PREZZO_DI_ACQUISTO || ''}">${row.PREZZO_DI_ACQUISTO || ''}</td>
            <td title="${row.PREZZO_DI_VENDITA || ''}">${row.PREZZO_DI_VENDITA || ''}</td>
            <td title="${row.PREZZO_UNITARIO_ACQUISTO || ''}">${row.PREZZO_UNITARIO_ACQUISTO || ''}</td>
            <td title="${row.PREZZO_UNITARIO_VENDITA || ''}">${row.PREZZO_UNITARIO_VENDITA || ''}</td>
            <td title="${row.RICAVO && row.RICAVO.replace ? row.RICAVO.replace(/<[^>]*>/g, '') : row.RICAVO || ''}">${row.RICAVO || ''}</td>
            <td title="${row.RICAVO_PERCENTUALE && row.RICAVO_PERCENTUALE.replace ? row.RICAVO_PERCENTUALE.replace(/<[^>]*>/g, '') : row.RICAVO_PERCENTUALE || ''}">${row.RICAVO_PERCENTUALE || ''}</td>
        `;
        
        tbody.appendChild(tr);
        
        // Applica le classi CSS per colorare solo i valori di ricavo
        const cells = tr.querySelectorAll('td');
        if (cells.length >= 10) {
            const ricavoCell = cells[9]; // 10ª colonna (indice 9)
            
            if (row.RICAVO) {
                const ricavoStr = row.RICAVO.replace('€', '').trim();
                const ricavoVal = convertNumber(ricavoStr);
                
                if (ricavoVal > 0) {
                    ricavoCell.classList.add('positive-value');
                } else if (ricavoVal < 0) {
                    ricavoCell.classList.add('negative-value');
                }
            }
        }
    });
    
    // Ordina la tabella per data di fine (dalla più recente alla più lontana)
    sortTable(document.getElementById('final-table'), 'data-fine', false); // Imposta l'ordinamento decrescente
}

// Funzione per creare il risultato finale
function createFinalResult() {
    const risultatoFinale = [];
    
    // Estraggo gli ISIN dai warrant exercise per l'abbinamento
    const warrantIsinDict = {};
    
    finalData.dfWarrant.forEach(row => {
        if (row.Descrizione && row.Descrizione.includes('ISIN')) {
            // Estrai l'ISIN dalla descrizione (formato: "for ISIN DE000XXX")
            const isin = row.Descrizione.split('ISIN ')[1].trim();
            const inEntrata = convertNumber(row.PREZZO_DI_VENDITA);
            warrantIsinDict[isin] = inEntrata;
        }
    });
    
    // Unisci buy e sell per descrizione
    const descUnicheSet = new Set([
        ...finalData.dfCommercioBuy.map(row => row.Descrizione),
        ...finalData.dfCommercioSell.map(row => row.Descrizione)
    ]);
    
    // Converti il Set in Array per iterare
    const descrizioniUniche = Array.from(descUnicheSet);
    
    // Ottieni tutte le righe originali dal dettaglio operazioni per calcolare le date
    const outputTableRows = document.querySelectorAll('#output-table tbody tr:not(.added-value-row)');
    
    // Funzione per estrarre la data di una riga
    const estraiData = (row) => {
        const dateCell = row.querySelector('td:first-child');
        if (dateCell) {
            const dateText = dateCell.textContent.trim();
            // Per date nel formato "12 apr 24"
            const dateParts = dateText.split(' ');
            if (dateParts.length === 3) {
                const day = parseInt(dateParts[0]);
                const monthNames = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
                const month = monthNames.indexOf(dateParts[1]);
                const year = parseInt('20' + dateParts[2]);
                if (month !== -1 && !isNaN(day) && !isNaN(year)) {
                    return new Date(year, month, day);
                }
            }
        }
        return null;
    };
    
    // Funzione per formattare la data nel formato "12 apr 24"
    const formatDate = (date) => {
        if (!date) return "";
        const day = date.getDate();
        const monthNames = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear().toString().substr(-2);
        return `${day} ${month} ${year}`;
    };
    
    // Funzione per trovare date min e max per una descrizione
    const trovateDate = (descrizione) => {
        const result = {
            dataInizio: null,
            dataFine: null,
        };
        
        // Estrai l'ISIN dalla descrizione (se presente)
        let isin = null;
        const descParts = descrizione.split(' ');
        if (descParts.length > 0) {
            isin = descParts[0]; // Il primo elemento è solitamente l'ISIN
        }
        
        // Cerca la descrizione in tutte le righe del dettaglio operazioni
        const dateDiAcquisto = [];
        const dateDiVendita = [];
        
        outputTableRows.forEach(row => {
            const descCell = row.querySelector('td:nth-child(4)');
            const tipoCell = row.querySelector('td:nth-child(2)');
            const tipoOperazioneCell = row.querySelector('td:nth-child(3)');
            
            // Controlla se la descrizione corrisponde esattamente o se è un Warrant Exercise contenente l'ISIN
            const isMatch = 
                (descCell && descCell.textContent.trim() === descrizione) || 
                (isin && tipoCell && tipoOperazioneCell && 
                 tipoCell.textContent.trim() === 'Redditi' && 
                 tipoOperazioneCell.textContent.trim() === 'Warrant Exercise' && 
                 descCell && descCell.textContent.trim().includes(isin));
            
            if (isMatch) {
                const date = estraiData(row);
                
                // Determina se è un acquisto o una vendita
                const tipoOperazione = tipoOperazioneCell ? tipoOperazioneCell.textContent.trim() : '';
                
                if (date) {
                    // Aggiungi alle date di acquisto
                    if (tipoOperazione === 'Buy trade' || tipoOperazione === 'Savings plan execution') {
                        dateDiAcquisto.push(date);
                    }
                    // Aggiungi alle date di vendita
                    else if (tipoOperazione === 'Sell trade' || tipoOperazione === 'Warrant Exercise') {
                        dateDiVendita.push(date);
                    }
                    // Per altri tipi, aggiungi a entrambi (casi non specificati)
                    else {
                        dateDiAcquisto.push(date);
                    }
                }
            }
        });
        
        // Imposta la data di inizio come la prima data di acquisto, se presente
        if (dateDiAcquisto.length > 0) {
            result.dataInizio = new Date(Math.min.apply(null, dateDiAcquisto));
        }
        
        // Imposta la data di fine come l'ultima data di vendita, se presente
        if (dateDiVendita.length > 0) {
            result.dataFine = new Date(Math.max.apply(null, dateDiVendita));
        } else {
            // Se non ci sono vendite, lascia la data di fine come null
            result.dataFine = null;
        }
        
        return result;
    };
    
    for (const descrizione of descrizioniUniche) {
        const buyRows = finalData.dfCommercioBuy.filter(row => row.Descrizione === descrizione);
        const sellRow = finalData.dfCommercioSell.find(row => row.Descrizione === descrizione);
        
        // Calcola date min e max per questa descrizione
        const { dataInizio, dataFine } = trovateDate(descrizione);
        const dataInizioFormatted = formatDate(dataInizio);
        const dataFineFormatted = dataFine ? formatDate(dataFine) : '';
        
        // Estrai l'ISIN o il codice identificativo dal buy_row (se esiste)
        let isinToMatch = null;
        if (buyRows.length > 0) {
            // Il primo elemento della descrizione è generalmente l'ISIN
            const descParts = descrizione.split(' ');
            if (descParts.length > 0) {
                // Verifica se sembra un ISIN (inizia con DE, CH, US, ecc. seguito da numeri/lettere)
                const isinPattern = /^([A-Z]{2}\d{10}|[A-Z]{2}[0-9A-Z]{9}\d)$/;
                if (isinPattern.test(descParts[0])) {
                    isinToMatch = descParts[0];
                }
            }
        }
        
        // Cerchiamo se c'è un warrant exercise corrispondente
        let warrantValue = 0;
        if (isinToMatch && warrantIsinDict[isinToMatch]) {
            warrantValue = warrantIsinDict[isinToMatch];
        }
        
        // Se abbiamo sia buy che sell per questo strumento
        if (buyRows.length > 0 && sellRow) {
            // Sommiamo tutte le quantità e prezzi di acquisto dai buy_rows
            let totalQty = 0;
            let totalOut = 0;
            
            for (const buyRow of buyRows) {
                const rowQty = convertNumber(buyRow.Quantita);
                const rowOut = convertNumber(buyRow.PREZZO_DI_ACQUISTO);
                totalQty += rowQty;
                totalOut += rowOut;
            }
            
            // Calcoliamo il prezzo unitario medio ponderato di acquisto
            const avgPriceBuy = totalQty > 0 ? totalOut / totalQty : 0;
            
            const inEntrata = sellRow.PREZZO_DI_VENDITA || '';
            
            // Convertiamo per calcolare il ricavo
            const inEntrataVal = convertNumber(inEntrata);
            
            // Calcolo del ricavo (IN ENTRATA - IN USCITA)
            const ricavo = inEntrataVal - totalOut;
            const ricavoStr = ricavo !== 0 ? normalizeNumericFormat(ricavo, true) : '';
            
            // Calcolo del ricavo percentuale
            let ricavoPercentuale = 0;
            if (totalOut > 0) {
                ricavoPercentuale = (ricavo / totalOut) * 100;
            }
            const ricavoPercStr = ricavoPercentuale !== 0 ? normalizeNumericFormat(ricavoPercentuale, true) : '';
            
            // Prezzo unitario di vendita
            const prezzoSell = sellRow.PREZZO_UNITARIO || '';
            
            // NUOVO: Calcola la quantità effettivamente venduta
            const soldQty = convertNumber(sellRow.Quantita);
            
            // NUOVO: Verifica se è una vendita parziale (quantità venduta < quantità acquistata)
            if (soldQty < totalQty) {
                // È una vendita parziale - creiamo una singola riga
                risultatoFinale.push({
                    DATA_INIZIO: dataInizioFormatted,
                    DATA_FINE: dataFineFormatted,
                    TIPO_OPERAZIONE: 'Sell trade Parziale',
                    Descrizione: descrizione,
                    Quantita: `${normalizeNumericFormat(totalQty, true)} azioni comprate\n${normalizeNumericFormat(soldQty, true)} azioni vendute`,
                    PREZZO_DI_ACQUISTO: normalizeNumericFormat(totalOut, true) + " €",
                    PREZZO_DI_VENDITA: inEntrata,
                    PREZZO_UNITARIO_ACQUISTO: normalizeNumericFormat(avgPriceBuy, true) + " €/azione",
                    PREZZO_UNITARIO_VENDITA: prezzoSell,
                    RICAVO: '',  // Rimosso il calcolo del ricavo per le vendite parziali
                    RICAVO_PERCENTUALE: ''  // Rimosso il calcolo percentuale
                });
                
                // Nota: non aggiungiamo più la riga Buy trade per le azioni rimanenti
                
            } else {
                // Vendita completa (come prima)
                risultatoFinale.push({
                    DATA_INIZIO: dataInizioFormatted,
                    DATA_FINE: dataFineFormatted,
                    TIPO_OPERAZIONE: 'Trade',
                    Descrizione: descrizione,
                    Quantita: normalizeNumericFormat(totalQty, true) + " azioni",
                    PREZZO_DI_ACQUISTO: normalizeNumericFormat(totalOut, true) + " €",
                    PREZZO_DI_VENDITA: inEntrata,
                    PREZZO_UNITARIO_ACQUISTO: normalizeNumericFormat(avgPriceBuy, true) + " €/azione",
                    PREZZO_UNITARIO_VENDITA: prezzoSell,
                    RICAVO: ricavoStr ? ricavoStr + " €" : "",
                    RICAVO_PERCENTUALE: ricavoPercStr ? ricavoPercStr + " %" : ""
                });
            }
        } else {
            // Se abbiamo solo buy o solo sell
            if (buyRows.length > 0) {
                // Sommiamo tutte le quantità e prezzi di acquisto dai buy_rows
                let totalQty = 0;
                let totalOut = 0;
                
                for (const buyRow of buyRows) {
                    const rowQty = convertNumber(buyRow.Quantita);
                    const rowOut = convertNumber(buyRow.PREZZO_DI_ACQUISTO);
                    totalQty += rowQty;
                    totalOut += rowOut;
                }
                
                // Calcoliamo il prezzo unitario medio ponderato
                const avgPriceBuy = totalQty > 0 ? totalOut / totalQty : 0;
                
                // Controlliamo se questo buy ha un warrant exercise corrispondente
                if (warrantValue > 0) {
                    // Calcoliamo il prezzo unitario di vendita
                    const prezzoVendita = totalQty > 0 ? warrantValue / totalQty : 0;
                    // Calcoliamo il ricavo
                    const ricavo = warrantValue - totalOut;
                    
                    // Calcolo del ricavo percentuale
                    let ricavoPercentuale = 0;
                    if (totalOut > 0) {
                        ricavoPercentuale = (ricavo / totalOut) * 100;
                    }
                    const ricavoPercStr = ricavoPercentuale !== 0 ? normalizeNumericFormat(ricavoPercentuale, true) : '';
                    
                    risultatoFinale.push({
                        DATA_INIZIO: dataInizioFormatted,
                        DATA_FINE: dataFineFormatted,
                        TIPO_OPERAZIONE: 'Trade K.O.',
                        Descrizione: descrizione,
                        Quantita: normalizeNumericFormat(totalQty, true) + " azioni",
                        PREZZO_DI_ACQUISTO: normalizeNumericFormat(totalOut, true) + " €",
                        PREZZO_DI_VENDITA: normalizeNumericFormat(warrantValue, true),
                        PREZZO_UNITARIO_ACQUISTO: normalizeNumericFormat(avgPriceBuy, true) + " €/azione",
                        PREZZO_UNITARIO_VENDITA: normalizeNumericFormat(prezzoVendita, true),
                        RICAVO: normalizeNumericFormat(ricavo, true) + " €",
                        RICAVO_PERCENTUALE: ricavoPercStr ? ricavoPercStr + " %" : ""
                    });
                } else {
                    risultatoFinale.push({
                        DATA_INIZIO: dataInizioFormatted,
                        DATA_FINE: '', // Lasciamo il campo DATA_FINE vuoto per i Buy trade
                        TIPO_OPERAZIONE: 'Buy trade',
                        Descrizione: descrizione,
                        Quantita: normalizeNumericFormat(totalQty, true) + " azioni",
                        PREZZO_DI_ACQUISTO: normalizeNumericFormat(totalOut, true) + " €",
                        PREZZO_DI_VENDITA: '',
                        PREZZO_UNITARIO_ACQUISTO: normalizeNumericFormat(avgPriceBuy, true) + " €/azione",
                        PREZZO_UNITARIO_VENDITA: '',
                        RICAVO: '',
                        RICAVO_PERCENTUALE: ''
                    });
                }
            }
            if (sellRow) {
                risultatoFinale.push({
                    DATA_INIZIO: dataInizioFormatted,
                    DATA_FINE: dataFineFormatted,
                    TIPO_OPERAZIONE: 'Sell trade',
                    Descrizione: sellRow.Descrizione,
                    Quantita: sellRow.Quantita && sellRow.Quantita.includes("azioni") ? sellRow.Quantita : (sellRow.Quantita || "") + " azioni",
                    PREZZO_DI_ACQUISTO: '',
                    PREZZO_DI_VENDITA: sellRow.PREZZO_DI_VENDITA && sellRow.PREZZO_DI_VENDITA.includes("€") ? sellRow.PREZZO_DI_VENDITA : (sellRow.PREZZO_DI_VENDITA || "") + " €",
                    PREZZO_UNITARIO_ACQUISTO: '',
                    PREZZO_UNITARIO_VENDITA: sellRow.PREZZO_UNITARIO && sellRow.PREZZO_UNITARIO.includes("€/azione") ? sellRow.PREZZO_UNITARIO : (sellRow.PREZZO_UNITARIO || "") + " €/azione",
                    RICAVO: '',
                    RICAVO_PERCENTUALE: ''
                });
            }
        }
    }
    
    // Aggiungiamo le righe per Your interest payment
    for (const row of finalData.dfInteressi) {
        // Assicuriamoci che il valore sia arrotondato prima di aggiungerlo
        let prezzoVendita = row.PREZZO_DI_VENDITA;
        if (prezzoVendita) {
            // Converto in numero, arrotondo e riconverto in stringa formattata
            const valoreNumerico = convertNumber(prezzoVendita);
            prezzoVendita = normalizeNumericFormat(valoreNumerico, true);
        }
        
        // Trova le date minima e massima per gli interessi
        let dataMinInteressi = "";
        let dataMaxInteressi = "";
        const dateInteressi = [];
        
        // Raccogli tutte le date degli interessi
        outputTableRows.forEach(row => {
            const tipoCell = row.querySelector('td:nth-child(3)');
            if (tipoCell && tipoCell.textContent.trim() === 'Your interest payment') {
                const date = estraiData(row);
                if (date) {
                    dateInteressi.push(date);
                }
            }
        });
        
        // Se abbiamo trovato date, calcoliamo min e max
        if (dateInteressi.length > 0) {
            dataMinInteressi = formatDate(new Date(Math.min.apply(null, dateInteressi)));
            dataMaxInteressi = formatDate(new Date(Math.max.apply(null, dateInteressi)));
        }
        
        // MODIFICATO: Sposta il valore da ricavo a prezzo di vendita
        const prezzoVenditaFormatted = prezzoVendita ? prezzoVendita + " €" : "";
        
        risultatoFinale.push({
            DATA_INIZIO: dataMinInteressi,
            DATA_FINE: dataMaxInteressi,
            TIPO_OPERAZIONE: 'Interessi Giacenza',
            Descrizione: row.Descrizione,
            Quantita: row.Quantita,
            PREZZO_DI_ACQUISTO: "0 €", // Imposta a 0 € invece che vuoto
            PREZZO_DI_VENDITA: prezzoVenditaFormatted, // Sposta qui il valore che era nel ricavo
            PREZZO_UNITARIO_ACQUISTO: '',
            PREZZO_UNITARIO_VENDITA: '',
            RICAVO: prezzoVenditaFormatted, // Mantiene il ricavo come prima (uguale al prezzo di vendita)
            RICAVO_PERCENTUALE: '' // Non calcoliamo percentuale perché il prezzo di acquisto è 0
        });
    }
    
    // Aggiungiamo le righe per Your Saveback payment
    for (const row of finalData.dfPremio) {
        // Assicuriamoci che il valore sia arrotondato prima di aggiungerlo
        let prezzoVendita = row.PREZZO_DI_VENDITA;
        if (prezzoVendita) {
            // Converto in numero, arrotondo e riconverto in stringa formattata
            const valoreNumerico = convertNumber(prezzoVendita);
            prezzoVendita = normalizeNumericFormat(valoreNumerico, true);
        }
        
        // Trova le date minima e massima per saveback
        let dataMinSaveback = "";
        let dataMaxSaveback = "";
        const dateSaveback = [];
        
        // Raccogli tutte le date dei saveback
        outputTableRows.forEach(row => {
            const tipoCell = row.querySelector('td:nth-child(3)');
            if (tipoCell && tipoCell.textContent.trim() === 'Your Saveback payment') {
                const date = estraiData(row);
                if (date) {
                    dateSaveback.push(date);
                }
            }
        });
        
        // Se abbiamo trovato date, calcoliamo min e max
        if (dateSaveback.length > 0) {
            dataMinSaveback = formatDate(new Date(Math.min.apply(null, dateSaveback)));
            dataMaxSaveback = formatDate(new Date(Math.max.apply(null, dateSaveback)));
        }
        
        // MODIFICATO: Sposta il valore da ricavo a prezzo di vendita
        const prezzoVenditaFormatted = prezzoVendita ? prezzoVendita + " €" : "";
        
        risultatoFinale.push({
            DATA_INIZIO: dataMinSaveback,
            DATA_FINE: dataMaxSaveback,
            TIPO_OPERAZIONE: 'Saveback',
            Descrizione: row.Descrizione,
            Quantita: row.Quantita,
            PREZZO_DI_ACQUISTO: "0 €", // Imposta a 0 € invece che vuoto
            PREZZO_DI_VENDITA: prezzoVenditaFormatted, // Sposta qui il valore che era nel ricavo
            PREZZO_UNITARIO_ACQUISTO: '',
            PREZZO_UNITARIO_VENDITA: '',
            RICAVO: prezzoVenditaFormatted, // Mantiene il ricavo come prima (uguale al prezzo di vendita)
            RICAVO_PERCENTUALE: '' // Non calcoliamo percentuale perché il prezzo di acquisto è 0
        });
    }
    
    // Filtriamo i warrant exercise che non sono stati abbinati
    const nonAbbinati = [];
    
    for (const row of finalData.dfWarrant) {
        if (row.Descrizione && row.Descrizione.includes('ISIN')) {
            const isin = row.Descrizione.split('ISIN ')[1].trim();
            let abbinato = false;
            
            // Cerco nelle descrizioni se c'è una corrispondenza
            for (const descUnica of descrizioniUniche) {
                const descParts = descUnica.split(' ');
                if (descParts.length > 0) {
                    // Verifica se il primo elemento è un ISIN e corrisponde
                    const isinPattern = /^([A-Z]{2}\d{10}|[A-Z]{2}[0-9A-Z]{9}\d)$/;
                    if (isinPattern.test(descParts[0]) && descParts[0] === isin) {
                        abbinato = true;
                        break;
                    }
                }
            }
            
            if (!abbinato) {
                // Cerca la data del Warrant Exercise nella tabella originale
                let dataWarrant = null;
                outputTableRows.forEach(tableRow => {
                    const descCell = tableRow.querySelector('td:nth-child(4)');
                    const tipoCell = tableRow.querySelector('td:nth-child(2)');
                    const tipoOpCell = tableRow.querySelector('td:nth-child(3)');
                    
                    // Trova la riga corrispondente nella tabella delle operazioni
                    if (descCell && tipoCell && tipoOpCell && 
                        descCell.textContent.trim().includes(isin) && 
                        tipoCell.textContent.trim() === 'Redditi' && 
                        tipoOpCell.textContent.trim() === 'Warrant Exercise') {
                        
                        // Estrai e converti la data
                        const dateCell = tableRow.querySelector('td:first-child');
                        if (dateCell) {
                            const dateText = dateCell.textContent.trim();
                            const date = estraiData(tableRow);
                            if (date) {
                                dataWarrant = date;
                            }
                        }
                    }
                });
                
                // Formatta la data solo se è stata trovata
                const dataWarrantFormatted = dataWarrant ? formatDate(dataWarrant) : "";
                
                nonAbbinati.push({
                    DATA_INIZIO: "", // Non abbiamo data di acquisto
                    DATA_FINE: dataWarrantFormatted, // Data del Warrant Exercise formattata
                    TIPO_OPERAZIONE: 'Trade K.O.',
                    Descrizione: row.Descrizione,
                    Quantita: row.Quantita,
                    PREZZO_DI_ACQUISTO: row.PREZZO_DI_ACQUISTO,
                    PREZZO_DI_VENDITA: row.PREZZO_DI_VENDITA,
                    PREZZO_UNITARIO_ACQUISTO: '',
                    PREZZO_UNITARIO_VENDITA: '',
                    RICAVO: row.PREZZO_DI_VENDITA, // Imposta il ricavo uguale al prezzo di vendita
                    RICAVO_PERCENTUALE: ''
                });
            }
        }
    }
    
    // Aggiungiamo i warrant non abbinati
    risultatoFinale.push(...nonAbbinati);
    
    // Renaming dei tipi di operazione
    risultatoFinale.forEach(row => {
        if (row.TIPO_OPERAZIONE === 'Your interest payment') {
            row.TIPO_OPERAZIONE = 'Interessi Giacenza';
        } else if (row.TIPO_OPERAZIONE === 'Your Saveback payment') {
            row.TIPO_OPERAZIONE = 'Saveback';
        }
    });
    
    // Ordina per tipo di operazione
    risultatoFinale.sort((a, b) => {
        const ordine = {
            'Buy trade': 1,
            'Trade': 2,
            'Sell trade Parziale': 3,
            'Sell trade': 4,
            'Warrant Exercise': 5,
            'Interessi Giacenza': 6,
            'Saveback': 7
        };
        
        const orderA = ordine[a.TIPO_OPERAZIONE] || 99;
        const orderB = ordine[b.TIPO_OPERAZIONE] || 99;
        return orderA - orderB;
    });
    
    // Aggiungiamo simboli ai valori
    risultatoFinale.forEach(row => {
        // Aggiungiamo "azioni" alla quantità se presente
        if (row.Quantita && row.Quantita.trim() !== '') {
            if (!row.Quantita.includes(" azioni")) {
                row.Quantita = `${row.Quantita} azioni`;
            }
        }
        
        // Aggiungiamo € ai prezzi
        if (row.PREZZO_DI_ACQUISTO && row.PREZZO_DI_ACQUISTO.trim() !== '') {
            if (!row.PREZZO_DI_ACQUISTO.includes(" €")) {
                row.PREZZO_DI_ACQUISTO = `${row.PREZZO_DI_ACQUISTO} €`;
            }
        }
        
        if (row.PREZZO_DI_VENDITA && row.PREZZO_DI_VENDITA.trim() !== '') {
            if (!row.PREZZO_DI_VENDITA.includes(" €")) {
                row.PREZZO_DI_VENDITA = `${row.PREZZO_DI_VENDITA} €`;
            }
        }
        
        if (row.RICAVO && row.RICAVO.trim() !== '') {
            if (!row.RICAVO.includes(" €")) {
                row.RICAVO = `${row.RICAVO} €`;
            }
        }
        
        // Aggiungiamo €/azione ai prezzi unitari
        if (row.PREZZO_UNITARIO_ACQUISTO && row.PREZZO_UNITARIO_ACQUISTO.trim() !== '') {
            if (!row.PREZZO_UNITARIO_ACQUISTO.includes(" €/azione")) {
                row.PREZZO_UNITARIO_ACQUISTO = `${row.PREZZO_UNITARIO_ACQUISTO} €/azione`;
            }
        }
        
        if (row.PREZZO_UNITARIO_VENDITA && row.PREZZO_UNITARIO_VENDITA.trim() !== '') {
            if (!row.PREZZO_UNITARIO_VENDITA.includes(" €/azione")) {
                row.PREZZO_UNITARIO_VENDITA = `${row.PREZZO_UNITARIO_VENDITA} €/azione`;
            }
        }
        
        // Aggiungiamo % al ricavo percentuale
        if (row.RICAVO_PERCENTUALE && row.RICAVO_PERCENTUALE.trim() !== '') {
            if (!row.RICAVO_PERCENTUALE.includes(" %")) {
                row.RICAVO_PERCENTUALE = `${row.RICAVO_PERCENTUALE} %`;
            }
        }
    });
    
    return risultatoFinale;
}

// Funzione per calcolare e visualizzare i totali
function calculateAndDisplayTotals() {
    // Elementi del riepilogo
    const totalTradeCard = document.querySelector('.summary-card:nth-child(1)');
    const totalInterestCard = document.querySelector('.summary-card:nth-child(2)');
    const totalSavebackCard = document.querySelector('.summary-card:nth-child(3)');
    const totalComprehensiveCard = document.querySelector('.summary-card:nth-child(4)');
    
    const totalTrade = document.getElementById('total-trade');
    const totalInterest = document.getElementById('total-interest');
    const totalSaveback = document.getElementById('total-saveback');
    const totalComprehensive = document.getElementById('total-comprehensive');
    
    // Nuovi elementi per guadagni e perdite
    const includeSavebackToggle = document.getElementById('include-saveback-toggle');
    const totalProfits = document.getElementById('total-profits');
    const totalLosses = document.getElementById('total-losses');
    
    // Ottieni tutte le righe del riepilogo finale
    const risultatoFinale = createFinalResult();
    
    // Calcolo totale dei Trade (esclusi i Buy trade che sono quelli attualmente investiti)
    const datiTrade = risultatoFinale.filter(row => 
        row.TIPO_OPERAZIONE === 'Trade' || row.TIPO_OPERAZIONE === 'Trade K.O.'
    );
    
    // Aggiungi anche le vendite parziali ma solo per il volume, non per i ricavi
    const datiVenditeParziali = risultatoFinale.filter(row => 
        row.TIPO_OPERAZIONE === 'Sell trade Parziale'
    );
    
    // Sommiamo i valori monetari
    let totaleAcquisto = 0;
    let totaleVendita = 0;
    let totaleRicavo = 0;
    
    // Variabile per tenere traccia dei valori aggiunti ai Buy trade
    let totaleValoriAggiuntiBuyTrade = 0;
    
    // Variabili per totale guadagni e perdite
    let totaleGuadagni = 0;
    let totalePerdite = 0;
    
    // Calcolo dei totali dai dati Trade
    datiTrade.forEach(row => {
        // Processa solo i dati di commercio con valori numerici
        if (row.PREZZO_DI_ACQUISTO) {
            const prezzoAcquisto = convertNumber(row.PREZZO_DI_ACQUISTO.replace('€', '').trim());
            totaleAcquisto += prezzoAcquisto;
        }
        
        if (row.PREZZO_DI_VENDITA) {
            const prezzoVendita = convertNumber(row.PREZZO_DI_VENDITA.replace('€', '').trim());
            totaleVendita += prezzoVendita;
        }
        
        if (row.RICAVO) {
            const ricavo = convertNumber(row.RICAVO.replace('€', '').trim());
            totaleRicavo += ricavo;
            
            // Aggiungi ai totali di guadagni o perdite
            if (ricavo > 0) {
                totaleGuadagni += ricavo;
            } else if (ricavo < 0) {
                totalePerdite += ricavo;
            }
        }
    });
    
    // Aggiungi le vendite parziali solo al volume di vendita, non al ricavo
    datiVenditeParziali.forEach(row => {
        if (row.PREZZO_DI_ACQUISTO) {
            // Non aggiungiamo al totale degli acquisti perché non vogliamo contare due volte
            // i prezzi di acquisto quando si completa la vendita
        }
        
        if (row.PREZZO_DI_VENDITA) {
            const prezzoVendita = convertNumber(row.PREZZO_DI_VENDITA.replace('€', '').trim());
            totaleVendita += prezzoVendita;
        }
        
        // Non calcoliamo ricavi per le vendite parziali
    });
    
    // Controlla se ci sono valori aggiunti associati a Buy trade
    if (additionalValues && additionalValues.length > 0) {
        additionalValues.forEach(item => {
            // Estrai l'ID della riga dal formato "row-X"
            const matches = item.rowId.match(/row-(\d+)/);
            if (matches && matches.length > 1) {
                const rowIndex = parseInt(matches[1]);
                
                // Se l'indice è valido
                if (rowIndex >= 0 && rowIndex < risultatoFinale.length) {
                    const row = risultatoFinale[rowIndex];
                    
                    // Verifica se la riga era originariamente un Buy trade
                    if (row && row.TIPO_OPERAZIONE === 'Buy trade') {
                        // Aggiungi l'intero valore al totale Trade
                        totaleValoriAggiuntiBuyTrade += item.value;
                        
                        // Per il calcolo del ricavo/profitto
                        if (row.PREZZO_DI_ACQUISTO) {
                            // Estrai il valore di acquisto originale
                            const prezzoAcquistoStr = row.PREZZO_DI_ACQUISTO.replace('€', '').trim();
                            const prezzoAcquisto = convertNumber(prezzoAcquistoStr);
                            
                            // Calcola il ricavo (valore aggiunto - prezzo acquisto)
                            const ricavoCalcolato = item.value - prezzoAcquisto;
                            
                            // Aggiungi al totale guadagni o perdite in base al segno del ricavo
                            if (ricavoCalcolato > 0) {
                                totaleGuadagni += ricavoCalcolato;
                            } else if (ricavoCalcolato < 0) {
                                totalePerdite += ricavoCalcolato;
                            }
                        } else {
                            // Se non abbiamo un prezzo di acquisto, trattiamo il valore aggiunto
                            // come un ricavo (se positivo) o una perdita (se negativo)
                            if (item.value > 0) {
                                totaleGuadagni += item.value;
                            } else if (item.value < 0) {
                                totalePerdite += item.value;
                            }
                        }
                    }
                }
            }
        });
    }
    
    // Estraggo i valori di interessi e saveback
    let interessiVal = 0;
    let savebackVal = 0;
    
    // Estrai il valore degli interessi
    const rigaInteressi = risultatoFinale.find(row => row.TIPO_OPERAZIONE === 'Interessi Giacenza');
    if (rigaInteressi && rigaInteressi.RICAVO) {
        // Rimuovo il simbolo € e converto in numero
        const interessiStr = rigaInteressi.RICAVO.replace('€', '').trim();
        interessiVal = convertNumber(interessiStr);
    }
    
    // Estrai il valore del saveback
    const rigaSaveback = risultatoFinale.find(row => row.TIPO_OPERAZIONE === 'Saveback');
    if (rigaSaveback && rigaSaveback.RICAVO) {
        // Rimuovo il simbolo € e converto in numero
        const savebackStr = rigaSaveback.RICAVO.replace('€', '').trim();
        savebackVal = convertNumber(savebackStr);
    }
    
    // Calcolo il ricavo totale complessivo
    const ricavoTotaleComplessivo = totaleRicavo + interessiVal + savebackVal;
    
    // Arrotondo i valori a 3 decimali per la visualizzazione
    const totaleRicavoArrotondato = Math.round(totaleRicavo * 1000) / 1000;
    const interessiValArrotondato = Math.round(interessiVal * 1000) / 1000;
    const savebackValArrotondato = Math.round(savebackVal * 1000) / 1000;
    const ricavoTotaleComplessivoArrotondato = Math.round(ricavoTotaleComplessivo * 1000) / 1000;
    
    // Aggiorna le card del riepilogo
    // Per il totale Trade, includiamo i valori aggiunti associati ai Buy trade
    const totalTradeValue = totaleRicavo + totaleValoriAggiuntiBuyTrade;
    totalTrade.textContent = `${normalizeNumericFormat(Math.round(totalTradeValue * 1000) / 1000, true)} €`;
    totalInterest.textContent = `${normalizeNumericFormat(interessiValArrotondato, true)} €`;
    totalSaveback.textContent = `${normalizeNumericFormat(savebackValArrotondato, true)} €`;
    // Per il totale complessivo, includiamo anche i valori aggiunti associati ai Buy trade
    // Nota: ora sommiamo solo gli interessi giacenza, non il saveback che è già incluso nel Ricavo Totale Trade
    const totalComprehensiveValue = totalTradeValue + interessiVal;
    totalComprehensive.textContent = `${normalizeNumericFormat(Math.round(totalComprehensiveValue * 1000) / 1000, true)} €`;
    
    // Assegna la classe appropriata in base al segno del valore per "Ricavo Totale Trade"
    totalTrade.classList.remove('positive-value', 'negative-value');
    if (totalTradeValue > 0) {
        totalTrade.classList.add('positive-value');
    } else if (totalTradeValue < 0) {
        totalTrade.classList.add('negative-value');
    }
    
    // Assegna la classe appropriata in base al segno del valore per "Ricavo Totale Complessivo"
    totalComprehensive.classList.remove('positive-value', 'negative-value');
    if (totalComprehensiveValue > 0) {
        totalComprehensive.classList.add('positive-value');
    } else if (totalComprehensiveValue < 0) {
        totalComprehensive.classList.add('negative-value');
    }
    
    // Aggiorna i valori dei guadagni e perdite (ora sempre senza includere saveback e interessi)
    totalProfits.textContent = `${normalizeNumericFormat(Math.round(totaleGuadagni * 1000) / 1000, true)} €`;
    totalLosses.textContent = `${normalizeNumericFormat(Math.round(totalePerdite * 1000) / 1000, true)} €`;
    
    // Assegna i colori ai guadagni e alle perdite (sempre positivo per guadagni, sempre negativo per perdite)
    totalProfits.classList.add('positive-value');
    totalLosses.classList.add('negative-value');
    
    // Gestisci anche la copia del Ricavo Totale Trade
    const tradeCopy = document.querySelector('.amount.trade-copy');
    if (tradeCopy) {
        tradeCopy.textContent = totalTrade.textContent;
        
        // Sincronizza anche i colori - prima rimuovo classi esistenti
        tradeCopy.classList.remove('positive-value', 'negative-value');
        if (totalTradeValue > 0) {
            tradeCopy.classList.add('positive-value');
        } else if (totalTradeValue < 0) {
            tradeCopy.classList.add('negative-value');
        }
    }
    
    // Aggiorna il grafico
    updateTradeChart();
}

// Funzione per popolare le tabelle e inizializzare il selettore di righe
function setupTablesAndSelectors() {
    // Definisci la funzione per inizializzare e aggiornare il selettore di righe
    const updateRowSelector = setupAdditionalValuesHandling();
    
    // Gestione dei tab
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            // Rimuovi la classe active da tutti i bottoni e pannelli
            tabBtns.forEach(b => b.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));
            
            // Aggiungi la classe active al bottone corrente e al pannello corrispondente
            btn.classList.add('active');
            const target = btn.getAttribute('data-tab');
            document.getElementById(target).classList.add('active');
            
            // Se stiamo cliccando sul tab di riepilogo finale, aggiorna il selettore di righe
            if (target === 'final-csv' && typeof updateRowSelector === 'function') {
                updateRowSelector();
            }
        });
    });
    
    // Esporta la funzione di aggiornamento per poterla chiamare dopo aver elaborato il file
    window.updateRowSelector = updateRowSelector;
    
    return updateRowSelector;
}

// Modifica la funzione processExcelFile per aggiornare il selettore di righe
function processExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Prendi il primo foglio
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // Converti in JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            // Estrai le intestazioni
            const headers = jsonData[0].map(header => header ? String(header).trim() : "");
            
            // Trova gli indici delle colonne IN ENTRATA e IN USCITA
            let inEntrataIdx, inUscitaIdx;
            
            headers.forEach((header, idx) => {
                const headerUpper = header ? header.toUpperCase() : "";
                if (headerUpper.includes("IN ENTRATA")) {
                    inEntrataIdx = idx;
                } else if (headerUpper.includes("IN USCITA")) {
                    inUscitaIdx = idx;
                }
            });
            
            // Se gli indici non sono stati trovati, assumiamo che siano nelle posizioni 4 e 5
            if (inEntrataIdx === undefined || inUscitaIdx === undefined) {
                inEntrataIdx = 4;
                inUscitaIdx = 5;
            }
            
            // Raccogli i dati validi
            const validRows = [];
            
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row || row.length === 0) continue;
                
                // Controlla se il tipo è nei tipi ammessi
                const tipo = row[1];
                if (!tipo || !allowedTypes.includes(tipo)) continue;
                
                // La data è nella colonna A
                const dataStr = row[0] ? String(row[0]) : "";
                const data = convertDate(dataStr);
                
                // Aggiungi la riga ai dati da ordinare
                validRows.push([row, tipo, data]);
            }
            
            // Ordina le righe: prima per tipo, poi per data
            validRows.sort((a, b) => {
                // Ordina prima per tipo
                if (a[1] < b[1]) return -1;
                if (a[1] > b[1]) return 1;
                
                // Se i tipi sono uguali, ordina per data
                return a[2] - b[2];
            });
            
            // Processa le righe ordinate
            const rows = [];
            
            for (const [rowData, tipoVal, data] of validRows) {
                // Estraggo i valori dalle celle originali
                const dataVal = rowData[0]; // DATA (colonna A)
                
                const descrizioneOriginale = rowData[2] ? String(rowData[2]) : "";
                
                // Separo il tipo di operazione dal resto della descrizione
                const [tipoOp, restoDescrizione] = separateOperation(descrizioneOriginale);
                
                // Estraggo e formatto la quantità
                const quantita = formatQuantity(descrizioneOriginale);
                
                // Rimuovo la parte ", quantity:" dalla descrizione se presente
                let descrizioneFinale = restoDescrizione;
                if (restoDescrizione.includes(", quantity")) {
                    descrizioneFinale = restoDescrizione.split(", quantity")[0];
                }
                
                // Estraggo i valori di IN ENTRATA e IN USCITA usando gli indici corretti
                const prezzoVendita = rowData[inEntrataIdx] !== undefined ? rowData[inEntrataIdx] : "";
                const prezzoAcquisto = rowData[inUscitaIdx] !== undefined ? rowData[inUscitaIdx] : "";
                
                // Normalizzo i valori di entrata/uscita (rimuovo eventuali simboli € e formatto)
                let prezzoVenditaFormatted = "";
                let prezzoAcquistoFormatted = "";
                
                if (prezzoVendita) {
                    prezzoVenditaFormatted = String(prezzoVendita).replace('€', '').trim();
                }
                
                if (prezzoAcquisto) {
                    prezzoAcquistoFormatted = String(prezzoAcquisto).replace('€', '').trim();
                }
                
                // Calcolo del prezzo unitario
                let prezzoUnitario = "";
                try {
                    // Converto quantità in numero
                    const quantitaNum = convertNumber(quantita);
                    
                    if (quantitaNum > 0) {
                        // Determino quale importo usare (entrata o uscita)
                        if (prezzoAcquistoFormatted && String(prezzoAcquistoFormatted).trim()) {
                            const importo = convertNumber(prezzoAcquistoFormatted);
                            if (importo > 0) {
                                // Prezzo = Importo pagato/Quantità
                                prezzoUnitario = formatNumber(importo/quantitaNum);
                            }
                        } else if (prezzoVenditaFormatted && String(prezzoVenditaFormatted).trim()) {
                            const importo = convertNumber(prezzoVenditaFormatted);
                            if (importo > 0) {
                                // Se è un'entrata, il prezzo è ancora Importo/Quantità
                                prezzoUnitario = formatNumber(importo/quantitaNum);
                            }
                        }
                    }
                } catch (e) {
                    console.error("Errore nel calcolo del prezzo unitario:", e);
                }
                
                // Aggiungo la riga formattata
                rows.push({
                    DATA: dataVal,
                    TIPO: tipoVal,
                    TIPO_OPERAZIONE: tipoOp,
                    Descrizione: descrizioneFinale,
                    Quantita: quantita,
                    PREZZO_DI_VENDITA: prezzoVenditaFormatted,
                    PREZZO_DI_ACQUISTO: prezzoAcquistoFormatted,
                    PREZZO_UNITARIO: prezzoUnitario
                });
            }
            
            // Salva i dati di output
            outputData = rows;
            
            // Processa i dati per il file finale
            processFinalData(rows);
            
            // Assicuriamoci che l'overlay di caricamento sia nascosto
            const loadingOverlay = document.querySelector('.loading-overlay');
            if (loadingOverlay) {
                loadingOverlay.style.display = 'none';
            }
            
            // Mostra la sezione dei valori aggiuntivi e dei risultati
            const additionalValuesSection = document.querySelector('.additional-values-section');
            const resultsSection = document.querySelector('.results-section');
            
            if (additionalValuesSection) {
                additionalValuesSection.style.display = 'block';
            }
            
            if (resultsSection) {
                resultsSection.style.display = 'block';
            }
            
            // Popola le tabelle e i riepiloghi
            populateOutputTable(rows);
            populateFinalTable();
            
            // Calcola e mostra i totali
            calculateAndDisplayTotals();
            
            // Aggiorna il titolo del riepilogo con la data più recente
            const latestDate = getLatestDate(rows);
            const summaryTitle = document.querySelector('.summary-container h2');
            if (summaryTitle && latestDate) {
                summaryTitle.textContent = `Riepilogo Ricavi al ${latestDate}`;
            }
            
            // Aggiorna il selettore di righe per i valori aggiuntivi
            console.log("Tentativo di aggiornamento del selettore dopo il caricamento della tabella finale");
            
            // Aggiungiamo un piccolo ritardo per assicurarci che tutto sia caricato
            setTimeout(() => {
                if (window.updateRowSelector && typeof window.updateRowSelector === 'function') {
                    console.log("Aggiornamento del selettore di righe...");
                    window.updateRowSelector();
                } else {
                    console.error("updateRowSelector non disponibile o non è una funzione");
                    // Tentiamo di recuperarla
                    if (typeof setupAdditionalValuesHandling === 'function') {
                        window.updateRowSelector = setupAdditionalValuesHandling();
                        if (window.updateRowSelector) {
                            window.updateRowSelector();
                        }
                    }
                }
            }, 200);
        } catch (error) {
            console.error("Errore durante l'elaborazione del file:", error);
            alert("Si è verificato un errore durante l'elaborazione del file. Controlla la console per i dettagli.");
            
            // Assicuriamoci che l'overlay di caricamento sia nascosto anche in caso di errore
            const loadingOverlay = document.querySelector('.loading-overlay');
            if (loadingOverlay) {
                loadingOverlay.style.display = 'none';
            }
        }
    };
    
    reader.onerror = function() {
        console.error("Errore nella lettura del file");
        alert("Si è verificato un errore nella lettura del file.");
        
        // Assicuriamoci che l'overlay di caricamento sia nascosto anche in caso di errore
        const loadingOverlay = document.querySelector('.loading-overlay');
        if (loadingOverlay) {
            loadingOverlay.style.display = 'none';
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Inizializza il setup delle tabelle quando il documento è pronto
document.addEventListener('DOMContentLoaded', () => {
    // ... existing code ...
    
    // Setup delle tabelle e dei selettori di righe
    setupTablesAndSelectors();
    
    // ... rest of the code ...
});

// Funzione per applicare valori aggiuntivi a righe specifiche del riepilogo
function applyAdditionalValuesToRows(results) {
    console.log("Applicazione valori aggiuntivi al riepilogo");
    // Clone i risultati per non modificare l'originale direttamente
    const modifiedResults = JSON.parse(JSON.stringify(results));
    
    // Se non ci sono valori aggiuntivi, restituisci i risultati originali
    if (!additionalValues || additionalValues.length === 0) {
        return modifiedResults;
    }
    
    console.log(`Trovati ${additionalValues.length} valori aggiuntivi da applicare`);
    
    // Mappa per tenere traccia delle righe già modificate
    const modifiedRowsMap = {};
    
    // Non ci sono più valori globali, tutti i valori sono specifici per riga
    const rowSpecificValues = additionalValues;
    
    console.log(`Valori specifici: ${rowSpecificValues.length}`);
    
    // Ottieni la data di oggi in formato "12 apr 24"
    const today = new Date();
    const day = today.getDate();
    const monthNames = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
    const month = monthNames[today.getMonth()];
    const year = today.getFullYear().toString().substr(-2);
    const todayFormatted = `${day} ${month} ${year}`;
    
    // Applica i valori specifici per riga
    rowSpecificValues.forEach(item => {
        // Estrai l'ID della riga dal formato "row-X"
        const matches = item.rowId.match(/row-(\d+)/);
        if (matches && matches.length > 1) {
            const rowIndex = parseInt(matches[1]);
            
            // Verifica che la riga esista nei risultati
            if (rowIndex >= 0 && rowIndex < modifiedResults.length) {
                const row = modifiedResults[rowIndex];
                const value = item.value;
                
                console.log(`Applicazione valore ${value} alla riga ${rowIndex} (${row.TIPO_OPERAZIONE})`);
                
                // Aggiorna la DATA_FINE (DATA ULTIMA VENDITA) con la data odierna
                row.DATA_FINE = todayFormatted;
                
                // In base al tipo di operazione, trattiamo il valore in modo diverso
                if (row.TIPO_OPERAZIONE === 'Buy trade') {
                    // Per i Buy trade, aggiungiamo il valore come prezzo di vendita
                    // e calcoliamo ricavo e percentuale
                    
                    // Estrai la quantità e il prezzo di acquisto
                    const quantita = convertNumber(row.Quantita.replace('azioni', '').trim());
                    const prezzoAcquisto = convertNumber(row.PREZZO_DI_ACQUISTO.replace('€', '').trim());
                    
                    // Imposta il prezzo di vendita
                    row.PREZZO_DI_VENDITA = `${normalizeNumericFormat(value, true)} €`;
                    
                    // Calcola il prezzo unitario di vendita se la quantità è valida
                    if (quantita > 0) {
                        const prezzoUnitarioVendita = value / quantita;
                        row.PREZZO_UNITARIO_VENDITA = `${normalizeNumericFormat(prezzoUnitarioVendita, true)} €/azione`;
                    } else {
                        row.PREZZO_UNITARIO_VENDITA = '';
                    }
                    
                    // Calcola il ricavo come valore aggiunto meno il prezzo di acquisto
                    const ricavo = value - prezzoAcquisto; // Modifica qui
                    row.RICAVO = `${normalizeNumericFormat(ricavo, true)} €`;
                    
                    // Calcola il ricavo percentuale
                    if (prezzoAcquisto > 0) {
                        const ricavoPercentuale = (ricavo / prezzoAcquisto) * 100;
                        row.RICAVO_PERCENTUALE = `${normalizeNumericFormat(ricavoPercentuale, true)} %`;
                    } else {
                        row.RICAVO_PERCENTUALE = '';
                    }
                    
                    // Cambia il tipo di operazione a 'Trade' invece di 'Buy trade'
                    row.TIPO_OPERAZIONE = 'Trade';
                    
                    console.log(`Riga ${rowIndex} convertita da Buy trade a Trade con ricavo ${ricavo}`);
                    
                } else if (row.TIPO_OPERAZIONE === 'Sell trade Parziale') {
                    // MODIFICATO: Quando si aggiunge un valore a una vendita parziale, la convertiamo in Trade completato
                    
                    // Estrai i valori esistenti
                    // Per il nuovo formato "X azioni comprate\nY azioni vendute", estraiamo solo la parte delle azioni vendute
                    let quantitaVenduta = 0;
                    if (row.Quantita && row.Quantita.includes('azioni vendute')) {
                        const match = row.Quantita.match(/(\d+[,.]?\d*)\s+azioni vendute/);
                        if (match) {
                            const qtyStr = match[1].replace(',', '.');
                            quantitaVenduta = parseFloat(qtyStr);
                        }
                    } else {
                        // Fallback al vecchio formato
                        quantitaVenduta = convertNumber(row.Quantita.replace('azioni', '').trim());
                    }
                    
                    // Estrai anche la quantità totale acquistata
                    let quantitaTotaleAcquistata = 0;
                    if (row.Quantita && row.Quantita.includes('azioni comprate')) {
                        const match = row.Quantita.match(/(\d+[,.]?\d*)\s+azioni comprate/);
                        if (match) {
                            const qtyStr = match[1].replace(',', '.');
                            quantitaTotaleAcquistata = parseFloat(qtyStr);
                        }
                    }
                    
                    // Calcola quantità residua (quella che stiamo vendendo ora)
                    const quantitaResidua = quantitaTotaleAcquistata - quantitaVenduta;
                    
                    const prezzoAcquisto = convertNumber(row.PREZZO_DI_ACQUISTO.replace('€', '').trim());
                    const vecchioPrezzoVendita = convertNumber(row.PREZZO_DI_VENDITA.replace('€', '').trim());
                    
                    // Estrai il vecchio prezzo unitario di vendita
                    let vecchioPrezzoUnitarioVendita = 0;
                    if (row.PREZZO_UNITARIO_VENDITA) {
                        vecchioPrezzoUnitarioVendita = convertNumber(row.PREZZO_UNITARIO_VENDITA.replace('€/azione', '').trim());
                    }
                    
                    // Calcola il prezzo unitario del nuovo valore aggiunto (valore aggiunto diviso quantità residua)
                    const nuovoPrezzoUnitario = quantitaResidua > 0 ? value / quantitaResidua : 0;
                    
                    // Calcola il nuovo prezzo di vendita totale
                    const nuovoPrezzoVendita = vecchioPrezzoVendita + value;
                    row.PREZZO_DI_VENDITA = `${normalizeNumericFormat(nuovoPrezzoVendita, true)} €`;
                    
                    // Calcola la media ponderata per il prezzo unitario di vendita
                    if (quantitaVenduta > 0 && quantitaResidua > 0) {
                        // Media ponderata: (prezzo1 * quantità1 + prezzo2 * quantità2) / (quantità1 + quantità2)
                        const prezzoUnitarioVendita = (vecchioPrezzoUnitarioVendita * quantitaVenduta + nuovoPrezzoUnitario * quantitaResidua) / (quantitaVenduta + quantitaResidua);
                        row.PREZZO_UNITARIO_VENDITA = `${normalizeNumericFormat(prezzoUnitarioVendita, true)} €/azione`;
                    } else {
                        row.PREZZO_UNITARIO_VENDITA = '';
                    }
                    
                    // Ora che l'operazione è completata, calcoliamo il ricavo
                    const ricavo = nuovoPrezzoVendita - prezzoAcquisto;
                    row.RICAVO = `${normalizeNumericFormat(ricavo, true)} €`;
                    
                    // Calcoliamo il ricavo percentuale
                    if (prezzoAcquisto > 0) {
                        const ricavoPercentuale = (ricavo / prezzoAcquisto) * 100;
                        row.RICAVO_PERCENTUALE = `${normalizeNumericFormat(ricavoPercentuale, true)} %`;
                    } else {
                        row.RICAVO_PERCENTUALE = '';
                    }
                    
                    // Quando si completa la vendita, aggiorniamo la quantità mostrata
                    // Mostra la quantità totale (venduta + residua)
                    row.Quantita = `${normalizeNumericFormat(quantitaVenduta + quantitaResidua, true)} azioni`;
                    
                    // Cambia il tipo di operazione a 'Trade' invece di 'Sell trade Parziale'
                    row.TIPO_OPERAZIONE = 'Trade';
                    
                    console.log(`Riga ${rowIndex} convertita da Sell trade Parziale a Trade con ricavo ${ricavo}`);
                    
                } else if (row.TIPO_OPERAZIONE === 'Trade') {
                    // Per i Trade esistenti, aggiorniamo il prezzo di vendita e ricalcoliamo
                    
                    // Estrai i valori esistenti
                    const quantitaTotale = convertNumber(row.Quantita.replace('azioni', '').trim());
                    const prezzoAcquisto = convertNumber(row.PREZZO_DI_ACQUISTO.replace('€', '').trim());
                    const vecchioPrezzoVendita = convertNumber(row.PREZZO_DI_VENDITA.replace('€', '').trim());
                    
                    // Estrai il vecchio prezzo unitario di vendita
                    let vecchioPrezzoUnitarioVendita = 0;
                    if (row.PREZZO_UNITARIO_VENDITA) {
                        vecchioPrezzoUnitarioVendita = convertNumber(row.PREZZO_UNITARIO_VENDITA.replace('€/azione', '').trim());
                    }
                    
                    // Stimiamo la quantità per la vendita precedente usando il prezzo di vendita e il prezzo unitario
                    let quantitaVenditaPrecedente = 0;
                    if (vecchioPrezzoUnitarioVendita > 0) {
                        quantitaVenditaPrecedente = vecchioPrezzoVendita / vecchioPrezzoUnitarioVendita;
                    }
                    
                    // Calcola il prezzo unitario del nuovo valore aggiunto
                    // Assumiamo che il nuovo valore sia per le azioni rimanenti (quantitaTotale - quantitaVenditaPrecedente)
                    const quantitaResidua = Math.max(0, quantitaTotale - quantitaVenditaPrecedente);
                    const nuovoPrezzoUnitario = quantitaResidua > 0 ? value / quantitaResidua : 0;
                    
                    // Calcola il nuovo prezzo di vendita totale
                    const nuovoPrezzoVendita = vecchioPrezzoVendita + value;
                    row.PREZZO_DI_VENDITA = `${normalizeNumericFormat(nuovoPrezzoVendita, true)} €`;
                    
                    // Calcola la media ponderata per il prezzo unitario di vendita
                    if (quantitaTotale > 0) {
                        let prezzoUnitarioVendita;
                        if (quantitaVenditaPrecedente > 0 && quantitaResidua > 0) {
                            // Media ponderata: (prezzo1 * quantità1 + prezzo2 * quantità2) / (quantità1 + quantità2)
                            prezzoUnitarioVendita = (vecchioPrezzoUnitarioVendita * quantitaVenditaPrecedente + nuovoPrezzoUnitario * quantitaResidua) / quantitaTotale;
                        } else {
                            // Se non abbiamo dati precedenti, semplicemente dividiamo il prezzo totale per la quantità
                            prezzoUnitarioVendita = nuovoPrezzoVendita / quantitaTotale;
                        }
                        row.PREZZO_UNITARIO_VENDITA = `${normalizeNumericFormat(prezzoUnitarioVendita, true)} €/azione`;
                    } else {
                        row.PREZZO_UNITARIO_VENDITA = '/';
                        row.Quantita = '/';
                    }
                    
                    // Ricalcola il ricavo
                    const nuovoRicavo = nuovoPrezzoVendita - prezzoAcquisto;
                    row.RICAVO = `${normalizeNumericFormat(nuovoRicavo, true)} €`;
                    
                    // Ricalcola il ricavo percentuale
                    if (prezzoAcquisto > 0) {
                        const ricavoPercentuale = (nuovoRicavo / prezzoAcquisto) * 100;
                        row.RICAVO_PERCENTUALE = `${normalizeNumericFormat(ricavoPercentuale, true)} %`;
                    } else {
                        row.RICAVO_PERCENTUALE = '';
                    }
                    
                    console.log(`Riga ${rowIndex} (Trade) aggiornata: nuovo prezzo vendita ${nuovoPrezzoVendita}, ricavo ${nuovoRicavo}`);
                    
                } else if (row.TIPO_OPERAZIONE === 'Interessi Giacenza' || row.TIPO_OPERAZIONE === 'Saveback') {
                    // Per Interessi e Saveback, aggiungiamo il valore direttamente al RICAVO e non al prezzo di vendita
                    
                    // Estrai il valore del ricavo attuale
                    const vecchioRicavo = convertNumber(row.RICAVO.replace('€', '').trim());
                    const nuovoRicavo = vecchioRicavo + value;
                    
                    // Aggiorna il ricavo con il nuovo valore
                    row.RICAVO = `${normalizeNumericFormat(nuovoRicavo, true)} €`;
                    
                    // Copiamo il valore del ricavo anche nella colonna PREZZO_DI_VENDITA
                    row.PREZZO_DI_VENDITA = row.RICAVO;
                    
                    // Aggiorna sempre la DATA_FINE con la data odierna
                    row.DATA_FINE = todayFormatted;
                    
                    console.log(`Riga ${rowIndex} (${row.TIPO_OPERAZIONE}) aggiornata: nuovo ricavo ${nuovoRicavo}`);
                }
                
                // Segna questa riga come modificata
                modifiedRowsMap[rowIndex] = true;
            }
        }
    });
    
    return modifiedResults;
}

// Modifica la funzione createFinalResult per applicare i valori aggiuntivi
const originalCreateFinalResult = window.createFinalResult;
window.createFinalResult = function() {
    // Ottieni il risultato originale
    const originalResults = originalCreateFinalResult();
    
    // Applica i valori aggiuntivi specifici per riga
    const modifiedResults = applyAdditionalValuesToRows(originalResults);
    
    return modifiedResults;
}; 

// Funzione per cercare nella tabella
function searchTable(table, searchTerm) {
    if (!table) return;
    
    const rows = table.querySelectorAll('tbody tr');
    
    // Se il termine di ricerca è vuoto, mostra tutte le righe e rimuovi l'evidenziazione
    if (!searchTerm || searchTerm.trim() === '') {
        rows.forEach(row => {
            row.style.display = '';
            row.querySelectorAll('td').forEach(cell => {
                cell.classList.remove('highlight');
            });
        });
        return;
    }
    
    // Altrimenti, filtra in base al termine di ricerca
    rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        const containsSearchTerm = text.includes(searchTerm);
        
        // Mostra o nascondi la riga
        row.style.display = containsSearchTerm ? '' : 'none';
        
        if (containsSearchTerm) {
            // Evidenzia il testo corrispondente nelle celle
            const cells = row.querySelectorAll('td');
            let foundMatch = false;
            
            cells.forEach(cell => {
                const cellText = cell.textContent.toLowerCase();
                if (cellText.includes(searchTerm)) {
                    cell.classList.add('highlight');
                    foundMatch = true;
                } else {
                    cell.classList.remove('highlight');
                }
            });
            
            // Se abbiamo trovato corrispondenze, scorri alla prima riga visibile
            if (foundMatch && !window.hasScrolledToMatch) {
                row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                window.hasScrolledToMatch = true;
                
                // Resetta il flag dopo un breve ritardo
                setTimeout(() => {
                    window.hasScrolledToMatch = false;
                }, 1000);
            }
        }
    });
}

// Funzione per scaricare la tabella come CSV
function downloadCSV(table, filename) {
    if (!table) return;
    
    // Ottieni tutte le righe della tabella
    const rows = table.querySelectorAll('tr');
    if (!rows || rows.length === 0) return;
    
    // Preparare l'output CSV
    let csvContent = '';
    const separator = ';'; // Separatore per il CSV (punto e virgola per Excel italiano)
    
    // Processa ogni riga
    rows.forEach(row => {
        const cells = row.querySelectorAll('th, td');
        if (!cells || cells.length === 0) return;
        
        const rowData = [];
        
        // Processa ogni cella
        cells.forEach(cell => {
            // Ottieni il testo e rimuovi spazi extra
            let text = cell.textContent.trim();
            
            // Gestisci caratteri speciali e virgolette
            if (text.includes(separator) || text.includes('"') || text.includes('\n')) {
                // Sostituisci le virgolette doppie con due virgolette doppie
                text = text.replace(/"/g, '""');
                // Aggiungi virgolette doppie intorno al testo
                text = `"${text}"`;
            }
            
            rowData.push(text);
        });
        
        // Aggiungi la riga al CSV
        csvContent += rowData.join(separator) + '\n';
    });
    
    // Crea un oggetto Blob per il download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    
    // Crea un link per il download
    const link = document.createElement('a');
    
    // Supporto per IE
    if (navigator.msSaveBlob) {
        navigator.msSaveBlob(blob, filename);
        return;
    }
    
    // Crea un URL per il blob
    const url = URL.createObjectURL(blob);
    
    // Imposta gli attributi per il download
    link.href = url;
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    
    // Aggiungi il link al documento
    document.body.appendChild(link);
    
    // Clicca sul link per avviare il download
    link.click();
    
    // Pulizia
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Gestione del file selezionato
function handleFiles(files) {
    if (files.length === 0) return;
    
    const file = files[0];
    const fileExtension = file.name.split('.').pop().toLowerCase();
    
    if (fileExtension !== 'xlsx' && fileExtension !== 'xls') {
        alert('Per favore, seleziona un file Excel (.xlsx o .xls)');
        return;
    }
    
    // Otteniamo i riferimenti agli elementi DOM necessari
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const fileInfo = document.getElementById('file-info');
    const uploadProgress = document.querySelector('.upload-progress');
    const progressBar = document.querySelector('.progress-bar-fill');
    const progressText = document.querySelector('.progress-text');
    const loadingOverlay = document.querySelector('.loading-overlay');
    const additionalValuesSection = document.querySelector('.additional-values-section');
    const resultsSection = document.querySelector('.results-section');
    const fileInput = document.getElementById('file-input');
    
    // Mostra la barra di progresso
    uploadProgress.style.display = 'block';
    
    // Nascondi il pulsante Elabora File durante il caricamento
    const processBtn = document.getElementById('process-btn');
    if (processBtn) {
        processBtn.style.display = 'none';
    }
    
    // Mostra le informazioni sul file ma nascondi il pulsante elabora
    fileInfo.style.display = 'block';
    
    // Visualizza le informazioni sul file
    fileName.textContent = file.name;
    
    // Formatta dimensione file
    const size = file.size;
    let formattedSize;
    
    if (size < 1024) {
        formattedSize = `${size} bytes`;
    } else if (size < 1024 * 1024) {
        formattedSize = `${(size / 1024).toFixed(2)} KB`;
    } else {
        formattedSize = `${(size / (1024 * 1024)).toFixed(2)} MB`;
    }
    
    fileSize.textContent = formattedSize;
    
    // Salva il file per l'elaborazione successiva
    window.selectedFile = file;
    
    // Resetta il valore dell'input file per permettere di ricaricare lo stesso file
    fileInput.value = '';
    
    // Simula un caricamento progressivo
    let progress = 0;
    const interval = setInterval(() => {
        progress += 5;
        progressBar.style.width = `${progress}%`;
        progressText.textContent = `${progress}%`;
        
        if (progress >= 100) {
            clearInterval(interval);
            
            // Nascondi la barra di progresso dopo un breve ritardo
            setTimeout(() => {
                uploadProgress.style.display = 'none';
                
                // Reset dei valori aggiuntivi quando si elabora un nuovo file
                additionalValues = [];
                
                // Elabora il file Excel direttamente senza mostrare il loading overlay
                processExcelFile(window.selectedFile);
            }, 800);
        }
    }, 50);
}

// Aggiungo la funzione handleHeaderClick all'inizio del file, sotto il listener DOMContentLoaded

// Funzione per gestire il click su un'intestazione di tabella per ordinamento
function handleHeaderClick(headers, table) {
    headers.forEach(header => {
        header.addEventListener('click', () => {
            const sortKey = header.getAttribute('data-sort');
            const isDateColumn = sortKey === 'data' || sortKey === 'data-inizio' || sortKey === 'data-fine';
            
            // Determina la direzione di ordinamento
            let ascending = true;
            if (header.classList.contains('sort-asc')) {
                ascending = false;
                header.classList.remove('sort-asc');
                header.classList.add('sort-desc');
            } else if (header.classList.contains('sort-desc')) {
                ascending = true;
                header.classList.remove('sort-desc');
                header.classList.add('sort-asc');
            } else {
                // Rimuovi le classi da tutti gli altri header
                headers.forEach(h => {
                    h.classList.remove('sort-asc', 'sort-desc');
                });
                // Aggiungi la classe al header corrente
                
                // Per le colonne di date, impostiamo l'ordinamento decrescente come default 
                // (dalla data più recente alla più lontana)
                if (isDateColumn) {
                    ascending = false;
                    header.classList.add('sort-desc');
                } else {
                    header.classList.add('sort-asc');
                }
            }
            
            // Esegui l'ordinamento
            sortTable(table, sortKey, ascending);
        });
    });
}

// Funzione per aggiungere un valore aggiuntivo alla tabella dettaglio operazioni
function addValueToOutputTable(value, description) {
    const outputTable = document.getElementById('output-table');
    if (!outputTable) return;
    
    const tbody = outputTable.querySelector('tbody');
    if (!tbody) return;
    
    // Crea una nuova riga
    const newRow = document.createElement('tr');
    newRow.className = 'added-value-row'; // Aggiungi una classe per identificare le righe aggiunte
    
    // Data corrente
    const today = new Date();
    const formattedDate = formatDate(today);
    
    // Estrai il tipo di operazione dalla descrizione (formato: "TIPO - Descrizione")
    const descriptionParts = description.split(' - ');
    const tipoOperazione = descriptionParts.length > 0 ? descriptionParts[0] : '';
    const soloDescrizione = descriptionParts.length > 1 ? descriptionParts.slice(1).join(' - ') : '';
    
    // Determina il tipo corretto in base al tipo di operazione
    let tipo = 'Commercio'; // Valore predefinito per la maggior parte delle operazioni
    let tipoOperazioneCorrect = 'Sell trade'; // Valore predefinito per le operazioni commerciali
    let quantita = '';
    let prezzoUnitario = '';
    let prezzoAcquisto = '';
    let prezzoVendita = `${normalizeNumericFormat(value, true)} €`; // Mostra sempre il valore nella colonna prezzo di vendita
    let remainingQty = 0;
    
    if (tipoOperazione === 'Interessi Giacenza') {
        tipo = 'Pagamento degli interessi';
        tipoOperazioneCorrect = 'Your interest payment';
        // Per gli interessi, mostriamo solo il prezzo di vendita senza calcolare quantità/prezzo unitario
    } else if (tipoOperazione === 'Saveback') {
        tipo = 'Premio';
        tipoOperazioneCorrect = 'Your Saveback payment';
        // Per i saveback, mostriamo solo il prezzo di vendita senza calcolare quantità/prezzo unitario
    } else {
        // Per tutti gli altri casi, inclusi i Buy trade/Sell trade, determiniamo le quantità
        
        // 1. Calcola la quantità totale acquistata e venduta dai dettagli operazioni
        let totalBuyQty = 0;
        let totalSellQty = 0;
        
        // Esamina tutte le righe della tabella di dettaglio operazioni per questa descrizione
        const outputTableRows = document.querySelectorAll('#output-table tbody tr');
        outputTableRows.forEach(row => {
            const rowDesc = row.cells[3].textContent.trim();
            const rowTipoOperazione = row.cells[2].textContent.trim();
            
            if (rowDesc === soloDescrizione) {
                // Estrai la quantità se presente
                const qtyCell = row.cells[4].textContent.trim();
                // Modifica l'espressione regolare per catturare numeri con virgole o punti
                const qtyMatch = qtyCell.match(/(\d+[,.]?\d*)/);
                
                if (qtyMatch) {
                    // Converti il valore in numero, sostituendo la virgola con punto
                    const qtyStr = qtyMatch[1].replace(',', '.');
                    const qty = parseFloat(qtyStr);
                    
                    // Consideriamo Buy trade e Savings plan execution come operazioni di acquisto
                    if (rowTipoOperazione.toLowerCase().includes('buy') || 
                        rowTipoOperazione.toLowerCase().includes('savings plan')) {
                        totalBuyQty += qty;
                        console.log(`Acquisto: ${rowTipoOperazione}, Quantità: ${qty}, Totale: ${totalBuyQty}`);
                    } else if (rowTipoOperazione.toLowerCase().includes('sell')) {
                        totalSellQty += qty;
                        console.log(`Vendita: ${rowTipoOperazione}, Quantità: ${qty}, Totale: ${totalSellQty}`);
                    }
                }
            }
        });
        
        // Calcola la quantità rimanente (acquistata - venduta)
        remainingQty = totalBuyQty - totalSellQty;
        
        console.log(`Descrizione: ${soloDescrizione}, Acquistate totali: ${totalBuyQty}, Vendute totali: ${totalSellQty}, Rimanenti: ${remainingQty}`);
        
        // 2. Se la quantità è stata calcolata correttamente, usa quel valore
        if (remainingQty > 0) {
            // Formatta il numero con 6 decimali per mantenere la precisione delle frazioni
            quantita = `${remainingQty.toFixed(6).replace(/\.?0+$/, '')} azioni`;
            
            // Per i Sell trade, non mostriamo il prezzo di acquisto nel dettaglio operazioni
            prezzoAcquisto = '';
        }
        // 3. Caso specifico per DE000VK1A0H3 o altri valori predefiniti
        else if (soloDescrizione.includes('DE000VK1A0H3')) {
            // Caso specifico per il Turbo auf APPLE INC.
            remainingQty = 199;
            quantita = `${remainingQty} azioni`;
            prezzoAcquisto = ''; // Rimuoviamo il prezzo di acquisto
        }
        // 4. In caso di fallimento totale, prova a trovare l'ultima operazione di acquisto
        else {
            // Cerca l'ultima operazione di acquisto per questa descrizione
            let lastBuyRow = null;
            
            // Esamina le righe in ordine inverso per trovare l'ultimo acquisto
            Array.from(outputTableRows).reverse().forEach(row => {
                if (!lastBuyRow) {
                    const rowDesc = row.cells[3].textContent.trim();
                    const rowTipoOperazione = row.cells[2].textContent.trim();
                    
                    if (rowDesc === soloDescrizione && rowTipoOperazione.toLowerCase().includes('buy')) {
                        lastBuyRow = row;
                    }
                }
            });
            
            // Se trovato, estrai la quantità (ma non il prezzo)
            if (lastBuyRow) {
                const qtyCell = lastBuyRow.cells[4].textContent.trim();
                const qtyMatch = qtyCell.match(/(\d+[,.]?\d*)/);
                if (qtyMatch) {
                    const qtyStr = qtyMatch[1].replace(',', '.');
                    remainingQty = parseFloat(qtyStr);
                } else {
                    remainingQty = 1;
                }
            } else {
                // Fallback a valori di default
                remainingQty = 1;
            }
            
            // Formatta il numero con 6 decimali per mantenere la precisione delle frazioni
            quantita = `${remainingQty.toFixed(6).replace(/\.?0+$/, '')} azioni`;
            prezzoAcquisto = ''; // Rimuoviamo il prezzo di acquisto
        }
        
        // Calcola il prezzo unitario (se abbiamo una quantità)
        if (remainingQty > 0) {
            const unitPrice = value / remainingQty;
            prezzoUnitario = `${normalizeNumericFormat(unitPrice, true)} €/azione`;
        }
    }
    
    // Crea le celle
    const cells = [
        formattedDate, // DATA
        tipo, // TIPO
        tipoOperazioneCorrect, // TIPO OPERAZIONE
        soloDescrizione, // DESCRIZIONE
        quantita, // QUANTITÀ
        prezzoAcquisto, // PREZZO DI ACQUISTO (ora vuoto per Sell trade)
        prezzoVendita, // PREZZO DI VENDITA
        prezzoUnitario // PREZZO UNITARIO
    ];
    
    // Popola la riga con le celle
    cells.forEach(cellText => {
        const cell = document.createElement('td');
        cell.textContent = cellText;
        newRow.appendChild(cell);
    });
    
    // Aggiungi la riga alla tabella
    tbody.appendChild(newRow);
    
    // Ordina la tabella per data (dalla più recente alla più vecchia)
    sortTable(outputTable, 'data', false); // false = ordine decrescente
    
    // Ricalcola i totali dopo aver aggiunto il nuovo valore
    calculateAndDisplayTotals();
    
    // Funzione interna per calcolare la quantità totale dal riepilogo per una descrizione/tipo specifico
    function calcTotalQuantityInFinalSummary(description, operationType) {
        let totalQty = 0;
        const finalTable = document.getElementById('final-table');
        
        if (!finalTable) return totalQty;
        
        const rows = finalTable.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            if (cells.length >= 5) {
                const rowDesc = cells[3].textContent.trim();
                const rowType = cells[2].textContent.trim();
                
                if (rowDesc === description && rowType === operationType) {
                    const qtyText = cells[4].textContent.trim();
                    const qtyMatch = qtyText.match(/(\d+[,.]?\d*)/);
                    if (qtyMatch && qtyMatch.length > 1) {
                        const qtyStr = qtyMatch[1].replace(',', '.');
                        totalQty += parseFloat(qtyStr);
                    }
                }
            }
        });
        
        return totalQty;
    }
    
    // Formatta la data corrente per il formato "12 apr 24"
    function formatDate(date) {
        const day = date.getDate();
        const monthNames = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear().toString().slice(-2);
        return `${day} ${month} ${year}`;
    }
}

// Funzione per creare o aggiornare il grafico
function updateTradeChart() {
    // Ottieni i dati dal riepilogo finale per le operazioni di tipo trade
    const risultatoFinale = createFinalResult();
    
    // Filtra le operazioni di tipo 'Trade', 'Trade K.O.' e 'Warrant Exercise'
    const datiTrade = risultatoFinale.filter(row => 
        row.TIPO_OPERAZIONE === 'Trade' || row.TIPO_OPERAZIONE === 'Trade K.O.' || row.TIPO_OPERAZIONE === 'Warrant Exercise'
    );
    
    // Strutture dati per guadagni, perdite e profitti
    let datiPerPeriodo = {};
    
    // Popola gli anni disponibili
    availableYears = [];
    
    // Plugin personalizzato per disegnare linee separatrici verticali tra i periodi
    const separatorPlugin = {
        id: 'periodSeparator',
        beforeDraw(chart) {
            const { ctx, chartArea, scales } = chart;
            const { left, right, top, bottom } = chartArea;
            
            if (chartPeriod === 'monthly') {
                // Memorizza i nomi dei mesi per rilevare i cambi
                const mesi = [];
                chart.data.labels.forEach(label => {
                    // Estrae il nome del mese dalla label (es. "Gen 2024" -> "Gen")
                    const mese = label.split(' ')[0];
                    mesi.push(mese);
                });
                
                // Disegna linee verticali al cambio del mese
                ctx.save();
                ctx.beginPath();
                ctx.lineWidth = 2;
                ctx.strokeStyle = 'rgba(0, 0, 0, 0.2)';
                ctx.setLineDash([]);
                
                for (let i = 1; i < mesi.length; i++) {
                    if (mesi[i] !== mesi[i - 1]) {
                        // Calcola la posizione x tra le barre
                        const xPos = (scales.x.getPixelForValue(i-1) + scales.x.getPixelForValue(i)) / 2;
                        
                        // Disegna la linea
                        ctx.moveTo(xPos, top);
                        ctx.lineTo(xPos, bottom);
                    }
                }
                
                ctx.stroke();
                ctx.restore();
            } else if (chartPeriod === 'weekly') {
                // Memorizza i numeri delle settimane per rilevare i cambi di mese
                const settimane = [];
                chart.data.labels.forEach(label => {
                    // Estrae la settimana dalla label (es. "Sett 1" -> estrae 1)
                    const settimana = parseInt(label.split(' ')[1]);
                    settimane.push(settimana);
                });
                
                // Disegna linee verticali per ogni settimana
                ctx.save();
                ctx.beginPath();
                ctx.lineWidth = 1; // Linea più sottile per non appesantire troppo la visualizzazione
                ctx.strokeStyle = 'rgba(0, 0, 0, 0.15)'; // Più trasparente per non dominare visivamente
                ctx.setLineDash([]);
                
                for (let i = 1; i < settimane.length; i++) {
                    // Calcola la posizione x tra le barre per ogni settimana
                    const xPos = (scales.x.getPixelForValue(i-1) + scales.x.getPixelForValue(i)) / 2;
                    
                    // Disegna la linea
                    ctx.moveTo(xPos, top);
                    ctx.lineTo(xPos, bottom);
                    
                    // Se è l'inizio di un nuovo mese (settimana 1 dopo altre settimane), falla più evidente
                    if (settimane[i] === 1 && settimane[i-1] > 1) {
                        // Prima chiudi il percorso corrente con linee sottili
                        ctx.stroke();
                        
                        // Inizia un nuovo percorso per la linea più spessa
                        ctx.beginPath();
                        ctx.lineWidth = 2;
                        ctx.strokeStyle = 'rgba(0, 0, 0, 0.3)';
                        
                        // Ridisegna la stessa linea ma più spessa
                        ctx.moveTo(xPos, top);
                        ctx.lineTo(xPos, bottom);
                        
                        // Disegna questa linea più spessa
                        ctx.stroke();
                        
                        // Resetta per continuare con le linee normali
                        ctx.beginPath();
                        ctx.lineWidth = 1;
                        ctx.strokeStyle = 'rgba(0, 0, 0, 0.15)';
                    }
                }
                
                ctx.stroke();
                ctx.restore();
            } else if (chartPeriod === 'yearly') {
                // Per la visualizzazione annuale, dobbiamo identificare i gruppi di anni
                const anni = [];
                chart.data.labels.forEach(label => {
                    // Estrae l'anno dalla label (es. "2023")
                    const anno = parseInt(label);
                    anni.push(anno);
                });
                
                // Disegna linee verticali tra gli anni
                ctx.save();
                ctx.beginPath();
                ctx.lineWidth = 2;
                ctx.strokeStyle = 'rgba(0, 0, 0, 0.2)';
                ctx.setLineDash([]);
                
                for (let i = 1; i < anni.length; i++) {
                    // Calcola la posizione x tra le barre per ogni anno
                    const xPos = (scales.x.getPixelForValue(i-1) + scales.x.getPixelForValue(i)) / 2;
                    
                    // Disegna la linea
                    ctx.moveTo(xPos, top);
                    ctx.lineTo(xPos, bottom);
                    
                    // Se passiamo anche da un decennio all'altro (es. da 2019 a 2020)
                    if (Math.floor(anni[i] / 10) !== Math.floor(anni[i-1] / 10)) {
                        // Prima chiudi il percorso corrente
                        ctx.stroke();
                        
                        // Inizia un nuovo percorso per la linea più spessa
                        ctx.beginPath();
                        ctx.lineWidth = 3;
                        ctx.strokeStyle = 'rgba(0, 0, 0, 0.3)';
                        
                        // Ridisegna la stessa linea ma più spessa
                        ctx.moveTo(xPos, top);
                        ctx.lineTo(xPos, bottom);
                        
                        // Disegna questa linea più spessa
                        ctx.stroke();
                        
                        // Resetta per continuare con le linee normali
                        ctx.beginPath();
                        ctx.lineWidth = 2;
                        ctx.strokeStyle = 'rgba(0, 0, 0, 0.2)';
                    }
                }
                
                ctx.stroke();
                ctx.restore();
            }
        }
    };
    
    // Plugin per disegnare una linea orizzontale a 0 euro
    const zeroLinePlugin = {
        id: 'zeroLine',
        // Cambiato da afterDraw a beforeDraw per renderlo sottostante al tooltip
        beforeDraw(chart) {
            const { ctx, chartArea, scales } = chart;
            const { left, right, top, bottom } = chartArea;
            
            if (!scales.y) return;
            
            // Ottieni la posizione Y del valore zero
            const zeroY = scales.y.getPixelForValue(0);
            
            // Disegna la linea orizzontale solo se zero è visibile nel grafico
            if (zeroY >= top && zeroY <= bottom) {
                ctx.save();
                ctx.beginPath();
                ctx.moveTo(left, zeroY);
                ctx.lineTo(right, zeroY);
                ctx.lineWidth = 2;
                ctx.strokeStyle = 'rgba(0, 0, 0, 0.3)';
                ctx.setLineDash([]);
                ctx.stroke();
                ctx.restore();
            }
        }
    };
    
    // Analizza le date di ultima vendita (DATA_FINE)
    datiTrade.forEach(row => {
        if (!row.DATA_FINE || !row.RICAVO) return;
        
        // Estrai la data
        const dateParts = row.DATA_FINE.split(' ');
        if (dateParts.length !== 3) return;
        
        const day = parseInt(dateParts[0]);
        const monthNames = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
        const monthIndex = monthNames.map(m => m.toLowerCase()).indexOf(dateParts[1].toLowerCase());
        if (monthIndex === -1) return;
        
        const year = parseInt('20' + dateParts[2]);
        if (isNaN(year)) return;
        
        // Aggiungi l'anno alla lista degli anni disponibili
        if (!availableYears.includes(year)) {
            availableYears.push(year);
        }
        
        // Estrai il valore del ricavo
        const ricavoStr = row.RICAVO.replace('€', '').trim();
        const ricavo = convertNumber(ricavoStr);
        
        // Determina la settimana dell'anno usando la norma ISO 8601
        const date = new Date(Date.UTC(year, monthIndex, day));
        const weekNumber = getWeekNumber(date);
        
        // Chiave per il periodo (settimana, mese o anno)
        const weekKey = `${weekNumber}/${year}`;
        const monthKey = `${monthIndex + 1}/${year}`;
        const yearKey = `${year}`;
        
        // Inizializza l'oggetto del periodo se non esiste
        if (chartPeriod === 'weekly') {
            if (!datiPerPeriodo[weekKey]) {
                datiPerPeriodo[weekKey] = { guadagni: 0, perdite: 0, profitti: 0 };
            }
            
            // Aggiungi il valore al periodo appropriato
            if (ricavo > 0) {
                datiPerPeriodo[weekKey].guadagni += ricavo;
                datiPerPeriodo[weekKey].profitti += ricavo; // Aggiungi ai profitti
            } else if (ricavo < 0) {
                // Salviamo le perdite come valore negativo per il grafico (rivolte verso il basso)
                datiPerPeriodo[weekKey].perdite += ricavo; // Manteniamo il segno negativo
                datiPerPeriodo[weekKey].profitti += ricavo; // Somma algebrica per i profitti
            }
        } else if (chartPeriod === 'monthly') {
            if (!datiPerPeriodo[monthKey]) {
                datiPerPeriodo[monthKey] = { guadagni: 0, perdite: 0, profitti: 0 };
            }
            
            // Aggiungi il valore al periodo appropriato
            if (ricavo > 0) {
                datiPerPeriodo[monthKey].guadagni += ricavo;
                datiPerPeriodo[monthKey].profitti += ricavo; // Aggiungi ai profitti
            } else if (ricavo < 0) {
                // Salviamo le perdite come valore negativo per il grafico (rivolte verso il basso)
                datiPerPeriodo[monthKey].perdite += ricavo; // Manteniamo il segno negativo
                datiPerPeriodo[monthKey].profitti += ricavo; // Somma algebrica per i profitti
            }
        } else {
            if (!datiPerPeriodo[yearKey]) {
                datiPerPeriodo[yearKey] = { guadagni: 0, perdite: 0, profitti: 0 };
            }
            
            // Aggiungi il valore al periodo appropriato
            if (ricavo > 0) {
                datiPerPeriodo[yearKey].guadagni += ricavo;
                datiPerPeriodo[yearKey].profitti += ricavo; // Aggiungi ai profitti
            } else if (ricavo < 0) {
                // Salviamo le perdite come valore negativo per il grafico (rivolte verso il basso)
                datiPerPeriodo[yearKey].perdite += ricavo; // Manteniamo il segno negativo
                datiPerPeriodo[yearKey].profitti += ricavo; // Somma algebrica per i profitti
            }
        }
    });
    
    // Ordina gli anni in ordine decrescente
    availableYears.sort((a, b) => b - a);
    
    // Popola il selettore degli anni
    const yearSelector = document.getElementById('chart-year-selector');
    if (yearSelector) {
        // Salva la selezione attuale
        const currentSelection = yearSelector.value;
        
        // Pulisci il selettore
        yearSelector.innerHTML = '';
        
        // Aggiungi le opzioni degli anni
        availableYears.forEach(year => {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = year;
            yearSelector.appendChild(option);
        });
        
        // Se non ci sono anni, aggiungi l'anno corrente
        if (availableYears.length === 0) {
            const option = document.createElement('option');
            option.value = new Date().getFullYear();
            option.textContent = new Date().getFullYear();
            yearSelector.appendChild(option);
        }
        
        // Ripristina la selezione precedente o imposta il primo anno disponibile
        if (currentSelection && availableYears.includes(parseInt(currentSelection))) {
            yearSelector.value = currentSelection;
        } else if (availableYears.length > 0) {
            yearSelector.value = availableYears[0];
            chartYear = availableYears[0];
        }
    }
    
    // Funzione per calcolare il numero della settimana in un anno secondo lo standard ISO 8601
    function getWeekNumber(date) {
        // Crea una copia della data
        const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        
        // Calcola il giovedì della stessa settimana
        d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
        
        // Inizio dell'anno che contiene il giovedì
        const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
        
        // Calcola il numero della settimana
        // 86400000 = millisecondi in un giorno
        return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    }
    
    // Funzione per ottenere la data di inizio di una settimana (ISO 8601 standard)
    function getFirstDayOfWeek(year, weekNumber) {
        // Calcola il primo giorno dell'anno
        const firstDayOfYear = new Date(Date.UTC(year, 0, 1));
        
        // Se il primo giorno è un venerdì, il primo giorno della settimana 1 è il 4 gennaio
        // Se è un sabato, è il 3 gennaio
        // Se è una domenica, è il 2 gennaio
        // Altrimenti, è lo stesso primo gennaio più i giorni necessari per arrivare a lunedì
        const dayOfWeek = firstDayOfYear.getDay() || 7; // Converti 0 (domenica) in 7
        const daysOffset = (dayOfWeek <= 4) ? 1 - dayOfWeek : 8 - dayOfWeek;
        
        // Primo giorno della prima settimana dell'anno
        const firstDayOfFirstWeek = new Date(Date.UTC(year, 0, 1 + daysOffset));
        
        // Calcola primo giorno della settimana richiesta
        const result = new Date(firstDayOfFirstWeek);
        result.setDate(firstDayOfFirstWeek.getDate() + (weekNumber - 1) * 7);
        
        return result;
    }
    
    // Funzione per ottenere la data di fine di una settimana (ISO 8601 standard)
    function getLastDayOfWeek(year, weekNumber) {
        // Ottieni il primo giorno della settimana richiesta (lunedì)
        const firstDay = getFirstDayOfWeek(year, weekNumber);
        
        // Aggiungi 6 giorni per ottenere l'ultimo giorno della stessa settimana (domenica)
        const lastDay = new Date(firstDay);
        lastDay.setDate(firstDay.getDate() + 6);
        
        return lastDay;
    }
    
    // Funzione per formattare una data nel formato "12 Set 2025"
    function formatDateShort(date) {
        const day = date.getDate();
        const monthNames = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear();
        
        return `${day} ${month} ${year}`;
    }
    
    // Variabili per tenere traccia dei dati da visualizzare
    let periodi = [];
    let guadagni = [];
    let perdite = [];
    let profitti = [];
    
    if (chartPeriod === 'weekly') {
        // Ottieni tutte le settimane dell'anno selezionato
        // Ci sono 52-53 settimane in un anno
        const totalWeeks = 53;
        
        // Ottieni la settimana corrente solo se stiamo visualizzando l'anno corrente
        const currentDate = new Date();
        const currentYear = currentDate.getFullYear();
        const isCurrentYear = chartYear === currentYear;
        
        // Se è l'anno corrente, mostra solo fino alla settimana attuale
        // altrimenti mostra tutte le settimane dell'anno
        const maxWeek = isCurrentYear ? getWeekNumber(currentDate) : totalWeeks;
        
        for (let i = 1; i <= maxWeek; i++) {
            const weekKey = `${i}/${chartYear}`;
            periodi.push(`Sett ${i}`);
            
            if (datiPerPeriodo[weekKey]) {
                guadagni.push(datiPerPeriodo[weekKey].guadagni);
                perdite.push(datiPerPeriodo[weekKey].perdite);
                profitti.push(datiPerPeriodo[weekKey].profitti);
            } else {
                guadagni.push(0);
                perdite.push(0);
                profitti.push(0);
            }
        }
    } else if (chartPeriod === 'monthly') {
        // Ottieni tutti i mesi dell'anno selezionato
        const monthNames = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
        
        // Ottieni il mese corrente solo se stiamo visualizzando l'anno corrente
        const currentDate = new Date();
        const currentYear = currentDate.getFullYear();
        const isCurrentYear = chartYear === currentYear;
        
        // Se è l'anno corrente, mostra solo fino al mese attuale
        // altrimenti mostra tutti i mesi dell'anno
        const maxMonth = isCurrentYear ? currentDate.getMonth() + 1 : 12; // +1 perché i mesi partono da 0
        
        for (let i = 0; i < maxMonth; i++) {
            const monthKey = `${i + 1}/${chartYear}`;
            periodi.push(monthNames[i]);
            
            if (datiPerPeriodo[monthKey]) {
                guadagni.push(datiPerPeriodo[monthKey].guadagni);
                perdite.push(datiPerPeriodo[monthKey].perdite);
                profitti.push(datiPerPeriodo[monthKey].profitti);
            } else {
                guadagni.push(0);
                perdite.push(0);
                profitti.push(0);
            }
        }
    } else {
        // Per il grafico annuale, mostra tutti gli anni disponibili in ordine crescente
        // Crea una copia dell'array degli anni e ordinali in ordine crescente per il grafico
        const yearsForChart = [...availableYears].sort((a, b) => a - b);
        
        yearsForChart.forEach(year => {
            const yearKey = `${year}`;
            periodi.push(year.toString());
            
            if (datiPerPeriodo[yearKey]) {
                guadagni.push(datiPerPeriodo[yearKey].guadagni);
                perdite.push(datiPerPeriodo[yearKey].perdite);
                profitti.push(datiPerPeriodo[yearKey].profitti);
            } else {
                guadagni.push(0);
                perdite.push(0);
                profitti.push(0);
            }
        });
    }
    
    // Ottieni il contesto del canvas
    const ctx = document.getElementById('trade-chart').getContext('2d');
    
    // Distruggi il grafico esistente se presente
    if (tradeChart) {
        tradeChart.destroy();
    }
    
    // Verifica se ci sono dati da visualizzare
    const hasDati = guadagni.some(val => val > 0) || perdite.some(val => val > 0) || profitti.some(val => val !== 0);
    
    // Se non ci sono dati, mostra un messaggio
    if (!hasDati) {
        // Pulisci il canvas
        ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
        
        // Mostra un messaggio
        ctx.font = '16px Arial';
        ctx.fillStyle = '#666';
        ctx.textAlign = 'center';
        ctx.fillText('Nessun dato disponibile per questo periodo', ctx.canvas.width / 2, ctx.canvas.height / 2);
        
        // Nascondi il selettore dell'anno se in modalità annuale
        if (yearSelector) {
            yearSelector.style.display = chartPeriod === 'monthly' ? 'block' : 'none';
        }
        
        return;
    }
    
    // Crea il nuovo grafico
    tradeChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: periodi,
            datasets: [
                {
                    label: 'Guadagni',
                    data: guadagni,
                    backgroundColor: 'rgba(46, 204, 113, 0.8)',
                    borderColor: 'rgba(46, 204, 113, 1)',
                    borderWidth: 1,
                    borderRadius: 6,
                    maxBarThickness: 40
                },
                {
                    label: 'Perdite',
                    data: perdite,
                    backgroundColor: 'rgba(231, 76, 60, 0.8)',
                    borderColor: 'rgba(231, 76, 60, 1)',
                    borderWidth: 1,
                    borderRadius: 6,
                    maxBarThickness: 40
                },
                {
                    label: 'Profitti',
                    data: profitti,
                    backgroundColor: 'rgba(52, 152, 219, 0.8)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 1,
                    borderRadius: 6,
                    maxBarThickness: 40,
                    // Utilizziamo il tipo 'line' per questo dataset per distinguerlo meglio
                    type: 'bar'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false,
            },
            scales: {
                y: {
                    beginAtZero: false, // Consenti valori negativi sull'asse Y
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)',
                        lineWidth: 1
                    },
                    border: {
                        dash: [4, 4]
                    },
                    ticks: {
                        font: {
                            size: 11,
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        padding: 8,
                        callback: function(value) {
                            return value.toLocaleString('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 0 });
                        }
                    },
                    title: {
                        display: true,
                        text: 'Euro (€)',
                        font: {
                            size: 13,
                            weight: 'bold',
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        padding: {top: 10, bottom: 10}
                    }
                },
                x: {
                    grid: {
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)',
                        lineWidth: 1,
                        drawOnChartArea: true,
                        drawTicks: false,
                        borderDash: [5, 5]
                    },
                    ticks: {
                        font: {
                            size: 11,
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        padding: 8,
                        callback: function(value, index, values) {
                            // Per la vista settimanale, mostra solo alcune etichette per evitare sovraffollamento
                            if (chartPeriod === 'weekly') {
                                return (index % 4 === 0 || index === 0 || index === values.length - 1) ? this.getLabelForValue(value) : '';
                            }
                            return this.getLabelForValue(value);
                        }
                    },
                    title: {
                        display: true,
                        text: chartPeriod === 'weekly' ? 'Settimana' : (chartPeriod === 'monthly' ? 'Mese' : 'Anno'),
                        font: {
                            size: 13,
                            weight: 'bold',
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        padding: {top: 10, bottom: 10}
                    }
                }
            },
            plugins: {
                tooltip: {
                    backgroundColor: 'rgba(255, 255, 255, 1)', // Rimossa trasparenza: da 0.9 a 1
                    titleColor: '#333',
                    titleFont: {
                        size: 13,
                        weight: 'bold',
                        family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                    },
                    bodyColor: '#333',
                    bodyFont: {
                        size: 12,
                        family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                    },
                    borderColor: 'rgba(0, 0, 0, 0.1)',
                    borderWidth: 1,
                    cornerRadius: 8,
                    boxPadding: 6,
                    usePointStyle: true,
                    callbacks: {
                        title: function(context) {
                            // Se siamo in modalità settimanale, mostra l'intervallo di date della settimana
                            if (chartPeriod === 'weekly') {
                                const weekIndex = context[0].dataIndex;
                                // Estrai il numero della settimana dal testo dell'etichetta (es. "Sett 7" -> 7)
                                const weekLabel = context[0].label;
                                const weekNumber = parseInt(weekLabel.split(' ')[1]);
                                
                                // Calcola le date di inizio e fine della settimana
                                const firstDay = getFirstDayOfWeek(chartYear, weekNumber);
                                const lastDay = getLastDayOfWeek(chartYear, weekNumber);
                                
                                // Formatta le date e restituisci l'intervallo
                                return `${formatDateShort(firstDay)} - ${formatDateShort(lastDay)}`;
                            }
                            // Altrimenti, usa il titolo predefinito
                            return context[0].label;
                        },
                        label: function(context) {
                            const value = context.raw;
                            // Per le perdite, mostriamo il valore assoluto per leggibilità, ma con il prefisso "-"
                            const formattedValue = 
                                context.dataset.label === 'Perdite' && value < 0 
                                    ? '-' + Math.abs(value).toLocaleString('it-IT', { style: 'currency', currency: 'EUR' })
                                    : value.toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
                            
                            return [
                                context.dataset.label + ': ' + formattedValue
                            ];
                        }
                    }
                },
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        usePointStyle: true,
                        padding: 20
                    }
                },
                title: {
                    display: true,
                    text: chartPeriod === 'weekly'
                        ? `Andamento Ricavi Settimanali ${chartYear}`
                        : (chartPeriod === 'monthly' 
                            ? `Andamento Ricavi Mensili ${chartYear}`
                            : 'Andamento Ricavi Annuali'),
                    font: {
                        size: 16,
                        weight: 'bold',
                        family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                    },
                    padding: {top: 10, bottom: 20},
                    color: '#333'
                }
            }
        },
        plugins: [separatorPlugin, zeroLinePlugin]
    });
    
    // Nascondi il selettore dell'anno solo in modalità annuale, mostralo in weekly e monthly
    if (yearSelector) {
        yearSelector.style.display = chartPeriod === 'yearly' ? 'none' : 'block';
    }
}

// Funzione per impostare i gestori eventi per i grafici
function setupChartEventHandlers() {
    // Selettori per i controlli del grafico
    const periodSelectors = document.querySelectorAll('input[name="chart-period"]');
    const yearSelector = document.getElementById('chart-year-selector');
    
    // Aggiungi anni al selettore se disponibili
    if (availableYears.length > 0) {
        // Ordina gli anni in ordine decrescente (dal più recente)
        availableYears.sort((a, b) => b - a);
        
        // Svuota il selettore
        yearSelector.innerHTML = '';
        
        // Aggiungi le opzioni
        availableYears.forEach(year => {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = year;
            yearSelector.appendChild(option);
        });
        
        // Seleziona l'anno corrente o il più recente disponibile
        const currentYear = new Date().getFullYear();
        if (availableYears.includes(currentYear)) {
            yearSelector.value = currentYear;
            chartYear = currentYear;
        } else if (availableYears.length > 0) {
            yearSelector.value = availableYears[0];
            chartYear = availableYears[0];
        }
    } else {
        // Se non ci sono dati, mostra solo l'anno corrente
        const currentYear = new Date().getFullYear();
        const option = document.createElement('option');
        option.value = currentYear;
        option.textContent = currentYear;
        yearSelector.innerHTML = '';
        yearSelector.appendChild(option);
        chartYear = currentYear;
    }
    
    // Gestisci il cambio di periodo (mensile/annuale)
    periodSelectors.forEach(radio => {
        radio.addEventListener('change', function() {
            // Effetto visivo: opacità ridotta durante l'aggiornamento
            const chartContainer = document.querySelector('.chart-container');
            chartContainer.style.opacity = '0.5';
            chartContainer.style.transition = 'opacity 0.3s ease';
            
            // Aggiorna la selezione del periodo
            chartPeriod = this.value;
            
            // Mostra il selettore dell'anno sia per settimanale che per mensile, ma non per annuale
            document.querySelector('.chart-controls-container').style.display = 
                chartPeriod === 'yearly' ? 'none' : 'block';
            
            // Aggiorna il grafico con breve ritardo per l'animazione
            setTimeout(() => {
                updateTradeChart();
                // Ripristina l'opacità
                chartContainer.style.opacity = '1';
            }, 300);
        });
    });
    
    // Gestisci il cambio di anno
    yearSelector.addEventListener('change', function() {
        // Effetto visivo: opacità ridotta durante l'aggiornamento
        const chartContainer = document.querySelector('.chart-container');
        chartContainer.style.opacity = '0.5';
        chartContainer.style.transition = 'opacity 0.3s ease';
        
        // Aggiorna l'anno selezionato
        chartYear = parseInt(this.value);
        
        // Aggiorna il grafico con breve ritardo per l'animazione
        setTimeout(() => {
            updateTradeChart();
            // Ripristina l'opacità
            chartContainer.style.opacity = '1';
        }, 300);
    });
    
    // Nascondi il selettore dell'anno se la visualizzazione è annuale
    if (chartPeriod === 'yearly') {
        document.querySelector('.chart-controls-container').style.display = 'none';
    }
}

// Aggiorniamo la funzione calculateAndDisplayTotals per aggiornare anche i grafici
const originalCalculateAndDisplayTotals = calculateAndDisplayTotals;
calculateAndDisplayTotals = function() {
    // Chiama la funzione originale
    originalCalculateAndDisplayTotals();
    
    // Aggiorna il grafico
    updateTradeChart();
};

// Aggiungiamo l'inizializzazione del grafico quando il documento è pronto
document.addEventListener('DOMContentLoaded', function() {
    // ... other code ...
    
    // Setup dei grafici
    setupChartEventHandlers();
    
    // ... other code ...
});

// Aggiorniamo la funzione populateFinalTable per aggiornare anche i grafici
const originalPopulateFinalTable = populateFinalTable;
populateFinalTable = function() {
    // Chiama la funzione originale
    originalPopulateFinalTable();
    
    // Aggiorna il grafico
    updateTradeChart();
};