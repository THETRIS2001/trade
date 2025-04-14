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
        downloadCSV(outputTable, 'output_data.csv');
    });
    
    downloadFinalBtn.addEventListener('click', () => {
        downloadCSV(finalTable, 'final_data.csv');
    });
    
    // Event listeners per ordinare le colonne delle tabelle
    const outputHeaders = outputTable.querySelectorAll('th[data-sort]');
    const finalHeaders = finalTable.querySelectorAll('th[data-sort]');
    
    // Applica gli event handlers alle intestazioni di entrambe le tabelle
    handleHeaderClick(outputHeaders, outputTable);
    handleHeaderClick(finalHeaders, finalTable);
    
    // Implemento la funzione a livello globale
    setupAdditionalValuesHandling = function() {
        // Funzione per aggiornare le opzioni del selettore di righe
        function updateRowSelectOptions() {
            console.log("Aggiornamento opzioni del selettore di righe");
            const associateRowSelect = document.getElementById('associate-row');
            
            if (!associateRowSelect) {
                console.error("Elemento select 'associate-row' non trovato!");
                return;
            }
            
            // Salva gli elementi selezionati prima di cancellare
            const selectedValue = associateRowSelect.value;
            
            // Pulisci completamente il selettore
            associateRowSelect.innerHTML = '';
            
            // Aggiungi solo l'opzione separatore iniziale
            const separatorOption = document.createElement('option');
            separatorOption.disabled = true;
            separatorOption.textContent = '-- Seleziona una riga dal riepilogo --';
            associateRowSelect.appendChild(separatorOption);
            
            // Assicurati che sia selezionata la prima opzione (il separatore)
            associateRowSelect.selectedIndex = 0;
            
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
                    // Crea la descrizione della riga
                    let rowDescription = `${row.TIPO_OPERAZIONE} - ${row.Descrizione || 'Senza descrizione'}`;
                    
                    // Controlla se questa descrizione è già stata aggiunta
                    if (addedDescriptions.has(rowDescription)) {
                        console.log(`Saltata riga duplicata: ${rowDescription}`);
                        return; // salta questa iterazione
                    }
                    
                    // Aggiungi la descrizione al set per controllare i duplicati
                    addedDescriptions.add(rowDescription);
                    
                    // Tronca la descrizione se è troppo lunga
                    const displayDescription = rowDescription.length > 60 
                        ? rowDescription.substring(0, 60) + '...' 
                        : rowDescription;
                    
                    const option = document.createElement('option');
                    option.value = `row-${index}`;
                    option.textContent = displayDescription;
                    
                    // Aggiungi attributi per i dettagli completi della riga
                    option.setAttribute('data-tipo', row.TIPO_OPERAZIONE || '');
                    option.setAttribute('data-desc', row.Descrizione || '');
                    
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
            }
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
            if (!selectedRowId || selectedRowId === 'global' || associateRowSelect.selectedIndex <= 0) {
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
            
            // Se ci sono dati elaborati, ricalcola i totali
            if (Object.keys(finalData).length > 0) {
                // Ricrea il risultato finale con i nuovi valori
                populateFinalTable();
                calculateAndDisplayTotals();
                
                // Aggiungi la riga al dettaglio operazioni
                addValueToOutputTable(newValue, rowDescription);
                
                // Aggiorna la data del riepilogo
                updateSummaryDate();
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
    }
    
    // Setup delle tabelle e dei selettori di righe
    setupTablesAndSelectors();
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
        
        // Crea celle per ogni colonna con attributo title per mostrare il testo completo
        tr.innerHTML = `
            <td title="${formattedDate || ''}">${formattedDate || ''}</td>
            <td title="${row.TIPO || ''}">${row.TIPO || ''}</td>
            <td title="${row.TIPO_OPERAZIONE || ''}">${row.TIPO_OPERAZIONE || ''}</td>
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
            <td title="${row.RICAVO || ''}">${row.RICAVO || ''}</td>
            <td title="${row.RICAVO_PERCENTUALE || ''}">${row.RICAVO_PERCENTUALE || ''}</td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Ordina la tabella per data di inizio (dalla più recente alla più lontana)
    sortTable(document.getElementById('final-table'), 'data-inizio', false); // Imposta l'ordinamento decrescente
}

// Funzione per creare il risultato finale
function createFinalResult() {
    const risultatoFinale = [];
    
    // Estraggo gli ISIN dai warrant exercise per l'abbinamento
    const warrantIsinDict = {};
    
    finalData.dfWarrant.forEach(row => {
        if (row.Descrizione && row.Descrizione.includes('ISIN')) {
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
        // Prepara il risultato
        const result = {
            dataInizio: null,
            dataFine: null
        };
        
        // Cerca la descrizione in tutte le righe del dettaglio operazioni
        const datePerDescrizione = [];
        outputTableRows.forEach(row => {
            const descCell = row.querySelector('td:nth-child(4)');
            if (descCell && descCell.textContent.trim() === descrizione) {
                const date = estraiData(row);
                if (date) {
                    datePerDescrizione.push(date);
                }
            }
        });
        
        // Se abbiamo trovato date, calcoliamo min e max
        if (datePerDescrizione.length > 0) {
            result.dataInizio = new Date(Math.min.apply(null, datePerDescrizione));
            result.dataFine = new Date(Math.max.apply(null, datePerDescrizione));
        }
        
        return result;
    };
    
    for (const descrizione of descrizioniUniche) {
        const buyRows = finalData.dfCommercioBuy.filter(row => row.Descrizione === descrizione);
        const sellRow = finalData.dfCommercioSell.find(row => row.Descrizione === descrizione);
        
        // Calcola date min e max per questa descrizione
        const { dataInizio, dataFine } = trovateDate(descrizione);
        const dataInizioFormatted = formatDate(dataInizio);
        const dataFineFormatted = formatDate(dataFine);
        
        // Estrai l'ISIN o il codice identificativo dal buy_row (se esiste)
        let isinToMatch = null;
        if (buyRows.length > 0) {
            const descParts = descrizione.split(' ');
            if (descParts.length > 0) {
                isinToMatch = descParts[0]; // Prendiamo la prima parte che dovrebbe essere l'ISIN
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
                        TIPO_OPERAZIONE: 'Trade',
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
        
        risultatoFinale.push({
            DATA_INIZIO: dataMinInteressi,
            DATA_FINE: dataMaxInteressi,
            TIPO_OPERAZIONE: 'Interessi Giacenza',
            Descrizione: row.Descrizione,
            Quantita: row.Quantita,
            PREZZO_DI_ACQUISTO: row.PREZZO_DI_ACQUISTO,
            PREZZO_DI_VENDITA: prezzoVendita,
            PREZZO_UNITARIO_ACQUISTO: '',
            PREZZO_UNITARIO_VENDITA: '',
            RICAVO: '',
            RICAVO_PERCENTUALE: ''
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
        
        risultatoFinale.push({
            DATA_INIZIO: dataMinSaveback,
            DATA_FINE: dataMaxSaveback,
            TIPO_OPERAZIONE: 'Saveback',
            Descrizione: row.Descrizione,
            Quantita: row.Quantita,
            PREZZO_DI_ACQUISTO: row.PREZZO_DI_ACQUISTO,
            PREZZO_DI_VENDITA: prezzoVendita,
            PREZZO_UNITARIO_ACQUISTO: '',
            PREZZO_UNITARIO_VENDITA: '',
            RICAVO: '',
            RICAVO_PERCENTUALE: ''
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
                if (descUnica.includes(isin)) {
                    abbinato = true;
                    break;
                }
            }
            
            if (!abbinato) {
                nonAbbinati.push({
                    TIPO_OPERAZIONE: 'Warrant Exercise',
                    Descrizione: row.Descrizione,
                    Quantita: row.Quantita,
                    PREZZO_DI_ACQUISTO: row.PREZZO_DI_ACQUISTO,
                    PREZZO_DI_VENDITA: row.PREZZO_DI_VENDITA,
                    PREZZO_UNITARIO_ACQUISTO: '',
                    PREZZO_UNITARIO_VENDITA: '',
                    RICAVO: '',
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
            'Sell trade': 3,
            'Warrant Exercise': 4,
            'Interessi Giacenza': 5,
            'Saveback': 6
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
    
    // Calcolo totale dei Trade (esclusi i Buy trade che sono quelli attualmente investiti)
    const datiTrade = createFinalResult().filter(row => row.TIPO_OPERAZIONE === 'Trade');
    
    // Sommiamo i valori monetari
    let totaleAcquisto = 0;
    let totaleVendita = 0;
    let totaleRicavo = 0;
    
    // Variabile per tenere traccia dei valori aggiunti ai Buy trade
    let totaleValoriAggiuntiBuyTrade = 0;
    
    datiTrade.forEach(row => {
        // Estraggo i valori numerici (rimuovendo i simboli)
        const acquisto = row.PREZZO_DI_ACQUISTO ? row.PREZZO_DI_ACQUISTO.replace('€', '').replace('/azione', '').trim() : '';
        const vendita = row.PREZZO_DI_VENDITA ? row.PREZZO_DI_VENDITA.replace('€', '').replace('/azione', '').trim() : '';
        const ricavo = row.RICAVO ? row.RICAVO.replace('€', '').replace('/azione', '').trim() : '';
        
        // Converto in numeri
        const acquistoVal = convertNumber(acquisto);
        const venditaVal = convertNumber(vendita);
        const ricavoVal = convertNumber(ricavo);
        
        // Aggiungo al totale
        totaleAcquisto += acquistoVal;
        totaleVendita += venditaVal;
        totaleRicavo += ricavoVal;
    });
    
    // Controlla se ci sono valori aggiunti associati a Buy trade
    if (additionalValues && additionalValues.length > 0) {
        additionalValues.forEach(item => {
            // Estrai l'ID della riga dal formato "row-X"
            const matches = item.rowId.match(/row-(\d+)/);
            if (matches && matches.length > 1) {
                const rowIndex = parseInt(matches[1]);
                const risultati = createFinalResult();
                
                // Se l'indice è valido
                if (rowIndex >= 0 && rowIndex < risultati.length) {
                    // Verifica se la riga era originariamente un Buy trade
                    // Questo controllo va fatto sulla descrizione associata al valore
                    if (item.description && item.description.includes("Buy trade")) {
                        // Aggiungi l'intero valore al totale
                        totaleValoriAggiuntiBuyTrade += item.value;
                        
                        // IMPORTANTE: Dobbiamo sottrarre il ricavo perché è già incluso nel totaleRicavo
                        // e non vogliamo contarlo due volte
                        const row = risultati[rowIndex];
                        if (row && row.RICAVO) {
                            const ricavo = convertNumber(row.RICAVO.replace('€', '').trim());
                            totaleRicavo -= ricavo; // Rimuoviamo il ricavo già conteggiato
                        }
                    }
                }
            }
        });
    }
    
    // Non ci sono più valori globali, ora tutti i valori sono associati a righe specifiche
    // e sono già inclusi nel calcolo del totaleRicavo attraverso le righe modificate
    
    // Estraggo i valori di interessi e saveback
    let interessiVal = 0;
    let savebackVal = 0;
    
    // Estrai il valore degli interessi
    const rigaInteressi = createFinalResult().find(row => row.TIPO_OPERAZIONE === 'Interessi Giacenza');
    if (rigaInteressi && rigaInteressi.PREZZO_DI_VENDITA) {
        // Rimuovo il simbolo € e converto in numero
        const interessiStr = rigaInteressi.PREZZO_DI_VENDITA.replace('€', '').trim();
        interessiVal = convertNumber(interessiStr);
    }
    
    // Estrai il valore del saveback
    const rigaSaveback = createFinalResult().find(row => row.TIPO_OPERAZIONE === 'Saveback');
    if (rigaSaveback && rigaSaveback.PREZZO_DI_VENDITA) {
        // Rimuovo il simbolo € e converto in numero
        const savebackStr = rigaSaveback.PREZZO_DI_VENDITA.replace('€', '').trim();
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
    const totalComprehensiveValue = ricavoTotaleComplessivo + totaleValoriAggiuntiBuyTrade;
    totalComprehensive.textContent = `${normalizeNumericFormat(Math.round(totalComprehensiveValue * 1000) / 1000, true)} €`;
    
    // Applica le classi di colore in base ai valori
    
    // Rimuovi tutte le classi di colore per iniziare
    if (totalTradeCard) {
        totalTradeCard.classList.remove('profit-positive', 'profit-negative');
        if (totalTradeValue > 0) {
            totalTradeCard.classList.add('profit-positive');
        } else if (totalTradeValue < 0) {
            totalTradeCard.classList.add('profit-negative');
        }
    }
    
    // Per il Ricavo Totale Complessivo
    if (totalComprehensiveCard) {
        totalComprehensiveCard.classList.remove('profit-positive', 'profit-negative');
        totalComprehensiveCard.classList.add('comprehensive'); // Aggiungi sempre questa classe
        if (totalComprehensiveValue > 0) {
            totalComprehensiveCard.classList.add('profit-positive');
        } else if (totalComprehensiveValue < 0) {
            totalComprehensiveCard.classList.add('profit-negative');
        }
    }
    
    // Per Interessi e Saveback
    if (totalInterestCard) {
        totalInterestCard.classList.add('interest');
    }
    
    if (totalSavebackCard) {
        totalSavebackCard.classList.add('saveback');
    }
}

// Funzione per popolare le tabelle e inizializzare il selettore di righe
function setupTablesAndSelectors() {
    // Definisci la funzione per inizializzare e aggiornare il selettore di righe
    const updateRowSelector = setupAdditionalValuesHandling();
    
    // Aggiungi un listener per il tab finale
    const finalTabBtn = document.querySelector('[data-tab="final-csv"]');
    if (finalTabBtn) {
        finalTabBtn.addEventListener('click', () => {
            // Quando si clicca sul tab di riepilogo, aggiorna il selettore di righe
            if (typeof updateRowSelector === 'function') {
                updateRowSelector();
            }
        });
    }
    
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
                    
                } else if (row.TIPO_OPERAZIONE === 'Trade') {
                    // Per i Trade esistenti, aggiorniamo il prezzo di vendita e ricalcoliamo
                    
                    // Estrai i valori esistenti
                    const prezzoAcquisto = convertNumber(row.PREZZO_DI_ACQUISTO.replace('€', '').trim());
                    const vecchioPrezzoVendita = convertNumber(row.PREZZO_DI_VENDITA.replace('€', '').trim());
                    
                    // Calcola il nuovo prezzo di vendita
                    const nuovoPrezzoVendita = vecchioPrezzoVendita + value;
                    row.PREZZO_DI_VENDITA = `${normalizeNumericFormat(nuovoPrezzoVendita, true)} €`;
                    
                    // Aggiorna il prezzo unitario di vendita
                    const quantita = convertNumber(row.Quantita.replace('azioni', '').trim());
                    if (quantita > 0) {
                        // Rimuovi la parte "€/azione" e mantieni solo il valore numerico
                        row.PREZZO_UNITARIO_VENDITA = `${normalizeNumericFormat(nuovoPrezzoVendita / quantita, true)} €/azione`;
                    } else {
                        row.PREZZO_UNITARIO_VENDITA = '/';
                        row.Quantita = '/';
                    }
                    
                    // Ricalcola il ricavo (aggiungendo direttamente il valore aggiunto)
                    // Correggiamo qui: il ricavo deve aumentare del valore aggiunto
                    // anziché essere calcolato come differenza tra nuovo prezzo di vendita e prezzo di acquisto
                    const vecchioRicavo = convertNumber(row.RICAVO.replace('€', '').trim());
                    const nuovoRicavo = vecchioRicavo + value;
                    row.RICAVO = `${normalizeNumericFormat(nuovoRicavo, true)} €`;
                    
                    // Ricalcola il ricavo percentuale
                    if (prezzoAcquisto > 0) {
                        const ricavoPercentuale = (nuovoRicavo / prezzoAcquisto) * 100;
                        row.RICAVO_PERCENTUALE = `${normalizeNumericFormat(ricavoPercentuale, true)} %`;
                    }
                    
                    console.log(`Riga ${rowIndex} (Trade) aggiornata: nuovo prezzo vendita ${nuovoPrezzoVendita}, ricavo ${nuovoRicavo}`);
                    
                } else if (row.TIPO_OPERAZIONE === 'Interessi Giacenza' || row.TIPO_OPERAZIONE === 'Saveback') {
                    // Per Interessi e Saveback, aggiungiamo il valore al prezzo di vendita
                    
                    const vecchioPrezzoVendita = convertNumber(row.PREZZO_DI_VENDITA.replace('€', '').trim());
                    const nuovoPrezzoVendita = vecchioPrezzoVendita + value;
                    
                    row.PREZZO_DI_VENDITA = `${normalizeNumericFormat(nuovoPrezzoVendita, true)} €`;
                    
                    // Aggiorna sempre la DATA_FINE con la data odierna
                    row.DATA_FINE = todayFormatted;
                    
                    console.log(`Riga ${rowIndex} (${row.TIPO_OPERAZIONE}) aggiornata: nuovo prezzo vendita ${nuovoPrezzoVendita}`);
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
    const tbody = document.querySelector('#output-table tbody');
    const tr = document.createElement('tr');
    
    // Aggiungi la classe per evidenziare questa riga
    tr.classList.add('added-value-row');
    
    // Ottieni la data di oggi in formato DD/MM/YYYY
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const monthNames = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
    const month = monthNames[today.getMonth()]; // Ottieni il nome del mese
    const year = today.getFullYear().toString().slice(2); // Prendi solo le ultime 2 cifre
    
    // Formatta la data come "13 Apr 25"
    const formattedDate = `${day} ${month} ${year}`;
    
    // Formatta il valore con la virgola come separatore decimale
    const formattedValue = normalizeNumericFormat(value, true) + " €";
    
    // Inizializza quantità e prezzo unitario vuoti
    let quantity = "";
    let unitPrice = "";
    
    // Inizializza tipo e tipo operazione con i valori predefiniti
    let tipoCell = "Commercio";
    let tipoOperazioneCell = "Sell trade";
    
    // Controlla il tipo di valore aggiunto in base alla descrizione
    const isBuyTrade = description.includes("Buy trade");
    const isInteressi = description.includes("Interessi Giacenza");
    const isSaveback = description.includes("Saveback");
    
    // Imposta i valori corretti in base al tipo
    if (isInteressi) {
        tipoCell = "Commercio";
        tipoOperazioneCell = "Savings plan execution";
    } else if (isSaveback) {
        tipoCell = "Premio";
        tipoOperazioneCell = "Your Saveback payment";
    }
    
    // Solo se è un Buy trade, cerca la quantità nel riepilogo finale
    if (isBuyTrade) {
        const finalTable = document.querySelector('#final-table tbody');
        const rows = finalTable.querySelectorAll('tr');
        
        // Estrai la parte della descrizione da confrontare (solo il testo dopo il primo trattino, senza lo spazio)
        const descriptionParts = description.split(' - ');
        let descToMatch = description;
        
        // Se la descrizione è nel formato "Tipo - Descrizione", prendiamo solo la Descrizione
        if (descriptionParts.length > 1) {
            descToMatch = descriptionParts.slice(1).join(' - '); // In caso ci siano più " - " nella descrizione
        }
        
        // Cerca nel riepilogo finale la riga con la descrizione corrispondente
        for (const row of rows) {
            const descCell = row.querySelector('td:nth-child(2)');
            if (descCell && descCell.textContent.trim() === descToMatch) {
                const quantityCell = row.querySelector('td:nth-child(3)');
                if (quantityCell) {
                    quantity = quantityCell.textContent.trim();
                    
                    // Calcola il prezzo unitario
                    const quantityNum = convertNumber(quantity.replace('azioni', '').trim());
                    if (quantityNum > 0) {
                        const unitPriceNum = value / quantityNum;
                        unitPrice = normalizeNumericFormat(unitPriceNum, true) + " €/azione";
                    }
                    
                    break;
                }
            }
        }
    }
    
    // Estrai la parte della descrizione da visualizzare
    let displayDesc = description;
    const descriptionParts = description.split(' - ');
    if (descriptionParts.length > 1) {
        displayDesc = descriptionParts.slice(1).join(' - '); // Prendi tutto dopo il primo "Tipo - "
    }
    
    // Crea celle per ogni colonna
    tr.innerHTML = `
        <td title="${formattedDate}">${formattedDate}</td>
        <td title="${tipoCell}">${tipoCell}</td>
        <td title="${tipoOperazioneCell}">${tipoOperazioneCell}</td>
        <td title="${displayDesc}">${displayDesc}</td>
        <td title="${quantity}">${quantity}</td>
        <td title=""></td>
        <td title="${formattedValue}">${formattedValue}</td>
        <td title="${unitPrice}">${unitPrice}</td>
    `;
    
    tbody.appendChild(tr);
    
    // Riordina la tabella dopo aver aggiunto la nuova riga
    sortTable(document.getElementById('output-table'), 'data', false); // Imposta l'ordinamento decrescente
    
    // Aggiorna la data del riepilogo
    updateSummaryDate();
}