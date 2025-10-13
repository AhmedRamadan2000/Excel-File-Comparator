// ---------- Header helpers ----------
function findCreditColumn(row) {
    for (let i = 0; i < row.length; i++) {
        const cell = row[i]?.toString().toLowerCase().trim();
        if (cell && (cell.includes('credit') || cell === 'cr' || cell.includes('amount credit'))) {
            return i;
        }
    }
    return -1;
}

function findDebitColumn(row) {
    for (let i = 0; i < row.length; i++) {
        const cell = row[i]?.toString().toLowerCase().trim();
        if (cell && (cell.includes('debit') || cell === 'dr' || cell.includes('amount debit'))) {
            return i;
        }
    }
    return -1;
}

function findDescriptionColumn(data) {
    for (let rowIndex = 0; rowIndex < Math.min(10, data.length); rowIndex++) {
        const row = data[rowIndex] || [];
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
            if (row[colIndex] &&
                row[colIndex].toString().toLowerCase().trim() === 'description') {
                return {
                    headerRow: rowIndex,
                    columnIndex: colIndex,
                    creditColumn: findCreditColumn(row),
                    debitColumn: findDebitColumn(row)
                };
            }
        }
    }
    return null;
}

// ---------- TP2P helper ----------
function isTP2PMatch(sourceDesc, compareDesc, compareRow, creditColumnIndex) {
    if (!sourceDesc) return false;
    if (!sourceDesc.toString().toUpperCase().trim().endsWith('TP2P')) return false;
    const baseCode = sourceDesc.toString().trim().slice(0, -4).trim();
    const compareText = compareDesc ? compareDesc.toString().trim() : '';
    const compareDescMatch = compareText === baseCode;
    let isCreditEntry = false;
    if (creditColumnIndex !== -1 && compareRow && compareRow[creditColumnIndex]) {
        const creditValue = parseFloat(compareRow[creditColumnIndex].toString().replace(/,/g, ''));
        isCreditEntry = !isNaN(creditValue) && creditValue > 0;
    }
    return compareDescMatch && isCreditEntry;
}

// ---------- Main comparison ----------
function performComparison(source, compare1, compare2) {
    const sourceDescInfo = findDescriptionColumn(source);
    const compare1DescInfo = compare1 ? findDescriptionColumn(compare1) : null;
    const compare2DescInfo = compare2 ? findDescriptionColumn(compare2) : null;

    const results = {
        statistics: {
            sourceRows: 0,
            compare1Rows: 0,
            compare2Rows: 0,
            matchingRows: 0,
            uniqueRows: 0,
            tp2pMatches: 0,
            exactMatches: 0,
            descriptionColumnFound: {
                source: !!sourceDescInfo,
                compare1: !!compare1DescInfo,
                compare2: !!compare2DescInfo
            },
            filesCompared: {
                wallet1: !!compare1,
                wallet2: !!compare2
            }
        },
        matches: [],
        unique: []
    };

    if (!sourceDescInfo) throw new Error('Description column not found in Bank Sheet file.');

    const sourceDataRows = source.slice(sourceDescInfo.headerRow + 1);
    const compare1DataRows = (compare1 && compare1DescInfo) ? compare1.slice(compare1DescInfo.headerRow + 1) : [];
    const compare2DataRows = (compare2 && compare2DescInfo) ? compare2.slice(compare2DescInfo.headerRow + 1) : [];

    results.statistics.sourceRows = sourceDataRows.length;
    results.statistics.compare1Rows = compare1DataRows.length;
    results.statistics.compare2Rows = compare2DataRows.length;

    function getDebitCredit(row, headerInfo) {
        const debit = (headerInfo && headerInfo.debitColumn >= 0) ? row[headerInfo.debitColumn] || '' : '';
        const credit = (headerInfo && headerInfo.creditColumn >= 0) ? row[headerInfo.creditColumn] || '' : '';
        return { debit, credit };
    }

    for (let i = 0; i < sourceDataRows.length; i++) {
        const sourceRow = sourceDataRows[i] || [];
        const sourceDescription = sourceRow[sourceDescInfo.columnIndex];
        if (!sourceDescription || sourceDescription.toString().trim() === '') continue;

        let foundInCompare1 = false;
        let foundInCompare2 = false;
        let matchType1 = '';
        let matchType2 = '';
        let matchingRows = { file1: [], file2: [] };

        if (compare1 && compare1DescInfo) {
            for (let j = 0; j < compare1DataRows.length; j++) {
                const row = compare1DataRows[j] || [];
                const desc = row[compare1DescInfo.columnIndex];
                if (!desc) continue;
                const amounts = getDebitCredit(row, compare1DescInfo);
                if (sourceDescription.toString().toLowerCase().trim() === desc.toString().toLowerCase().trim()) {
                    foundInCompare1 = true; matchType1='exact';
                    matchingRows.file1.push({...row, debit: amounts.debit, credit: amounts.credit, rowIndex: j + compare1DescInfo.headerRow + 2});
                } else if (isTP2PMatch(sourceDescription, desc, row, compare1DescInfo.creditColumn)) {
                    foundInCompare1 = true; matchType1='TP2P';
                    matchingRows.file1.push({...row, debit: amounts.debit, credit: amounts.credit, rowIndex: j + compare1DescInfo.headerRow + 2});
                }
            }
        }

        if (compare2 && compare2DescInfo) {
            for (let j = 0; j < compare2DataRows.length; j++) {
                const row = compare2DataRows[j] || [];
                const desc = row[compare2DescInfo.columnIndex];
                if (!desc) continue;
                const amounts = getDebitCredit(row, compare2DescInfo);
                if (sourceDescription.toString().toLowerCase().trim() === desc.toString().toLowerCase().trim()) {
                    foundInCompare2 = true; matchType2='exact';
                    matchingRows.file2.push({...row, debit: amounts.debit, credit: amounts.credit, rowIndex: j + compare2DescInfo.headerRow + 2});
                } else if (isTP2PMatch(sourceDescription, desc, row, compare2DescInfo.creditColumn)) {
                    foundInCompare2 = true; matchType2='TP2P';
                    matchingRows.file2.push({...row, debit: amounts.debit, credit: amounts.credit, rowIndex: j + compare2DescInfo.headerRow + 2});
                }
            }
        }

        const amounts = getDebitCredit(sourceRow, sourceDescInfo);

        if (foundInCompare1 || foundInCompare2) {
            results.matches.push({
                rowIndex: i + sourceDescInfo.headerRow + 2,
                description: sourceDescription,
                data: sourceRow,
                debit: amounts.debit,
                credit: amounts.credit,
                foundInFile1: foundInCompare1,
                foundInFile2: foundInCompare2,
                matchType1, matchType2,
                matchingRows
            });
            results.statistics.matchingRows++;
            if (matchType1==='TP2P' || matchType2==='TP2P') results.statistics.tp2pMatches++;
            else results.statistics.exactMatches++;
        } else {
            results.unique.push({
                rowIndex: i + sourceDescInfo.headerRow + 2,
                description: sourceDescription,
                data: sourceRow,
                debit: amounts.debit,
                credit: amounts.credit
            });
            results.statistics.uniqueRows++;
        }
    }

    return results;
}

let sourceData = null;
let compareData1 = null;
let compareData2 = null;
let comparisonResults = null;

// File upload handlers
document.getElementById('sourceFile').addEventListener('change', function(e) {
    handleFileUpload(e, 'sourceFileName', 'source');
});

document.getElementById('compareFile1').addEventListener('change', function(e) {
    handleFileUpload(e, 'compareFile1Name', 'compare1');
});

document.getElementById('compareFile2').addEventListener('change', function(e) {
    handleFileUpload(e, 'compareFile2Name', 'compare2');
});

async function handleFileUpload(event, fileNameElementId, type) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById(fileNameElementId).textContent = file.name;

    try {
        const data = await readExcelFile(file);
        
        switch(type) {
            case 'source':
                sourceData = data;
                break;
            case 'compare1':
                compareData1 = data;
                break;
            case 'compare2':
                compareData2 = data;
                break;
        }

        checkAllFilesLoaded();
    } catch (error) {
        alert('Error reading file: ' + error.message);
    }
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                
                // Get the first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = function() {
            reject(new Error('Failed to read file'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

function checkAllFilesLoaded() {
    const compareBtn = document.getElementById('compareBtn');
    
    // Enable button if we have source file and at least one comparison file
    if (sourceData && (compareData1 || compareData2)) {
        compareBtn.disabled = false;
        
        // Update button text based on which files are loaded
        if (compareData1 && compareData2) {
            compareBtn.textContent = 'Compare with Both Wallets';
        } else if (compareData1) {
            compareBtn.textContent = 'Compare with Wallet 1';
        } else if (compareData2) {
            compareBtn.textContent = 'Compare with Wallet 2';
        }
    } else {
        compareBtn.disabled = true;
        compareBtn.textContent = 'Upload Bank Sheet + At Least 1 Wallet';
    }
}

async function compareFiles() {
    if (!sourceData) {
        alert('Please upload the Bank Sheet file first.');
        return;
    }
    
    if (!compareData1 && !compareData2) {
        alert('Please upload at least one Balad Wallet file for comparison.');
        return;
    }

    document.getElementById('loading').style.display = 'block';
    document.getElementById('results').style.display = 'none';

    // Simulate processing time for better UX
    await new Promise(resolve => setTimeout(resolve, 500));

    try {
        comparisonResults = performComparison(sourceData, compareData1, compareData2);
        displayResults(comparisonResults);
        
        document.getElementById('loading').style.display = 'none';
        document.getElementById('results').style.display = 'block';
    } catch (error) {
        document.getElementById('loading').style.display = 'none';
        alert('Error during comparison: ' + error.message);
    }
}

// REMOVE this duplicate and incomplete performComparison function

function displayResults(results) {
    // Display statistics
    const statsGrid = document.getElementById('statsGrid');
    
    // Check if description columns were found
    const descStatus = results.statistics.descriptionColumnFound;
    const filesCompared = results.statistics.filesCompared;
    
    let statusHtml = `
        <div class="stat-card ${descStatus.source ? 'success' : 'error'}">
            <span class="stat-number">${descStatus.source ? '✓' : '✗'}</span>
            Bank Sheet Column
        </div>
    `;
    
    if (filesCompared.wallet1) {
        statusHtml += `
            <div class="stat-card ${descStatus.compare1 ? 'success' : 'error'}">
                <span class="stat-number">${descStatus.compare1 ? '✓' : '✗'}</span>
                Balad Wallet 1 Column
            </div>
        `;
    }
    
    if (filesCompared.wallet2) {
        statusHtml += `
            <div class="stat-card ${descStatus.compare2 ? 'success' : 'error'}">
                <span class="stat-number">${descStatus.compare2 ? '✓' : '✗'}</span>
                Balad Wallet 2 Column
            </div>
        `;
    }
    
    statsGrid.innerHTML = `
        <div class="stat-card">
            <span class="stat-number">${results.statistics.sourceRows}</span>
            Bank Sheet Rows
        </div>
        <div class="stat-card">
            <span class="stat-number">${results.statistics.matchingRows}</span>
            Total Matches
        </div>
        <div class="stat-card tp2p-matches">
            <span class="stat-number">${results.statistics.tp2pMatches || 0}</span>
            TP2P Matches
        </div>
        <div class="stat-card exact-matches">
            <span class="stat-number">${results.statistics.exactMatches || 0}</span>
            Exact Matches
        </div>
        <div class="stat-card">
            <span class="stat-number">${results.statistics.uniqueRows}</span>
            Unique Descriptions
        </div>
        <div class="stat-card">
            <span class="stat-number">${results.statistics.sourceRows > 0 ? Math.round((results.statistics.matchingRows / results.statistics.sourceRows) * 100) : 0}%</span>
            Match Rate
        </div>
        ${statusHtml}
    `;

    // Display matching descriptions
    const matchingResults = document.getElementById('matchingResults');
    if (results.matches.length > 0) {
        // Build table header dynamically based on which files are compared
        let headerHtml = `
            <tr>
                <th>Row #</th>
                <th>Description</th>
                <th>Debit</th>
                <th>Credit</th>
                <th>Match Type</th>
        `;
        
        if (filesCompared.wallet1) {
            headerHtml += `<th>Found in Balad Wallet 1</th>`;
        }
        if (filesCompared.wallet2) {
            headerHtml += `<th>Found in Balad Wallet 2</th>`;
        }
        
        headerHtml += `</tr>`;
        
        let tableHtml = `
            <table class="result-table">
                <thead>
                    ${headerHtml}
                </thead>
                <tbody>
        `;
        
        results.matches.forEach(match => {
            // Determine overall match type
            let overallMatchType = 'Exact';
            if (match.matchType1 === 'TP2P' || match.matchType2 === 'TP2P') {
                overallMatchType = 'TP2P';
            }
            
            // Extract debit and credit from the row data
            const debit = match.debit || '';
            const credit = match.credit || '';
            
            let rowHtml = `
                <tr>
                    <td>${match.rowIndex}</td>
                    <td><strong>${match.description || 'N/A'}</strong></td>
                    <td>${debit}</td>
                    <td>${credit}</td>
                    <td class="${overallMatchType === 'TP2P' ? 'tp2p-match' : 'exact-match'}">${overallMatchType}</td>
            `;
            
            // Format status for Balad Wallet 1 (only if file was uploaded)
            if (filesCompared.wallet1) {
                let wallet1Status = '';
                if (match.foundInFile1) {
                    if (match.matchType1 === 'TP2P') {
                        wallet1Status = '<span class="tp2p-status">Canceled and Credited to Balance</span>';
                    } else {
                        wallet1Status = '<span class="exact-status">✓ Exact</span>';
                    }
                } else {
                    wallet1Status = '<span class="not-found-status">✗</span>';
                }
                rowHtml += `<td>${wallet1Status}</td>`;
            }
            
            // Format status for Balad Wallet 2 (only if file was uploaded)
            if (filesCompared.wallet2) {
                let wallet2Status = '';
                if (match.foundInFile2) {
                    if (match.matchType2 === 'TP2P') {
                        wallet2Status = '<span class="tp2p-status">Canceled and Credited to Balance</span>';
                    } else {
                        wallet2Status = '<span class="exact-status">✓ Exact</span>';
                    }
                } else {
                    wallet2Status = '<span class="not-found-status">✗</span>';
                }
                rowHtml += `<td>${wallet2Status}</td>`;
            }
            
            rowHtml += `</tr>`;
            tableHtml += rowHtml;
        });
        
        tableHtml += '</tbody></table>';
        tableHtml += `<p><strong>Total matches: ${results.matches.length}</strong></p>`;
        
        matchingResults.innerHTML = tableHtml;
    } else {
        matchingResults.innerHTML = '<p>No matching descriptions found between the files.</p>';
    }

    // Display unique descriptions
    const uniqueResults = document.getElementById('uniqueResults');
    if (results.unique.length > 0) {
        let tableHtml = `
            <table class="result-table">
                <thead>
                    <tr>
                        <th>Row #</th>
                        <th>Description</th>
                        <th>Debit</th>
                        <th>Credit</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        results.unique.forEach(unique => {
            const debit = unique.debit || '';
            const credit = unique.credit || '';
            
            tableHtml += `
                <tr>
                    <td>${unique.rowIndex}</td>
                    <td><strong>${unique.description || 'N/A'}</strong></td>
                    <td>${debit}</td>
                    <td>${credit}</td>
                </tr>
            `;
        });
        
        tableHtml += '</tbody></table>';
        tableHtml += `<p><strong>Total unique descriptions: ${results.unique.length}</strong></p>`;
        
        uniqueResults.innerHTML = tableHtml;
    } else {
        uniqueResults.innerHTML = '<p>No unique descriptions found. All bank descriptions have matches in Balad wallet files.</p>';
    }
}
/* The following duplicate and invalid code block is removed to fix syntax errors. */

function exportResults() {
    if (!comparisonResults) {
        alert('No results to export. Please run a comparison first.');
        return;
    }

    const file1Name = document.getElementById('compareFile1Name').textContent || 'File 1';
    const file2Name = document.getElementById('compareFile2Name').textContent || 'File 2';

    // Generate CSV content
    const csvContent = generateCSV(comparisonResults, file1Name, file2Name);

    // Create a blob and download as .csv (Excel compatible)
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'BankComparisonResults.csv'; // File name
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}


function generateCSV(results, file1Name = 'File 1', file2Name = 'File 2') {
    let csv = [
        'Type',
        'Row Number',
        'Description',
        'Debit',
        'Credit',
        'Balance',
        `Found in ${file1Name}`,
        `Found in ${file2Name}`,
        'Match Details'
    ].join(',') + '\n';

    // Matches
    results.matches.forEach(match => {
        const description = match.description || 'N/A';
        const debit = match.debit || '';
        const credit = match.credit || '';
        const balance = match.data[match.data.length - 1] || '';

        csv += [
            'Match',
            match.rowIndex,
            `"${description}"`,
            `"${debit}"`,
            `"${credit}"`,
            `"${balance}"`,
            match.foundInFile1,
            match.foundInFile2,
            '' // Empty for matches
        ].join(',') + '\n';
    });

    // Unique
    results.unique.forEach(unique => {
        const description = unique.description || 'N/A';
        const debit = unique.debit || '';
        const credit = unique.credit || '';
        const balance = unique.data[unique.data.length - 1] || '';

        csv += [
            'Unique',
            unique.rowIndex,
            `"${description}"`,
            `"${debit}"`,
            `"${credit}"`,
            `"${balance}"`,
            '', '', '"No matches found"'
        ].join(',') + '\n';
    });

    return csv;
}
