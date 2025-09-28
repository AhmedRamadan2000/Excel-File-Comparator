function findDescriptionColumn(data) {
    // Look through multiple rows to find the header row with "description"
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

function findCreditColumn(headerRow) {
    for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i] && 
            headerRow[i].toString().toLowerCase().trim() === 'credit') {
            return i;
        }
    }
    return -1;
}

function findDebitColumn(headerRow) {
    for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i] && 
            headerRow[i].toString().toLowerCase().trim() === 'debit') {
            return i;
        }
    }
    return -1;
}

function isTP2PMatch(sourceDesc, compareDesc, compareRow, creditColumnIndex) {
    // Check if source description ends with TP2P
    if (!sourceDesc.toString().toUpperCase().endsWith('TP2P')) {
        return false;
    }
    
    // Extract the base code without TP2P
    const baseCode = sourceDesc.toString().slice(0, -4); // Remove last 4 characters (TP2P)
    
    // Check if compare description matches the base code
    const compareDescMatch = compareDesc.toString().trim() === baseCode.trim();
    
    // Check if the amount is in credit column (not empty and not zero)
    let isCreditEntry = false;
    if (creditColumnIndex !== -1 && compareRow[creditColumnIndex]) {
        const creditValue = parseFloat(compareRow[creditColumnIndex].toString().replace(/,/g, ''));
        isCreditEntry = !isNaN(creditValue) && creditValue > 0;
    }
    
    return compareDescMatch && isCreditEntry;
}

function performComparison(source, compare1, compare2) {
    // Find description column info for each file
    const sourceDescInfo = findDescriptionColumn(source);
    const compare1DescInfo = findDescriptionColumn(compare1);
    const compare2DescInfo = findDescriptionColumn(compare2);

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
                source: sourceDescInfo !== null,
                compare1: compare1DescInfo !== null,
                compare2: compare2DescInfo !== null
            }
        },
        matches: [],
        unique: []
    };

    // If source doesn't have description column, return error
    if (!sourceDescInfo) {
        throw new Error('Description column not found in source file. Please ensure your source file has a column named "Description" (case insensitive).');
    }

    // Get data rows after the header row for each file
    const sourceDataRows = source.slice(sourceDescInfo.headerRow + 1);
    const compare1DataRows = compare1DescInfo ? compare1.slice(compare1DescInfo.headerRow + 1) : [];
    const compare2DataRows = compare2DescInfo ? compare2.slice(compare2DescInfo.headerRow + 1) : [];

    // Update row counts
    results.statistics.sourceRows = sourceDataRows.length;
    results.statistics.compare1Rows = compare1DataRows.length;
    results.statistics.compare2Rows = compare2DataRows.length;

    // Compare each source row based on description column
    for (let i = 0; i < sourceDataRows.length; i++) {
        const sourceRow = sourceDataRows[i];
        const sourceDescription = sourceRow[sourceDescInfo.columnIndex];
        
        // Skip empty descriptions or empty rows
        if (!sourceDescription || sourceDescription.toString().trim() === '') {
            continue;
        }
        
        let foundInCompare1 = false;
        let foundInCompare2 = false;
        let matchType1 = '';
        let matchType2 = '';
        let matchingRows = {
            file1: [],
            file2: []
        };
        
        // Check if description exists in compare1
        if (compare1DescInfo) {
            for (let j = 0; j < compare1DataRows.length; j++) {
                const compare1Description = compare1DataRows[j][compare1DescInfo.columnIndex];
                if (compare1Description) {
                    // First try exact match
                    if (sourceDescription.toString().toLowerCase().trim() === 
                        compare1Description.toString().toLowerCase().trim()) {
                        foundInCompare1 = true;
                        matchType1 = 'exact';
                        matchingRows.file1.push({
                            rowIndex: j + compare1DescInfo.headerRow + 2,
                            data: compare1DataRows[j],
                            matchType: 'exact'
                        });
                    }
                    // Then try TP2P match
                    else if (isTP2PMatch(sourceDescription, compare1Description, 
                             compare1DataRows[j], compare1DescInfo.creditColumn)) {
                        foundInCompare1 = true;
                        matchType1 = 'TP2P';
                        matchingRows.file1.push({
                            rowIndex: j + compare1DescInfo.headerRow + 2,
                            data: compare1DataRows[j],
                            matchType: 'TP2P'
                        });
                    }
                }
            }
        }
        
        // Check if description exists in compare2
        if (compare2DescInfo) {
            for (let k = 0; k < compare2DataRows.length; k++) {
                const compare2Description = compare2DataRows[k][compare2DescInfo.columnIndex];
                if (compare2Description) {
                    // First try exact match
                    if (sourceDescription.toString().toLowerCase().trim() === 
                        compare2Description.toString().toLowerCase().trim()) {
                        foundInCompare2 = true;
                        matchType2 = 'exact';
                        matchingRows.file2.push({
                            rowIndex: k + compare2DescInfo.headerRow + 2,
                            data: compare2DataRows[k],
                            matchType: 'exact'
                        });
                    }
                    // Then try TP2P match
                    else if (isTP2PMatch(sourceDescription, compare2Description, 
                             compare2DataRows[k], compare2DescInfo.creditColumn)) {
                        foundInCompare2 = true;
                        matchType2 = 'TP2P';
                        matchingRows.file2.push({
                            rowIndex: k + compare2DescInfo.headerRow + 2,
                            data: compare2DataRows[k],
                            matchType: 'TP2P'
                        });
                    }
                }
            }
        }
        
        if (foundInCompare1 || foundInCompare2) {
            results.matches.push({
                rowIndex: i + sourceDescInfo.headerRow + 2,
                data: sourceRow,
                description: sourceDescription,
                foundInFile1: foundInCompare1,
                foundInFile2: foundInCompare2,
                matchType1: matchType1,
                matchType2: matchType2,
                matchingRows: matchingRows
            });
            results.statistics.matchingRows++;
            
            // Count match types
            if (matchType1 === 'TP2P' || matchType2 === 'TP2P') {
                results.statistics.tp2pMatches++;
            } else {
                results.statistics.exactMatches++;
            }
        } else {
            results.unique.push({
                rowIndex: i + sourceDescInfo.headerRow + 2,
                data: sourceRow,
                description: sourceDescription
            });
            results.statistics.uniqueRows++;
        }
    }

    return results;
}let sourceData = null;
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
    if (sourceData && compareData1 && compareData2) {
        compareBtn.disabled = false;
        compareBtn.textContent = 'Start Comparison';
    }
}

async function compareFiles() {
    if (!sourceData || !compareData1 || !compareData2) {
        alert('Please upload all three files first.');
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

/* Duplicate performComparison function removed to fix syntax error */

function displayResults(results) {
    // Display statistics
    const statsGrid = document.getElementById('statsGrid');
    
    // Check if description columns were found
    const descStatus = results.statistics.descriptionColumnFound;
    const statusHtml = `
        <div class="stat-card ${descStatus.source ? 'success' : 'error'}">
            <span class="stat-number">${descStatus.source ? '✓' : '✗'}</span>
            Source Desc Column
        </div>
        <div class="stat-card ${descStatus.compare1 ? 'success' : 'error'}">
            <span class="stat-number">${descStatus.compare1 ? '✓' : '✗'}</span>
            File 1 Desc Column
        </div>
        <div class="stat-card ${descStatus.compare2 ? 'success' : 'error'}">
            <span class="stat-number">${descStatus.compare2 ? '✓' : '✗'}</span>
            File 2 Desc Column
        </div>
    `;
    
    statsGrid.innerHTML = `
        <div class="stat-card">
            <span class="stat-number">${results.statistics.sourceRows}</span>
            Source Rows
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
        let tableHtml = `
            <table class="result-table">
                <thead>
                    <tr>
                        <th>Row #</th>
                        <th>Description</th>
                        <th>Match Type</th>
                        <th>Found in File 1</th>
                        <th>Found in File 2</th>
                        <th>Match Details</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        results.matches.slice(0, 100).forEach(match => {
            const matchDetails = [];
            if (match.foundInFile1) {
                matchDetails.push(`File 1: ${match.matchingRows.file1.length} ${match.matchType1} matches`);
            }
            if (match.foundInFile2) {
                matchDetails.push(`File 2: ${match.matchingRows.file2.length} ${match.matchType2} matches`);
            }
            
            // Determine overall match type
            let overallMatchType = 'Exact';
            if (match.matchType1 === 'TP2P' || match.matchType2 === 'TP2P') {
                overallMatchType = 'TP2P';
            }
            
            tableHtml += `
                <tr>
                    <td>${match.rowIndex}</td>
                    <td><strong>${match.description || 'N/A'}</strong></td>
                    <td class="${overallMatchType === 'TP2P' ? 'tp2p-match' : 'exact-match'}">${overallMatchType}</td>
                    <td class="${match.foundInFile1 ? 'success' : 'error'}">${match.foundInFile1 ? '✓ ' + match.matchType1 : '✗'}</td>
                    <td class="${match.foundInFile2 ? 'success' : 'error'}">${match.foundInFile2 ? '✓ ' + match.matchType2 : '✗'}</td>
                    <td>${matchDetails.join(', ')}</td>
                </tr>
            `;
        });
        
        tableHtml += '</tbody></table>';
        
        if (results.matches.length > 100) {
            tableHtml += `<p><em>Showing first 100 results. Total matches: ${results.matches.length}</em></p>`;
        }
        
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
                        <th>Unique Description</th>
                        <th>Full Row Data</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        results.unique.slice(0, 100).forEach(unique => {
            tableHtml += `
                <tr>
                    <td>${unique.rowIndex}</td>
                    <td><strong>${unique.description || 'N/A'}</strong></td>
                    <td>${JSON.stringify(unique.data)}</td>
                </tr>
            `;
        });
        
        tableHtml += '</tbody></table>';
        
        if (results.unique.length > 100) {
            tableHtml += `<p><em>Showing first 100 results. Total unique descriptions: ${results.unique.length}</em></p>`;
        }
        
        uniqueResults.innerHTML = tableHtml;
    } else {
        uniqueResults.innerHTML = '<p>No unique descriptions found. All source descriptions have matches in comparison files.</p>';
    }
}

function exportResults(format) {
    if (!comparisonResults) {
        alert('No results to export. Please run a comparison first.');
        return;
    }

    let content, filename, mimeType;

    if (format === 'csv') {
        content = generateCSV(comparisonResults);
        filename = 'excel_comparison_results.csv';
        mimeType = 'text/csv';
    } else if (format === 'json') {
        content = JSON.stringify(comparisonResults, null, 2);
        filename = 'excel_comparison_results.json';
        mimeType = 'application/json';
    }

    const blob = new Blob([content], { type: mimeType });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

function generateCSV(results) {
    let csv = 'Type,Row Number,Description,Full Row Data,Found in File 1,Found in File 2,Match Details\n';
    
    results.matches.forEach(match => {
        const data = JSON.stringify(match.data).replace(/"/g, '""');
        const description = match.description ? match.description.toString().replace(/"/g, '""') : 'N/A';
        const matchDetails = [];
        
        if (match.foundInFile1) {
            matchDetails.push(`File1:${match.matchingRows.file1.length}matches`);
        }
        if (match.foundInFile2) {
            matchDetails.push(`File2:${match.matchingRows.file2.length}matches`);
        }
        
        csv += `Match,${match.rowIndex},"${description}","${data}",${match.foundInFile1},${match.foundInFile2},"${matchDetails.join(', ')}"\n`;
    });
    
    results.unique.forEach(unique => {
        const data = JSON.stringify(unique.data).replace(/"/g, '""');
        const description = unique.description ? unique.description.toString().replace(/"/g, '""') : 'N/A';
        csv += `Unique,${unique.rowIndex},"${description}","${data}",,,"No matches found"\n`;
    });
    
    return csv;
}