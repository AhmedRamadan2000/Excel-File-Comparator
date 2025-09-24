function findDescriptionColumn(data) {
    // Look through multiple rows to find the header row with "description"
    for (let rowIndex = 0; rowIndex < Math.min(10, data.length); rowIndex++) {
        const row = data[rowIndex] || [];
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
            if (row[colIndex] && 
                row[colIndex].toString().toLowerCase().trim() === 'description') {
                return {
                    headerRow: rowIndex,
                    columnIndex: colIndex
                };
            }
        }
    }
    return null;
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
        throw new Error('Description column not found in Bank File. Please ensure your Bank File has a column named "Description" (case insensitive).');
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
        let matchingRows = {
            file1: [],
            file2: []
        };
        
        // Check if description exists in compare1
        if (compare1DescInfo) {
            for (let j = 0; j < compare1DataRows.length; j++) {
                const compare1Description = compare1DataRows[j][compare1DescInfo.columnIndex];
                if (compare1Description && 
                    sourceDescription.toString().toLowerCase().trim() === 
                    compare1Description.toString().toLowerCase().trim()) {
                    foundInCompare1 = true;
                    matchingRows.file1.push({
                        rowIndex: j + compare1DescInfo.headerRow + 2, // Actual Excel row number
                        data: compare1DataRows[j]
                    });
                }
            }
        }
        
        // Check if description exists in compare2
        if (compare2DescInfo) {
            for (let k = 0; k < compare2DataRows.length; k++) {
                const compare2Description = compare2DataRows[k][compare2DescInfo.columnIndex];
                if (compare2Description && 
                    sourceDescription.toString().toLowerCase().trim() === 
                    compare2Description.toString().toLowerCase().trim()) {
                    foundInCompare2 = true;
                    matchingRows.file2.push({
                        rowIndex: k + compare2DescInfo.headerRow + 2, // Actual Excel row number
                        data: compare2DataRows[k]
                    });
                }
            }
        }
        
        if (foundInCompare1 || foundInCompare2) {
            results.matches.push({
                rowIndex: i + sourceDescInfo.headerRow + 2, // Actual Excel row number
                data: sourceRow,
                description: sourceDescription,
                foundInFile1: foundInCompare1,
                foundInFile2: foundInCompare2,
                matchingRows: matchingRows
            });
            results.statistics.matchingRows++;
        } else {
            results.unique.push({
                rowIndex: i + sourceDescInfo.headerRow + 2, // Actual Excel row number
                data: sourceRow,
                description: sourceDescription
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

// Removed duplicate and incomplete performComparison function

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
            BDC EGP Desc Column
        </div>
        <div class="stat-card ${descStatus.compare2 ? 'success' : 'error'}">
            <span class="stat-number">${descStatus.compare2 ? '✓' : '✗'}</span>
            BDC EGP - Euro based Desc Column
        </div>
    `;
    
    statsGrid.innerHTML = `
        <div class="stat-card">
            <span class="stat-number">${results.statistics.sourceRows}</span>
            Source Rows
        </div>
        <div class="stat-card">
            <span class="stat-number">${results.statistics.matchingRows}</span>
            Matching Descriptions
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
                        <th>Full Row Data</th>
                        <th>Found in BDC EGP</th>
                        <th>Found in BDC EGP - Euro based</th>
                        <th>Matching Details</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        results.matches.slice(0, 100).forEach(match => {
            const matchDetails = [];
            if (match.foundInFile1) {
                matchDetails.push(`BDC EGP: ${match.matchingRows.file1.length} matches`);
            }
            if (match.foundInFile2) {
                matchDetails.push(`BDC EGP - Euro based: ${match.matchingRows.file2.length} matches`);
            }
            
            tableHtml += `
                <tr>
                    <td>${match.rowIndex}</td>
                    <td><strong>${match.description || 'N/A'}</strong></td>
                    <td>${JSON.stringify(match.data).substring(0, 100)}...</td>
                    <td class="${match.foundInFile1 ? 'success' : 'error'}">${match.foundInFile1 ? '✓' : '✗'}</td>
                    <td class="${match.foundInFile2 ? 'success' : 'error'}">${match.foundInFile2 ? '✓' : '✗'}</td>
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
    let csv = 'Type,Row Number,Description,Full Row Data,Found in BDC EGP,Found in BDC EGP - Euro based,Match Details\n';
    
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