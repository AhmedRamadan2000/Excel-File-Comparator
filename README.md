# Excel File Comparator

A modern, web-based tool to compare Excel files and identify matching and unique rows across multiple datasets.

## Features

- **Multi-file Comparison**: Compare one source Excel file against two comparison files
- **Real-time Processing**: Fast, client-side processing with visual feedback
- **Detailed Results**: View statistics, matching rows, and unique rows
- **Export Functionality**: Export results as CSV or JSON files
- **Responsive Design**: Works on desktop, tablet, and mobile devices
- **Modern UI**: Clean, professional interface with animations and gradients

## Supported File Formats

- Excel files (.xlsx, .xls)
- CSV files (.csv)

## How to Use

### Method 1: Direct File Opening
1. Download all project files to a folder
2. Open `index.html` in your web browser
3. Upload your three Excel files (1 source + 2 comparison files)
4. Click "Start Comparison" to analyze the data
5. View results and export if needed

### Method 2: Using Live Server (Recommended for Development)
1. Install Node.js from [nodejs.org](https://nodejs.org/)
2. Navigate to the project folder in terminal/command prompt
3. Run: `npm install`
4. Run: `npm start`
5. Open your browser to `http://localhost:3000`

## Project Structure

```
excel-comparator/
├── index.html          # Main HTML structure
├── style.css           # All styling and animations
├── script.js           # JavaScript functionality
├── package.json        # Project configuration
└── README.md          # This file
```

## How It Works

1. **File Upload**: Uses HTML5 File API to read Excel files
2. **Data Processing**: SheetJS library converts Excel data to JSON format
3. **Comparison Logic**: Compares each row in the source file against all rows in comparison files
4. **Results Display**: Shows statistics, matches, and unique rows in organized tables
5. **Export**: Generates downloadable CSV or JSON files with results

## Comparison Logic

- **Exact Matching**: Rows are compared using exact string matching (JSON.stringify)
- **Row-by-Row**: Each source row is checked against all rows in comparison files
- **Statistics**: Calculates match rates and provides detailed breakdowns
- **Categorization**: Results are split into "Matching Rows" and "Unique Rows"

## Browser Compatibility

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## Dependencies

- **SheetJS**: For Excel file processing (loaded via CDN)
- No additional runtime dependencies

## Development

To contribute or modify this project:

1. Clone the repository
2. Make your changes to HTML, CSS, or JavaScript files
3. Test by opening `index.html` in a browser
4. For development server: run `npm start`

## Performance Considerations

- **Client-side Processing**: All processing happens in the browser
- **Memory Usage**: Large files (>10MB) may cause performance issues
- **Display Limits**: Only first 100 results shown in tables (full data available in exports)

## Future Enhancements

- [ ] Column-specific comparison options
- [ ] Fuzzy matching for similar (not exact) rows
- [ ] Multiple sheet support
- [ ] Database integration
- [ ] Progress bars for large files
- [ ] Advanced filtering options
- [ ] Bulk file processing

## Troubleshooting

**Files won't upload:**
- Ensure files are in supported formats (.xlsx, .xls, .csv)
- Check file size (very large files may cause issues)
- Try refreshing the page

**Comparison not working:**
- Make sure all three files are uploaded
- Check that files contain data (not empty)
- Verify files have readable content

**Results seem incorrect:**
- Remember that comparison uses exact matching
- Check for extra spaces or formatting differences in your data
- Export results to verify the raw comparison data

## License

MIT License - feel free to use and modify as needed.

## Support

For issues or questions, please check the troubleshooting section above or create an issue in the project repository.
