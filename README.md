
# UPI Search Automation Tool

This tool helps financial institutions search for the appropriate Unique Product Identifier (UPI) for their OTC derivatives trades, supporting the HKTR reform requirements effective September 2025.

## Features

- Supports both FX (Foreign Exchange) and IR (Interest Rate) derivatives
- Allows loading of UPI data from DSB in JSON format
- Loads trade details from Excel files
- Provides intuitive column mapping for different bank data formats
- Automatically suggests column mappings based on common naming conventions
- Performs intelligent UPI matching using a scoring system
- Exports results to Excel for further analysis or integration

## Requirements

- Python 3.7 or higher
- Required Python packages:
  - pandas
  - openpyxl
  - tkinter (for GUI version)

## Installation

1. Ensure you have Python installed on your system
2. Install required packages:
   ```
   pip install pandas openpyxl
   ```
3. Download the tool files:
   - `upi_search_tool.py` - GUI version
   - `upi_search_batch.py` - Command-line batch processing version
   - `upi_search_test_cases.py` - Test cases

## Usage - GUI Version

Run the GUI version for interactive use:

```
python upi_search_tool.py
```

Steps:
1. In the "Upload Files" tab, browse and select your UPI JSON file and trade Excel file
2. Select the appropriate asset class (FX or IR)
3. Click "Load Data" to load the files
4. In the "Map Columns" tab, verify or adjust the automatic column mapping
5. Click "Map Columns & Search UPIs" to start the search process
6. View the results in the "Results" tab
7. Export the results to Excel using the "Export Results to Excel" button

## Usage - Batch Processing

For batch processing or integration into other systems:

```
python upi_search_batch.py --upi <upi_file.json> --trade <trade_file.xlsx> --asset-class <FX|IR> --output <output_file.xlsx>
```

Arguments:
- `--upi`: Path to the UPI JSON file (required)
- `--trade`: Path to the trade Excel file (required)
- `--asset-class`: Asset class, either "FX" or "IR" (default: "FX")
- `--output`: Path to output Excel file (default: results_YYYY-MM-DD.xlsx)

## Testing

To run the included test cases with sample data:

```
python upi_search_test_cases.py
```

## UPI Data Format

The tool expects UPI data in JSON format with the following structure:

```json
{
  "upis": [
    {
      "upiCode": "ABCDEFGHIJKL",
      "assetClass": "Rates",
      "instrumentType": "Swap",
      "product": "Fixed_Float",
      "underlying": {
        "referenceRate": "USD-LIBOR-3M",
        "currency": "USD",
        "term": "3M"
      },
      "fixedRate": {
        "currency": "USD",
        "term": "6M"
      },
      "deliveryType": "Cash"
    }
  ]
}
```

## Trade Data Format

The tool can work with various Excel formats for trade data. You will be able to map your specific columns to the required UPI search attributes through the interface.

## Notes

- This tool provides matching suggestions based on a scoring system. Always verify the suggested UPIs before using them for regulatory reporting.
- Recommended to test with a small subset of trades before processing your entire portfolio.
- The scoring threshold is set to 50% by default. You can adjust this in the code if needed.

## Support

For questions or support, please contact your technical team or system administrator.
