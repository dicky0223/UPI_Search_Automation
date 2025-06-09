# UPI Search Automation Tool - Complete User Guide

## Overview

The UPI Search Automation Tool is designed to help financial institutions automatically find appropriate Unique Product Identifiers (UPIs) for their OTC derivatives trades, supporting the Hong Kong Trade Repository (HKTR) reform requirements effective September 2025.

## Tool Components

### 1. GUI Version (`upi_search_tool.py`)
- Interactive graphical interface
- Best for manual operation and testing
- Supports file browsing and visual column mapping
- Real-time results display

### 2. Batch Processing Version (`upi_search_batch.py`)
- Command-line interface
- Best for automated processing and integration
- Supports bulk processing of large datasets
- Can be scheduled or called from other systems

### 3. Test Cases (`upi_search_test_cases.py`)
- Validates tool functionality
- Demonstrates expected behavior
- Useful for troubleshooting

## Step-by-Step Usage Guide

### Step 1: Prepare Your Data

#### UPI Data (JSON Format)
Download the latest UPI reference data from the Derivatives Service Bureau (DSB) in JSON format. The expected structure is:

```json
{
  "upis": [
    {
      "upiCode": "UNIQUE12CHAR",
      "assetClass": "ForeignExchange" or "Rates",
      "instrumentType": "Forward", "Swap", "Option",
      "product": "Vanilla", "NDF", "Fixed_Float", etc.,
      "underlying": {
        "currencyPair": "USD/CNY",
        "referenceRate": "USD-LIBOR-3M",
        "currency": "USD",
        "term": "3M"
      },
      "deliveryType": "Cash" or "Physical"
    }
  ]
}
```

#### Trade Data (Excel Format)
Prepare your trade data in Excel format with columns containing relevant trade attributes. Common column names include:
- TradeID, Asset Class, Product Type, Currency Pair
- Reference Rate, Term, Settlement Type
- Option Type, Option Style (for options)

### Step 2: Using the GUI Version

1. **Launch the Tool**
   ```bash
   python upi_search_tool.py
   ```

2. **Upload Files Tab**
   - Browse and select your UPI JSON file
   - Browse and select your trade Excel file
   - Select asset class (FX or IR)
   - Click "Load Data"

3. **Map Columns Tab**
   - The tool will automatically suggest column mappings
   - Review and adjust mappings as needed
   - Mark irrelevant attributes as "N/A"
   - Click "Map Columns & Search UPIs"

4. **Results Tab**
   - View detailed search results
   - See match scores and explanations
   - Export results to Excel for further analysis

### Step 3: Using the Batch Processing Version

For automated or large-scale processing:

```bash
python upi_search_batch.py \
  --upi dsb_upi_data.json \
  --trade my_trades.xlsx \
  --asset-class FX \
  --output upi_search_results.xlsx
```

## Column Mapping Guide

### FX (Foreign Exchange) Attributes
| UPI Attribute | Common Column Names | Required |
|---------------|-------------------|----------|
| Asset Class | AssetClass, Asset_Class, Product_Class | Yes |
| Instrument Type | InstrumentType, ProductType, TradeType | Yes |
| Product Type | Product, ForwardType, OptionType | Yes |
| Currency Pair | CcyPair, CurrencyPair, Ccy_Pair | Yes |
| Settlement Currency | SettlementCcy, Settlement_Currency | Optional |
| Option Type | OptionType, CallPut | Optional |
| Option Style | OptionStyle, Exercise_Style | Optional |
| Delivery Type | DeliveryType, SettlementType | Optional |

### IR (Interest Rate) Attributes
| UPI Attribute | Common Column Names | Required |
|---------------|-------------------|----------|
| Asset Class | AssetClass, Asset_Class | Yes |
| Instrument Type | InstrumentType, ProductType | Yes |
| Product Type | Product, SwapType, Swap_Type | Yes |
| Reference Rate | RefRate, ReferenceRate, IndexRate | Yes |
| Currency | Currency, Ccy, Currency1 | Yes |
| Term | Term, Tenor, Term1 | Yes |
| Other Leg Reference Rate | RefRate2, OtherLegRate | Optional |
| Other Leg Currency | Currency2, OtherLegCcy | Optional |
| Other Leg Term | Term2, OtherLegTerm | Optional |
| Delivery Type | DeliveryType, SettlementType | Optional |

## Scoring System

The tool uses a weighted scoring system to match trades with UPIs:

### FX Scoring Weights
- Asset Class Match: 20 points
- Instrument Type Match: 20 points
- Product Match: 20 points
- Currency Pair Match: 20 points
- Settlement Currency Match: 10 points
- Option Type/Style Match: 5 points each
- Delivery Type Match: 10 points

### IR Scoring Weights
- Asset Class Match: 20 points
- Instrument Type Match: 20 points
- Product Match: 20 points
- Reference Rate Match: 15 points
- Currency Match: 10 points
- Term Match: 10 points
- Other Leg Attributes: 5-10 points each

**Minimum Score Threshold: 50 points** (adjustable in code)

## Best Practices

### Data Preparation
1. Ensure your trade data is clean and consistent
2. Use standardized codes where possible (e.g., ISO currency codes)
3. Remove or handle blank/null values appropriately

### Column Mapping
1. Review auto-suggested mappings carefully
2. Use "N/A" for truly irrelevant attributes
3. Test with a small sample before processing large datasets

### Result Validation
1. Review matches with scores below 70 manually
2. Validate high-confidence matches (score > 90) with sample checks
3. Export results for audit trails

### Performance Optimization
1. Process large datasets in batches
2. Use the batch processing version for automated workflows
3. Consider parallel processing for very large datasets

## Troubleshooting

### Common Issues

**Issue: No matches found**
- Check data formats and column mappings
- Verify UPI data contains relevant asset classes
- Review scoring threshold settings

**Issue: Low match scores**
- Inconsistent data formats between trade and UPI data
- Missing key attributes in trade data
- Need to adjust mapping or data standardization

**Issue: File loading errors**
- Verify file formats (JSON for UPI, Excel for trades)
- Check file permissions and paths
- Ensure required Python packages are installed

### Error Messages

| Error | Cause | Solution |
|-------|-------|----------|
| "No UPI records found" | Empty or malformed UPI JSON | Verify UPI file format and content |
| "Error loading files" | File access or format issues | Check file paths and permissions |
| "No matching UPI found with sufficient confidence" | Low match scores | Review mappings and data quality |

## Integration with Trading Systems

### API Integration
The batch processing version can be easily integrated into existing workflows:

```python
# Example integration
import subprocess

def search_upis_for_trades(upi_file, trade_file, asset_class):
    cmd = [
        'python', 'upi_search_batch.py',
        '--upi', upi_file,
        '--trade', trade_file,
        '--asset-class', asset_class,
        '--output', 'results.xlsx'
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode == 0
```

### Scheduled Processing
Set up automated UPI searches using cron (Linux/Mac) or Task Scheduler (Windows):

```bash
# Daily UPI search at 9 AM
0 9 * * * /usr/bin/python3 /path/to/upi_search_batch.py --upi latest_upi.json --trade daily_trades.xlsx --asset-class FX
```

## Regulatory Compliance Notes

- This tool provides matching suggestions based on available data
- Always verify UPI matches before regulatory reporting
- Maintain audit trails of UPI assignments
- Review and update UPI reference data regularly from DSB
- Consider implementing additional validation rules for your specific use cases

## Support and Maintenance

### Regular Updates
- Update UPI reference data from DSB regularly
- Review and update column mapping logic as needed
- Monitor match accuracy and adjust scoring weights if necessary

### Backup and Recovery
- Maintain backups of configuration and mapping settings
- Document custom modifications for future reference
- Test tool functionality after system updates

For technical support, consult your system administrator or development team.