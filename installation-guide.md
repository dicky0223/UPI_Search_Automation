# Installation Script for UPI Search Automation Tool

## Quick Installation Guide

This script will help you set up the UPI Search Automation Tool on your local computer.

### Prerequisites

1. **Python 3.7 or higher** must be installed on your system
   - Download from: https://www.python.org/downloads/
   - During installation, make sure to check "Add Python to PATH"

2. **Basic Command Line Knowledge**
   - Windows: Command Prompt or PowerShell
   - Mac/Linux: Terminal

### Installation Steps

#### Step 1: Create a Project Directory

Create a new folder for the UPI Search Tool:

```bash
# Windows
mkdir UPI_Search_Tool
cd UPI_Search_Tool

# Mac/Linux
mkdir UPI_Search_Tool
cd UPI_Search_Tool
```

#### Step 2: Install Required Python Packages

Run the following command to install necessary packages:

```bash
pip install pandas openpyxl
```

If you encounter permission issues, try:

```bash
# Windows
pip install --user pandas openpyxl

# Mac/Linux
pip3 install --user pandas openpyxl
```

#### Step 3: Download Tool Files

Place the following files in your UPI_Search_Tool directory:
- `upi_search_tool.py` - GUI version
- `upi_search_batch.py` - Batch processing version
- `upi_search_test_cases.py` - Test cases
- `README.md` - Documentation
- Sample data files (for testing)

#### Step 4: Verify Installation

Test the installation by running:

```bash
# Test the batch version
python upi_search_batch.py --help

# Test the GUI version (if you have a graphical interface)
python upi_search_tool.py
```

### Quick Start Example

#### Using Sample Data

1. **Prepare Sample UPI Data** (`sample_upi.json`):
```json
{
  "upis": [
    {
      "upiCode": "FXUSDEUR0001",
      "assetClass": "ForeignExchange",
      "instrumentType": "Forward",
      "product": "Vanilla",
      "underlying": {
        "currencyPair": "USD/EUR"
      },
      "deliveryType": "Physical"
    },
    {
      "upiCode": "IRUSDFLT0001",
      "assetClass": "Rates",
      "instrumentType": "Swap",
      "product": "Fixed_Float",
      "underlying": {
        "referenceRate": "USD-LIBOR-3M",
        "currency": "USD",
        "term": "3M"
      },
      "deliveryType": "Cash"
    }
  ]
}
```

2. **Prepare Sample Trade Data** (`sample_trades.xlsx`):

| TradeID | AssetClass | ProductType | CcyPair | RefRate | SettlementType |
|---------|------------|-------------|---------|---------|----------------|
| T001 | FX | Forward | USD/EUR | | Physical |
| T002 | Rates | Swap | | USD-LIBOR-3M | Cash |

3. **Run the Tool**:
```bash
python upi_search_batch.py \
  --upi sample_upi.json \
  --trade sample_trades.xlsx \
  --asset-class FX \
  --output test_results.xlsx
```

### Troubleshooting Installation

#### Common Issues

**Issue: "python" command not found**
- Solution: Make sure Python is installed and added to PATH
- Try using `python3` instead of `python`

**Issue: "pip" command not found**
- Solution: Try `python -m pip` instead of `pip`

**Issue: Permission denied when installing packages**
- Solution: Use `--user` flag: `pip install --user pandas openpyxl`

**Issue: "pandas not found" when running the tool**
- Solution: Verify pandas installation: `python -c "import pandas; print('OK')"`

**Issue: GUI version doesn't start**
- Solution: Ensure you have a graphical interface (not applicable for servers)
- Use the batch version instead: `upi_search_batch.py`

#### Advanced Installation Options

**Using Virtual Environment (Recommended)**:
```bash
# Create virtual environment
python -m venv upi_tool_env

# Activate virtual environment
# Windows:
upi_tool_env\Scripts\activate
# Mac/Linux:
source upi_tool_env/bin/activate

# Install packages
pip install pandas openpyxl

# Deactivate when done
deactivate
```

**Using conda (if you have Anaconda/Miniconda)**:
```bash
conda create -n upi_tool python=3.9
conda activate upi_tool
conda install pandas openpyxl
```

### Configuration

#### Environment Variables (Optional)

Set default paths for frequently used files:

```bash
# Windows
set UPI_DEFAULT_PATH=C:\path\to\upi\files
set TRADE_DEFAULT_PATH=C:\path\to\trade\files

# Mac/Linux
export UPI_DEFAULT_PATH=/path/to/upi/files
export TRADE_DEFAULT_PATH=/path/to/trade/files
```

#### Custom Configuration File

Create `config.ini` for default settings:

```ini
[DEFAULT]
upi_file_path = ./data/upi_data.json
min_score_threshold = 50
output_directory = ./results

[FX]
default_mapping = fx_mapping.json

[IR]
default_mapping = ir_mapping.json
```

### Performance Optimization

#### For Large Datasets

1. **Increase Memory Allocation**:
```bash
python -Xmx4g upi_search_batch.py [arguments]
```

2. **Process in Chunks**:
Split large Excel files into smaller chunks (e.g., 1000 rows each)

3. **Use SSD Storage**:
Store data files on SSD for faster I/O

#### Monitoring Performance

Track processing time:
```bash
# Windows
powershell "Measure-Command { python upi_search_batch.py [args] }"

# Mac/Linux
time python upi_search_batch.py [args]
```

### Next Steps

1. **Read the User Guide**: `upi-search-tool-guide.md`
2. **Run Test Cases**: `python upi_search_test_cases.py`
3. **Try with Your Data**: Start with a small sample
4. **Schedule Regular Processing**: Set up automated runs

### Getting Help

- Check the README.md for basic usage
- Review the User Guide for detailed instructions
- Run test cases to verify functionality
- Contact your system administrator for technical issues

### Security Considerations

- Keep UPI reference data files secure
- Implement appropriate access controls
- Regular security updates for Python and packages
- Consider using encrypted storage for sensitive trade data

---

**Installation Complete!** You should now be able to use the UPI Search Automation Tool.