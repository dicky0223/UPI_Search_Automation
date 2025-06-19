import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import json
from datetime import datetime
import os

class UPISearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("UPI Search Automation Tool")
        self.root.geometry("1000x700")
        
        # Data storage
        self.upi_data = None
        self.trade_data = None
        self.results = None
        self.asset_class = tk.StringVar(value="FX")
        
        # File paths
        self.upi_file_path = tk.StringVar()
        self.trade_file_path = tk.StringVar()
        
        # Column mappings
        self.column_mappings = {}
        
        self.create_widgets()
    
    def create_widgets(self):
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: File Upload
        upload_frame = ttk.Frame(notebook)
        notebook.add(upload_frame, text="Upload Files")
        self.create_upload_tab(upload_frame)
        
        # Tab 2: Column Mapping
        mapping_frame = ttk.Frame(notebook)
        notebook.add(mapping_frame, text="Map Columns")
        self.create_mapping_tab(mapping_frame)
        
        # Tab 3: Results
        results_frame = ttk.Frame(notebook)
        notebook.add(results_frame, text="Results")
        self.create_results_tab(results_frame)
    
    def create_upload_tab(self, parent):
        # UPI File Selection
        ttk.Label(parent, text="UPI JSON File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent, textvariable=self.upi_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(parent, text="Browse", command=self.browse_upi_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Trade File Selection
        ttk.Label(parent, text="Trade Excel File:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent, textvariable=self.trade_file_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(parent, text="Browse", command=self.browse_trade_file).grid(row=1, column=2, padx=5, pady=5)
        
        # Asset Class Selection
        ttk.Label(parent, text="Asset Class:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        asset_frame = ttk.Frame(parent)
        asset_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Radiobutton(asset_frame, text="FX", variable=self.asset_class, value="FX").pack(side=tk.LEFT)
        ttk.Radiobutton(asset_frame, text="IR", variable=self.asset_class, value="IR").pack(side=tk.LEFT, padx=(10, 0))
        
        # Load Data Button
        ttk.Button(parent, text="Load Data", command=self.load_data).grid(row=3, column=1, pady=20)
        
        # Status Text
        self.status_text = scrolledtext.ScrolledText(parent, height=10, width=80)
        self.status_text.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky=tk.NSEW)
        
        # Configure grid weights
        parent.grid_rowconfigure(4, weight=1)
        parent.grid_columnconfigure(1, weight=1)
    
    def create_mapping_tab(self, parent):
        # Instructions
        instructions = ttk.Label(parent, text="Map your Excel columns to UPI attributes. Select 'N/A' for attributes not present in your data.")
        instructions.pack(pady=10)
        
        # Mapping frame with scrollbar
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.mapping_frame = ttk.Frame(canvas)
        
        self.mapping_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.mapping_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Map and Search Button
        ttk.Button(parent, text="Map Columns & Search UPIs", command=self.map_and_search).pack(pady=10)
    
    def create_results_tab(self, parent):
        # Results display
        self.results_text = scrolledtext.ScrolledText(parent, height=25, width=100)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Export button
        ttk.Button(parent, text="Export Results to Excel", command=self.export_results).pack(pady=10)
    
    def browse_upi_file(self):
        filename = filedialog.askopenfilename(
            title="Select UPI JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.upi_file_path.set(filename)
    
    def browse_trade_file(self):
        filename = filedialog.askopenfilename(
            title="Select Trade Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.trade_file_path.set(filename)
    
    def log_status(self, message):
        self.status_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.status_text.see(tk.END)
        self.root.update()
    
    def load_data(self):
        try:
            if not self.upi_file_path.get() or not self.trade_file_path.get():
                messagebox.showerror("Error", "Please select both UPI and trade files.")
                return
            
            self.log_status("Loading UPI data...")
            with open(self.upi_file_path.get(), 'r') as f:
                self.upi_data = json.load(f)
            
            self.log_status("Loading trade data...")
            self.trade_data = pd.read_excel(self.trade_file_path.get())
            
            self.log_status(f"Loaded {len(self.upi_data.get('upis', []))} UPI records")
            self.log_status(f"Loaded {len(self.trade_data)} trade records")
            
            # Apply CNH handling logic
            self.apply_cnh_handling()
            
            # Create column mapping interface
            self.create_column_mapping_interface()
            
            self.log_status("Data loaded successfully. Please proceed to Map Columns tab.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading files: {str(e)}")
            self.log_status(f"Error: {str(e)}")
    
    def apply_cnh_handling(self):
        """Apply CNH-specific handling logic to trade data"""
        self.log_status("Applying CNH handling logic...")
        
        # Create new columns for CNH handling if they don't exist
        if 'ProcessedUseCase' not in self.trade_data.columns:
            self.trade_data['ProcessedUseCase'] = ''
        if 'ProcessedPlaceofSettlement' not in self.trade_data.columns:
            self.trade_data['ProcessedPlaceofSettlement'] = ''
        if 'ProcessedCurrency' not in self.trade_data.columns:
            self.trade_data['ProcessedCurrency'] = ''
        
        cnh_trades_count = 0
        
        for idx, row in self.trade_data.iterrows():
            # Check for CNH in any currency-related columns
            is_cnh_trade = False
            
            # Check all columns for CNH currency
            for col in self.trade_data.columns:
                if pd.notna(row[col]) and str(row[col]).upper() in ['CNH', 'CNY']:
                    is_cnh_trade = True
                    # Normalize CNH to CNY for UPI matching
                    if str(row[col]).upper() == 'CNH':
                        self.trade_data.at[idx, 'ProcessedCurrency'] = 'CNY'
                    else:
                        self.trade_data.at[idx, 'ProcessedCurrency'] = str(row[col]).upper()
                    break
            
            if is_cnh_trade:
                cnh_trades_count += 1
                
                # Set PlaceofSettlement to Hong Kong for CNH trades
                self.trade_data.at[idx, 'ProcessedPlaceofSettlement'] = 'Hong Kong'
                
                # Determine UseCase based on InstrumentType
                instrument_type = self.get_instrument_type_from_row(row)
                
                if instrument_type:
                    if instrument_type.upper() == 'SWAP':
                        self.trade_data.at[idx, 'ProcessedUseCase'] = 'Non_Deliverable_FX_Swap'
                    elif instrument_type.upper() in ['FORWARD', 'OPTION']:
                        self.trade_data.at[idx, 'ProcessedUseCase'] = 'Non_Standard'
        
        if cnh_trades_count > 0:
            self.log_status(f"Applied CNH handling to {cnh_trades_count} trades")
            self.log_status("CNH trades will use:")
            self.log_status("- Currency: CNY (normalized from CNH)")
            self.log_status("- PlaceofSettlement: Hong Kong")
            self.log_status("- UseCase: Non_Deliverable_FX_Swap (Swaps) or Non_Standard (Forwards/Options)")
        else:
            self.log_status("No CNH trades detected")
    
    def get_instrument_type_from_row(self, row):
        """Extract instrument type from trade row"""
        # Common column names for instrument type
        instrument_type_columns = [
            'InstrumentType', 'Instrument_Type', 'ProductType', 'Product_Type',
            'TradeType', 'Trade_Type', 'Type', 'Instrument'
        ]
        
        for col in instrument_type_columns:
            if col in row.index and pd.notna(row[col]):
                return str(row[col])
        
        return None
    
    def create_column_mapping_interface(self):
        # Clear existing mapping widgets
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        
        # Get trade data columns
        trade_columns = ['N/A'] + list(self.trade_data.columns)
        
        # Define UPI attributes based on asset class
        if self.asset_class.get() == "FX":
            upi_attributes = [
                'Asset Class', 'Instrument Type', 'Product Type', 'Currency Pair',
                'Settlement Currency', 'Option Type', 'Option Style', 'Delivery Type',
                'Place of Settlement'  # Added for CNH handling
            ]
        else:  # IR
            upi_attributes = [
                'Asset Class', 'Instrument Type', 'Product Type', 'Reference Rate',
                'Currency', 'Term', 'Other Leg Reference Rate', 'Other Leg Currency',
                'Other Leg Term', 'Delivery Type'
            ]
        
        # Create mapping widgets
        self.mapping_vars = {}
        for i, attr in enumerate(upi_attributes):
            ttk.Label(self.mapping_frame, text=f"{attr}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            
            var = tk.StringVar()
            combobox = ttk.Combobox(self.mapping_frame, textvariable=var, values=trade_columns, width=30)
            combobox.grid(row=i, column=1, padx=5, pady=2)
            
            # Auto-suggest mapping
            suggested = self.suggest_column_mapping(attr, trade_columns)
            if suggested:
                var.set(suggested)
            else:
                var.set('N/A')
            
            self.mapping_vars[attr] = var
    
    def suggest_column_mapping(self, upi_attribute, trade_columns):
        """Suggest column mapping based on common naming patterns"""
        suggestions = {
            'Asset Class': ['AssetClass', 'Asset_Class', 'Product_Class', 'Class'],
            'Instrument Type': ['InstrumentType', 'Instrument_Type', 'ProductType', 'Product_Type', 'TradeType', 'Type'],
            'Product Type': ['Product', 'ProductType', 'Product_Type', 'SubProduct', 'ForwardType'],
            'Currency Pair': ['CcyPair', 'CurrencyPair', 'Ccy_Pair', 'Currency_Pair', 'Pair'],
            'Settlement Currency': ['SettlementCcy', 'Settlement_Currency', 'SettleCcy', 'ProcessedCurrency'],
            'Option Type': ['OptionType', 'Option_Type', 'CallPut', 'Call_Put'],
            'Option Style': ['OptionStyle', 'Option_Style', 'ExerciseStyle', 'Exercise_Style'],
            'Delivery Type': ['DeliveryType', 'Delivery_Type', 'SettlementType', 'Settlement_Type'],
            'Place of Settlement': ['PlaceofSettlement', 'Place_of_Settlement', 'ProcessedPlaceofSettlement'],
            'Reference Rate': ['RefRate', 'ReferenceRate', 'Reference_Rate', 'IndexRate', 'Index_Rate'],
            'Currency': ['Currency', 'Ccy', 'Currency1', 'ProcessedCurrency'],
            'Term': ['Term', 'Tenor', 'Term1', 'Maturity'],
            'Other Leg Reference Rate': ['RefRate2', 'OtherLegRate', 'Other_Leg_Rate', 'IndexRate2'],
            'Other Leg Currency': ['Currency2', 'OtherLegCcy', 'Other_Leg_Currency'],
            'Other Leg Term': ['Term2', 'OtherLegTerm', 'Other_Leg_Term', 'Tenor2']
        }
        
        for suggestion in suggestions.get(upi_attribute, []):
            for col in trade_columns:
                if col.lower() == suggestion.lower():
                    return col
        
        return None
    
    def map_and_search(self):
        if self.trade_data is None:
            messagebox.showerror("Error", "Please load data first.")
            return
        
        try:
            # Get column mappings
            self.column_mappings = {}
            for attr, var in self.mapping_vars.items():
                if var.get() != 'N/A':
                    self.column_mappings[attr] = var.get()
            
            self.log_status("Starting UPI search...")
            
            # Perform UPI search
            self.results = self.search_upis()
            
            # Display results
            self.display_results()
            
            self.log_status("UPI search completed.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during UPI search: {str(e)}")
            self.log_status(f"Error: {str(e)}")
    
    def search_upis(self):
        results = []
        
        for idx, trade in self.trade_data.iterrows():
            best_match = None
            best_score = 0
            
            # Extract trade attributes using column mappings
            trade_attrs = self.extract_trade_attributes(trade)
            
            # Search through UPI data
            for upi in self.upi_data.get('upis', []):
                score = self.calculate_match_score(trade_attrs, upi)
                
                if score > best_score:
                    best_score = score
                    best_match = upi
            
            # Prepare result
            result = {
                'Trade_Index': idx,
                'Best_UPI': best_match['upiCode'] if best_match else 'No Match',
                'Match_Score': best_score,
                'Trade_Attributes': trade_attrs,
                'UPI_Details': best_match if best_match else {}
            }
            
            # Add original trade data
            for col in self.trade_data.columns:
                result[f'Original_{col}'] = trade[col]
            
            results.append(result)
        
        return results
    
    def extract_trade_attributes(self, trade):
        """Extract trade attributes using column mappings and CNH processing"""
        attrs = {}
        
        for upi_attr, trade_col in self.column_mappings.items():
            if trade_col in trade.index:
                value = trade[trade_col]
                if pd.notna(value):
                    attrs[upi_attr] = str(value)
        
        # Apply CNH-specific overrides
        if 'ProcessedUseCase' in trade.index and pd.notna(trade['ProcessedUseCase']) and trade['ProcessedUseCase']:
            attrs['Product Type'] = trade['ProcessedUseCase']
        
        if 'ProcessedPlaceofSettlement' in trade.index and pd.notna(trade['ProcessedPlaceofSettlement']) and trade['ProcessedPlaceofSettlement']:
            attrs['Place of Settlement'] = trade['ProcessedPlaceofSettlement']
        
        if 'ProcessedCurrency' in trade.index and pd.notna(trade['ProcessedCurrency']) and trade['ProcessedCurrency']:
            # Override currency-related attributes with processed currency
            if 'Currency' in attrs:
                attrs['Currency'] = trade['ProcessedCurrency']
            if 'Settlement Currency' in attrs:
                attrs['Settlement Currency'] = trade['ProcessedCurrency']
        
        return attrs
    
    def calculate_match_score(self, trade_attrs, upi):
        """Calculate match score between trade and UPI"""
        score = 0
        max_score = 100
        
        if self.asset_class.get() == "FX":
            # FX scoring weights
            weights = {
                'Asset Class': 20,
                'Instrument Type': 20,
                'Product Type': 20,
                'Currency Pair': 20,
                'Settlement Currency': 10,
                'Option Type': 5,
                'Option Style': 5,
                'Delivery Type': 10,
                'Place of Settlement': 10
            }
        else:
            # IR scoring weights
            weights = {
                'Asset Class': 20,
                'Instrument Type': 20,
                'Product Type': 20,
                'Reference Rate': 15,
                'Currency': 10,
                'Term': 10,
                'Other Leg Reference Rate': 10,
                'Other Leg Currency': 5,
                'Other Leg Term': 5,
                'Delivery Type': 10
            }
        
        # Calculate score based on matches
        for attr, weight in weights.items():
            if attr in trade_attrs:
                trade_value = str(trade_attrs[attr]).upper()
                upi_value = self.get_upi_attribute_value(upi, attr)
                
                if upi_value and trade_value == upi_value.upper():
                    score += weight
        
        return score
    
    def get_upi_attribute_value(self, upi, attribute):
        """Extract attribute value from UPI record"""
        mapping = {
            'Asset Class': lambda u: u.get('assetClass', ''),
            'Instrument Type': lambda u: u.get('instrumentType', ''),
            'Product Type': lambda u: u.get('product', ''),
            'Currency Pair': lambda u: u.get('underlying', {}).get('currencyPair', ''),
            'Settlement Currency': lambda u: u.get('underlying', {}).get('settlementCurrency', ''),
            'Option Type': lambda u: u.get('optionType', ''),
            'Option Style': lambda u: u.get('optionStyle', ''),
            'Delivery Type': lambda u: u.get('deliveryType', ''),
            'Place of Settlement': lambda u: u.get('placeOfSettlement', ''),
            'Reference Rate': lambda u: u.get('underlying', {}).get('referenceRate', ''),
            'Currency': lambda u: u.get('underlying', {}).get('currency', ''),
            'Term': lambda u: u.get('underlying', {}).get('term', ''),
            'Other Leg Reference Rate': lambda u: u.get('otherLeg', {}).get('referenceRate', ''),
            'Other Leg Currency': lambda u: u.get('otherLeg', {}).get('currency', ''),
            'Other Leg Term': lambda u: u.get('otherLeg', {}).get('term', '')
        }
        
        if attribute in mapping:
            return mapping[attribute](upi)
        
        return ''
    
    def display_results(self):
        self.results_text.delete(1.0, tk.END)
        
        if not self.results:
            self.results_text.insert(tk.END, "No results to display.\n")
            return
        
        # Summary
        total_trades = len(self.results)
        matched_trades = sum(1 for r in self.results if r['Match_Score'] >= 50)
        high_confidence = sum(1 for r in self.results if r['Match_Score'] >= 80)
        
        summary = f"""UPI Search Results Summary
========================
Total Trades: {total_trades}
Matched Trades (Score ≥ 50): {matched_trades}
High Confidence Matches (Score ≥ 80): {high_confidence}
Match Rate: {(matched_trades/total_trades)*100:.1f}%

Detailed Results:
================
"""
        self.results_text.insert(tk.END, summary)
        
        # Detailed results
        for i, result in enumerate(self.results[:10]):  # Show first 10 results
            detail = f"""
Trade {i+1}:
  UPI: {result['Best_UPI']}
  Score: {result['Match_Score']}/100
  Trade Attributes: {result['Trade_Attributes']}
  
"""
            self.results_text.insert(tk.END, detail)
        
        if len(self.results) > 10:
            self.results_text.insert(tk.END, f"\n... and {len(self.results) - 10} more results. Export to Excel to see all results.\n")
    
    def export_results(self):
        if not self.results:
            messagebox.showerror("Error", "No results to export.")
            return
        
        try:
            # Prepare data for export
            export_data = []
            
            for result in self.results:
                row = {
                    'Trade_Index': result['Trade_Index'],
                    'Best_UPI': result['Best_UPI'],
                    'Match_Score': result['Match_Score'],
                    'Trade_Attributes': str(result['Trade_Attributes']),
                    'UPI_Details': str(result['UPI_Details'])
                }
                
                # Add original trade data
                for key, value in result.items():
                    if key.startswith('Original_'):
                        row[key] = value
                
                export_data.append(row)
            
            # Create DataFrame and export
            df = pd.DataFrame(export_data)
            
            # Ask for save location
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Results"
            )
            
            if filename:
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.log_status(f"Results exported to {filename}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting results: {str(e)}")
            self.log_status(f"Export error: {str(e)}")

def main():
    root = tk.Tk()
    app = UPISearchTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()