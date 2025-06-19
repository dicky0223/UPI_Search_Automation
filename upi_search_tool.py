import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
from datetime import datetime
import os

class UPISearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("UPI Search Automation Tool")
        self.root.geometry("1200x800")
        
        # Data storage
        self.upi_data = None
        self.trade_data = None
        self.results = None
        self.column_mappings = {}
        
        # Create GUI
        self.create_widgets()
    
    def create_widgets(self):
        """Create the main GUI widgets"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_upload_tab()
        self.create_mapping_tab()
        self.create_results_tab()
    
    def create_upload_tab(self):
        """Create the file upload tab"""
        upload_frame = ttk.Frame(self.notebook)
        self.notebook.add(upload_frame, text="Upload Files")
        
        # UPI File Section
        upi_frame = ttk.LabelFrame(upload_frame, text="UPI Reference Data", padding=10)
        upi_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.upi_file_var = tk.StringVar()
        ttk.Label(upi_frame, text="UPI JSON File:").pack(anchor=tk.W)
        file_frame = ttk.Frame(upi_frame)
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(file_frame, textvariable=self.upi_file_var, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="Browse", command=self.browse_upi_file).pack(side=tk.RIGHT, padx=(5,0))
        
        # Trade File Section
        trade_frame = ttk.LabelFrame(upload_frame, text="Trade Data", padding=10)
        trade_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.trade_file_var = tk.StringVar()
        ttk.Label(trade_frame, text="Trade Excel File:").pack(anchor=tk.W)
        file_frame2 = ttk.Frame(trade_frame)
        file_frame2.pack(fill=tk.X, pady=5)
        ttk.Entry(file_frame2, textvariable=self.trade_file_var, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame2, text="Browse", command=self.browse_trade_file).pack(side=tk.RIGHT, padx=(5,0))
        
        # Asset Class Selection
        asset_frame = ttk.LabelFrame(upload_frame, text="Asset Class", padding=10)
        asset_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.asset_class_var = tk.StringVar(value="FX")
        ttk.Radiobutton(asset_frame, text="FX (Foreign Exchange)", variable=self.asset_class_var, value="FX").pack(anchor=tk.W)
        ttk.Radiobutton(asset_frame, text="IR (Interest Rate)", variable=self.asset_class_var, value="IR").pack(anchor=tk.W)
        
        # Load Button
        ttk.Button(upload_frame, text="Load Data", command=self.load_data).pack(pady=20)
        
        # Status
        self.status_var = tk.StringVar(value="Ready to load files...")
        ttk.Label(upload_frame, textvariable=self.status_var).pack(pady=5)
    
    def create_mapping_tab(self):
        """Create the column mapping tab"""
        mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(mapping_frame, text="Map Columns")
        
        # Instructions
        instructions = ttk.Label(mapping_frame, text="Map your trade data columns to UPI attributes. Auto-suggestions are provided based on common naming patterns.")
        instructions.pack(pady=10)
        
        # Mapping area
        self.mapping_canvas = tk.Canvas(mapping_frame)
        self.mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=self.mapping_canvas.yview)
        self.mapping_scrollable_frame = ttk.Frame(self.mapping_canvas)
        
        self.mapping_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_scrollable_frame, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=self.mapping_scrollbar.set)
        
        self.mapping_canvas.pack(side="left", fill="both", expand=True, padx=10)
        self.mapping_scrollbar.pack(side="right", fill="y")
        
        # Search button
        button_frame = ttk.Frame(mapping_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(button_frame, text="Map Columns & Search UPIs", command=self.search_upis).pack()
    
    def create_results_tab(self):
        """Create the results display tab"""
        results_frame = ttk.Frame(self.notebook)
        self.notebook.add(results_frame, text="Results")
        
        # Results tree
        columns = ("Trade Index", "Best UPI", "Match Score", "Trade Attributes")
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=20)
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=200)
        
        # Scrollbars for results
        v_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        h_scrollbar = ttk.Scrollbar(results_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack results tree and scrollbars
        self.results_tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        
        # Export button
        export_frame = ttk.Frame(results_frame)
        export_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(export_frame, text="Export Results to Excel", command=self.export_results).pack()
    
    def browse_upi_file(self):
        """Browse for UPI JSON file"""
        filename = filedialog.askopenfilename(
            title="Select UPI JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.upi_file_var.set(filename)
    
    def browse_trade_file(self):
        """Browse for trade Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Trade Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.trade_file_var.set(filename)
    
    def load_data(self):
        """Load UPI and trade data"""
        try:
            # Load UPI data
            upi_file = self.upi_file_var.get()
            if not upi_file:
                messagebox.showerror("Error", "Please select a UPI JSON file")
                return
            
            with open(upi_file, 'r') as f:
                self.upi_data = json.load(f)
            
            # Load trade data
            trade_file = self.trade_file_var.get()
            if not trade_file:
                messagebox.showerror("Error", "Please select a trade Excel file")
                return
            
            self.trade_data = pd.read_excel(trade_file)
            
            # Apply CNH handling
            self.apply_cnh_handling()
            
            # Auto-map columns
            self.auto_map_columns()
            
            # Update status
            self.status_var.set(f"Loaded {len(self.upi_data.get('upis', []))} UPIs and {len(self.trade_data)} trades")
            
            # Switch to mapping tab
            self.notebook.select(1)
            
            messagebox.showinfo("Success", "Data loaded successfully! Please review column mappings.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
    
    def apply_cnh_handling(self):
        """Apply CNH-specific handling logic to trade data"""
        # Create new columns for CNH handling if they don't exist
        if 'ProcessedUseCase' not in self.trade_data.columns:
            self.trade_data['ProcessedUseCase'] = ''
        if 'ProcessedPlaceofSettlement' not in self.trade_data.columns:
            self.trade_data['ProcessedPlaceofSettlement'] = ''
        if 'ProcessedCurrency' not in self.trade_data.columns:
            self.trade_data['ProcessedCurrency'] = ''
        
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
                # Set PlaceofSettlement to Hong Kong for CNH trades
                self.trade_data.at[idx, 'ProcessedPlaceofSettlement'] = 'Hong Kong'
                
                # Determine UseCase based on InstrumentType
                instrument_type = self.get_instrument_type_from_row(row)
                
                if instrument_type:
                    if instrument_type.upper() == 'SWAP':
                        self.trade_data.at[idx, 'ProcessedUseCase'] = 'Non_Deliverable_FX_Swap'
                    elif instrument_type.upper() in ['FORWARD', 'OPTION']:
                        self.trade_data.at[idx, 'ProcessedUseCase'] = 'Non_Standard'
    
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
    
    def auto_map_columns(self):
        """Automatically map columns based on common naming patterns"""
        if self.trade_data is None:
            return
        
        trade_columns = list(self.trade_data.columns)
        asset_class = self.asset_class_var.get()
        
        if asset_class == "FX":
            upi_attributes = [
                'Asset Class', 'Instrument Type', 'Product Type', 'Currency Pair',
                'Settlement Currency', 'Option Type', 'Option Style', 'Delivery Type',
                'Place of Settlement'
            ]
        else:  # IR
            upi_attributes = [
                'Asset Class', 'Instrument Type', 'Product Type', 'Reference Rate',
                'Currency', 'Term', 'Other Leg Reference Rate', 'Other Leg Currency',
                'Other Leg Term', 'Delivery Type'
            ]
        
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
        
        self.column_mappings = {}
        
        for attr in upi_attributes:
            for suggestion in suggestions.get(attr, []):
                for col in trade_columns:
                    if col.lower() == suggestion.lower():
                        self.column_mappings[attr] = col
                        break
                if attr in self.column_mappings:
                    break
        
        # Create mapping widgets
        self.create_mapping_widgets(upi_attributes, trade_columns)
    
    def create_mapping_widgets(self, upi_attributes, trade_columns):
        """Create widgets for column mapping"""
        # Clear existing widgets
        for widget in self.mapping_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.mapping_vars = {}
        
        # Add "N/A" option to trade columns
        column_options = ["N/A"] + trade_columns
        
        for i, attr in enumerate(upi_attributes):
            frame = ttk.Frame(self.mapping_scrollable_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Attribute label
            ttk.Label(frame, text=f"{attr}:", width=25).pack(side=tk.LEFT)
            
            # Dropdown for column selection
            var = tk.StringVar()
            if attr in self.column_mappings:
                var.set(self.column_mappings[attr])
            else:
                var.set("N/A")
            
            combo = ttk.Combobox(frame, textvariable=var, values=column_options, width=30)
            combo.pack(side=tk.LEFT, padx=10)
            
            self.mapping_vars[attr] = var
    
    def search_upis(self):
        """Perform UPI search"""
        if self.trade_data is None or self.upi_data is None:
            messagebox.showerror("Error", "Please load data first")
            return
        
        try:
            # Update column mappings from GUI
            self.column_mappings = {}
            for attr, var in self.mapping_vars.items():
                if var.get() != "N/A":
                    self.column_mappings[attr] = var.get()
            
            # Perform search
            results = []
            asset_class = self.asset_class_var.get()
            
            for idx, trade in self.trade_data.iterrows():
                best_match = None
                best_score = 0
                
                # Extract trade attributes using column mappings
                trade_attrs = self.extract_trade_attributes(trade)
                
                # Search through UPI data
                for upi in self.upi_data.get('upis', []):
                    score = self.calculate_match_score(trade_attrs, upi, asset_class)
                    
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
            
            self.results = results
            
            # Display results
            self.display_results()
            
            # Switch to results tab
            self.notebook.select(2)
            
            # Show summary
            matched_trades = sum(1 for r in results if r['Match_Score'] >= 50)
            messagebox.showinfo("Search Complete", 
                              f"Processed {len(results)} trades.\n"
                              f"Found matches for {matched_trades} trades (score ≥ 50).")
            
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")
    
    def extract_trade_attributes(self, trade):
        """Extract trade attributes using column mappings and CNH processing"""
        attrs = {}
        
        for upi_attr, trade_col in self.column_mappings.items():
            if trade_col in trade.index:
                value = trade[trade_col]
                if pd.notna(value):
                    attrs[upi_attr] = str(value)
        
        # Extract individual currencies from currency pair for bidirectional matching
        if 'Currency Pair' in attrs:
            currency_pair = attrs['Currency Pair']
            if '/' in currency_pair:
                currencies = currency_pair.split('/')
                if len(currencies) == 2:
                    attrs['TradeNotionalCurrency'] = currencies[0].strip()
                    attrs['TradeOtherNotionalCurrency'] = currencies[1].strip()
        
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
            
            # Update individual currencies if CNH was processed to CNY
            if 'TradeNotionalCurrency' in attrs and attrs['TradeNotionalCurrency'].upper() == 'CNH':
                attrs['TradeNotionalCurrency'] = 'CNY'
            if 'TradeOtherNotionalCurrency' in attrs and attrs['TradeOtherNotionalCurrency'].upper() == 'CNH':
                attrs['TradeOtherNotionalCurrency'] = 'CNY'
        
        return attrs
    
    def calculate_match_score(self, trade_attrs, upi, asset_class):
        """Calculate match score between trade and UPI with bidirectional currency matching"""
        score = 0
        
        if asset_class == "FX":
            # FX scoring weights - removed Currency Pair, added individual currency matching
            weights = {
                'Asset Class': 20,
                'Instrument Type': 20,
                'Product Type': 20,
                'Notional Currency': 10,
                'Other Notional Currency': 10,
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
            if attr in ['Notional Currency', 'Other Notional Currency'] and asset_class == "FX":
                # Handle bidirectional currency matching for FX
                if self.match_currencies_bidirectional(trade_attrs, upi):
                    score += weights['Notional Currency'] + weights['Other Notional Currency']
                    break  # Only count this once for both currencies
            elif attr in trade_attrs:
                trade_value = str(trade_attrs[attr]).upper()
                upi_value = self.get_upi_attribute_value(upi, attr)
                
                if upi_value and trade_value == upi_value.upper():
                    score += weight
        
        return score
    
    def match_currencies_bidirectional(self, trade_attrs, upi):
        """Check if trade currencies match UPI currencies in either order"""
        # Get trade currencies
        trade_ccy1 = trade_attrs.get('TradeNotionalCurrency', '').upper()
        trade_ccy2 = trade_attrs.get('TradeOtherNotionalCurrency', '').upper()
        
        # Get UPI currencies
        upi_ccy1 = self.get_upi_attribute_value(upi, 'Notional Currency').upper()
        upi_ccy2 = self.get_upi_attribute_value(upi, 'Other Notional Currency').upper()
        
        # Check if currencies are available
        if not all([trade_ccy1, trade_ccy2, upi_ccy1, upi_ccy2]):
            return False
        
        # Check both orders: (trade1, trade2) == (upi1, upi2) OR (trade1, trade2) == (upi2, upi1)
        return ((trade_ccy1 == upi_ccy1 and trade_ccy2 == upi_ccy2) or 
                (trade_ccy1 == upi_ccy2 and trade_ccy2 == upi_ccy1))
    
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
            'Other Leg Term': lambda u: u.get('otherLeg', {}).get('term', ''),
            'Notional Currency': lambda u: self.extract_currency_from_pair(u.get('underlying', {}).get('currencyPair', ''), 0),
            'Other Notional Currency': lambda u: self.extract_currency_from_pair(u.get('underlying', {}).get('currencyPair', ''), 1)
        }
        
        if attribute in mapping:
            return mapping[attribute](upi)
        
        return ''
    
    def extract_currency_from_pair(self, currency_pair, index):
        """Extract individual currency from currency pair (e.g., 'USD/EUR' -> 'USD' or 'EUR')"""
        if '/' in currency_pair:
            currencies = currency_pair.split('/')
            if len(currencies) > index:
                return currencies[index].strip()
        return ''
    
    def display_results(self):
        """Display search results in the tree view"""
        # Clear existing results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Add results
        for result in self.results:
            self.results_tree.insert("", "end", values=(
                result['Trade_Index'],
                result['Best_UPI'],
                f"{result['Match_Score']:.1f}",
                str(result['Trade_Attributes'])[:100] + "..." if len(str(result['Trade_Attributes'])) > 100 else str(result['Trade_Attributes'])
            ))
    
    def export_results(self):
        """Export results to Excel file"""
        if not self.results:
            messagebox.showerror("Error", "No results to export")
            return
        
        try:
            # Ask for save location
            filename = filedialog.asksaveasfilename(
                title="Save Results",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not filename:
                return
            
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
            df.to_excel(filename, index=False)
            
            messagebox.showinfo("Success", f"Results exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export results: {str(e)}")

def main():
    root = tk.Tk()
    app = UPISearchTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()