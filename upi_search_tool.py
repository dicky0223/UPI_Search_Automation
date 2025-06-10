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
        self.asset_class = tk.StringVar(value="FX")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_upload_tab()
        self.create_mapping_tab()
        self.create_results_tab()
        
    def create_upload_tab(self):
        # Upload Files Tab
        upload_frame = ttk.Frame(self.notebook)
        self.notebook.add(upload_frame, text="Upload Files")
        
        # UPI File Selection
        ttk.Label(upload_frame, text="UPI JSON File:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.upi_file_var = tk.StringVar()
        ttk.Entry(upload_frame, textvariable=self.upi_file_var, width=50).grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(upload_frame, text="Browse", command=self.browse_upi_file).grid(row=0, column=2, padx=10, pady=5)
        
        # Trade File Selection
        ttk.Label(upload_frame, text="Trade Excel File:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.trade_file_var = tk.StringVar()
        ttk.Entry(upload_frame, textvariable=self.trade_file_var, width=50).grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(upload_frame, text="Browse", command=self.browse_trade_file).grid(row=1, column=2, padx=10, pady=5)
        
        # Asset Class Selection
        ttk.Label(upload_frame, text="Asset Class:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        asset_frame = ttk.Frame(upload_frame)
        asset_frame.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(asset_frame, text="FX", variable=self.asset_class, value="FX").pack(side=tk.LEFT)
        ttk.Radiobutton(asset_frame, text="IR", variable=self.asset_class, value="IR").pack(side=tk.LEFT, padx=(20, 0))
        
        # Load Data Button
        ttk.Button(upload_frame, text="Load Data", command=self.load_data).grid(row=3, column=1, pady=20)
        
        # Status Label
        self.status_label = ttk.Label(upload_frame, text="Ready to load data...")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=10)
        
    def create_mapping_tab(self):
        # Column Mapping Tab
        mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(mapping_frame, text="Map Columns")
        
        # Instructions
        instructions = ttk.Label(mapping_frame, text="Map your Excel columns to UPI attributes. Select 'N/A' for attributes not present in your data.")
        instructions.pack(pady=10)
        
        # Mapping area (will be populated after data load)
        self.mapping_area = ttk.Frame(mapping_frame)
        self.mapping_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Map and Search Button
        self.map_search_button = ttk.Button(mapping_frame, text="Map Columns & Search UPIs", command=self.map_and_search)
        self.map_search_button.pack(pady=10)
        
    def create_results_tab(self):
        # Results Tab
        results_frame = ttk.Frame(self.notebook)
        self.notebook.add(results_frame, text="Results")
        
        # Results area
        self.results_tree = ttk.Treeview(results_frame)
        self.results_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbars for results
        v_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.results_tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.results_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Export Button
        ttk.Button(results_frame, text="Export Results to Excel", command=self.export_results).pack(pady=10)
        
    def browse_upi_file(self):
        filename = filedialog.askopenfilename(
            title="Select UPI JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.upi_file_var.set(filename)
            
    def browse_trade_file(self):
        filename = filedialog.askopenfilename(
            title="Select Trade Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.trade_file_var.set(filename)
            
    def load_data(self):
        try:
            # Load UPI data
            if not self.upi_file_var.get():
                messagebox.showerror("Error", "Please select a UPI JSON file")
                return
                
            with open(self.upi_file_var.get(), 'r') as f:
                self.upi_data = json.load(f)
                
            # Load trade data
            if not self.trade_file_var.get():
                messagebox.showerror("Error", "Please select a trade Excel file")
                return
                
            self.trade_data = pd.read_excel(self.trade_file_var.get())
            
            self.status_label.config(text=f"Loaded {len(self.upi_data.get('upis', []))} UPIs and {len(self.trade_data)} trades")
            
            # Create mapping interface
            self.create_mapping_interface()
            
            # Switch to mapping tab
            self.notebook.select(1)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading files: {str(e)}")
            
    def create_mapping_interface(self):
        # Clear existing mapping interface
        for widget in self.mapping_area.winfo_children():
            widget.destroy()
            
        # Get mapping fields based on asset class
        if self.asset_class.get() == "FX":
            mapping_fields = self.get_fx_mapping_fields()
        else:
            mapping_fields = self.get_ir_mapping_fields()
            
        # Create mapping dropdowns
        self.mapping_vars = {}
        
        # Headers
        ttk.Label(self.mapping_area, text="UPI Attribute", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ttk.Label(self.mapping_area, text="Your Excel Column", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Get column names from trade data
        column_names = ["N/A"] + list(self.trade_data.columns)
        
        row = 1
        for field in mapping_fields:
            ttk.Label(self.mapping_area, text=field).grid(row=row, column=0, padx=10, pady=2, sticky="w")
            
            var = tk.StringVar()
            # Auto-suggest mapping based on common column names
            suggested_mapping = self.suggest_column_mapping(field, column_names)
            var.set(suggested_mapping)
            
            dropdown = ttk.Combobox(self.mapping_area, textvariable=var, values=column_names, width=30)
            dropdown.grid(row=row, column=1, padx=10, pady=2, sticky="w")
            
            self.mapping_vars[field] = var
            row += 1
            
    def get_fx_mapping_fields(self):
        """Get mapping fields for FX asset class"""
        base_fields = [
            "Asset Class",
            "Instrument Type",
            "Product Type",
            "Notional Currency",
            "Other Notional Currency",
            "Delivery Type"
        ]
        
        # Add product-specific fields
        product_specific_fields = {
            "Forward": ["Settlement Currency"],
            "NDF": ["Settlement Currency"],
            "Non_Standard": [
                "Underlying Asset Type",
                "Return or Payout Trigger", 
                "Option Type",
                "Option Exercise Style",
                "Valuation Method or Trigger",
                "Settlement Currency",
                "Place of Settlement"
            ],
            "Digital_Option": [
                "Option Type",
                "Option Exercise Style", 
                "Valuation Method or Trigger",
                "Settlement Currency"
            ],
            "Vanilla_Option": [
                "Option Type",
                "Option Exercise Style"
            ],
            "FX_Swap": [],
            "Non_Deliverable_FX_Swap": [
                "Settlement Currency",
                "Place of Settlement"
            ]
        }
        
        # For now, include all possible fields
        all_fields = base_fields + [
            "Settlement Currency",
            "Option Type", 
            "Option Exercise Style",
            "Valuation Method or Trigger",
            "Underlying Asset Type",
            "Return or Payout Trigger",
            "Place of Settlement"
        ]
        
        return list(dict.fromkeys(all_fields))  # Remove duplicates while preserving order
        
    def get_ir_mapping_fields(self):
        """Get mapping fields for IR asset class"""
        return [
            "Asset Class",
            "Instrument Type", 
            "Product Type",
            "Notional Currency",
            "Reference Rate",
            "Reference Rate Term Value",
            "Reference Rate Term Unit",
            "Other Leg Reference Rate",
            "Other Leg Reference Rate Term Value", 
            "Other Leg Reference Rate Term Unit",
            "Other Notional Currency",
            "Notional Schedule",
            "Delivery Type"
        ]
        
    def suggest_column_mapping(self, field, column_names):
        """Suggest column mapping based on field name"""
        field_lower = field.lower()
        
        # Common mapping suggestions
        suggestions = {
            "asset class": ["assetclass", "asset_class", "product_class", "asset"],
            "instrument type": ["instrumenttype", "instrument_type", "producttype", "product_type", "tradetype", "trade_type"],
            "product type": ["product", "producttype", "product_type", "usecase", "use_case"],
            "notional currency": ["notionalcurrency", "notional_currency", "currency", "ccy", "currency1", "ccy1"],
            "other notional currency": ["othernotionalcurrency", "other_notional_currency", "currency2", "ccy2", "othercurrency", "other_currency"],
            "delivery type": ["deliverytype", "delivery_type", "settlementtype", "settlement_type"],
            "settlement currency": ["settlementcurrency", "settlement_currency", "settleccy", "settle_ccy"],
            "option type": ["optiontype", "option_type", "callput", "call_put"],
            "option exercise style": ["optionexercisestyle", "option_exercise_style", "exercisestyle", "exercise_style"],
            "reference rate": ["referencerate", "reference_rate", "refrate", "ref_rate", "indexrate", "index_rate"],
            "place of settlement": ["placeofsettlement", "place_of_settlement", "settlementplace", "settlement_place"]
        }
        
        if field_lower in suggestions:
            for col in column_names[1:]:  # Skip "N/A"
                col_lower = col.lower().replace(" ", "").replace("_", "")
                for suggestion in suggestions[field_lower]:
                    if suggestion in col_lower:
                        return col
                        
        return "N/A"
        
    def is_offshore_cny(self, trade_row):
        """Check if trade involves offshore CNH/CNY currency"""
        notional_currency = trade_row.get("Notional Currency", "")
        other_notional_currency = trade_row.get("Other Notional Currency", "")
        
        offshore_currencies = ["CNH", "CNY"]
        
        return (str(notional_currency).upper() in offshore_currencies or 
                str(other_notional_currency).upper() in offshore_currencies)
        
    def map_and_search(self):
        try:
            if self.trade_data is None or self.upi_data is None:
                messagebox.showerror("Error", "Please load data first")
                return
                
            # Create mapped trade data
            mapped_trades = []
            
            for _, trade_row in self.trade_data.iterrows():
                mapped_trade = {}
                
                for field, var in self.mapping_vars.items():
                    column = var.get()
                    if column != "N/A" and column in self.trade_data.columns:
                        mapped_trade[field] = trade_row[column]
                    else:
                        mapped_trade[field] = None
                        
                mapped_trades.append(mapped_trade)
                
            # Search for UPIs
            self.results = []
            
            for i, trade in enumerate(mapped_trades):
                result = self.find_matching_upi(trade, i)
                self.results.append(result)
                
            # Display results
            self.display_results()
            
            # Switch to results tab
            self.notebook.select(2)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during mapping and search: {str(e)}")
            
    def find_matching_upi(self, trade, trade_index):
        """Find matching UPI for a trade with enhanced logic for offshore CNY and currency pair reversal"""
        
        # Check if this is an offshore CNY trade
        is_offshore = self.is_offshore_cny(trade)
        instrument_type = trade.get("Instrument Type", "")
        
        # Determine target product type based on offshore status
        original_product = trade.get("Product Type", "")
        target_products = [original_product]
        
        if is_offshore:
            if instrument_type in ["Forward", "Option"]:
                target_products = ["Non_Standard"] + target_products
            elif instrument_type == "Swap":
                target_products = ["Non_Deliverable_FX_Swap"] + target_products
        
        best_match = None
        best_score = 0
        
        # Try each target product type
        for target_product in target_products:
            # First try with original currency pair
            matches = self.search_upis_for_product(trade, target_product, is_offshore)
            
            for upi in matches:
                score = self.calculate_upi_score(trade, upi, False)  # False = not reversed
                if score > best_score:
                    best_score = score
                    best_match = upi
            
            # If no good match found, try with reversed currency pair
            if best_score < 50:  # Threshold for trying reversal
                reversed_trade = self.create_reversed_currency_trade(trade)
                if reversed_trade != trade:  # Only if reversal actually changed something
                    matches = self.search_upis_for_product(reversed_trade, target_product, is_offshore)
                    
                    for upi in matches:
                        score = self.calculate_upi_score(reversed_trade, upi, True)  # True = reversed
                        if score > best_score:
                            best_score = score
                            best_match = upi
            
            # If we found a good match, stop searching
            if best_score >= 70:
                break
        
        return {
            "Trade Index": trade_index + 1,
            "Best UPI": best_match.get("Identifier", {}).get("UPI", "No match found") if best_match else "No match found",
            "Score": best_score,
            "Match Details": self.get_match_details(trade, best_match) if best_match else "No sufficient match found",
            "Is Offshore": is_offshore,
            "Target Products": ", ".join(target_products)
        }
    
    def search_upis_for_product(self, trade, target_product, is_offshore):
        """Search UPIs for a specific product type with offshore filtering"""
        matches = []
        
        for upi in self.upi_data.get("upis", []):
            # Parse UPI attributes
            if self.asset_class.get() == "FX":
                upi_attrs = self.parse_fx_attributes(upi)
            else:
                upi_attrs = self.parse_ir_attributes(upi)
            
            # Check if this UPI matches the target product
            if upi_attrs.get("Product Type") != target_product:
                continue
            
            # For offshore trades, check place of settlement
            if is_offshore and target_product in ["Non_Standard", "Non_Deliverable_FX_Swap"]:
                place_of_settlement = upi_attrs.get("Place of Settlement", "")
                if place_of_settlement != "Hong Kong":
                    continue
            
            matches.append(upi)
        
        return matches
    
    def create_reversed_currency_trade(self, trade):
        """Create a copy of trade with reversed currency pair"""
        reversed_trade = trade.copy()
        
        notional_currency = trade.get("Notional Currency")
        other_notional_currency = trade.get("Other Notional Currency")
        
        if notional_currency and other_notional_currency:
            reversed_trade["Notional Currency"] = other_notional_currency
            reversed_trade["Other Notional Currency"] = notional_currency
        
        return reversed_trade
        
    def parse_fx_attributes(self, upi):
        """Parse FX UPI attributes from the UPI record"""
        attributes = {}
        
        # Parse Header information
        header = upi.get("Header", {})
        attributes["Asset Class"] = header.get("AssetClass", "")
        attributes["Instrument Type"] = header.get("InstrumentType", "")
        attributes["Product Type"] = header.get("UseCase", "")
        
        # Parse Derived information
        derived = upi.get("Derived", {})
        attributes["Underlying Asset Type"] = derived.get("UnderlyingAssetType", "")
        attributes["Return or Payout Trigger"] = derived.get("ReturnorPayoutTrigger", "")
        attributes["Valuation Method or Trigger"] = derived.get("ValuationMethodorTrigger", "")
        
        # Parse Attributes information
        attrs = upi.get("Attributes", {})
        attributes["Notional Currency"] = attrs.get("NotionalCurrency", "")
        attributes["Other Notional Currency"] = attrs.get("OtherNotionalCurrency", "")
        attributes["Settlement Currency"] = attrs.get("SettlementCurrency", "")
        attributes["Delivery Type"] = attrs.get("DeliveryType", "")
        attributes["Option Type"] = attrs.get("OptionType", "")
        attributes["Option Exercise Style"] = attrs.get("OptionExerciseStyle", "")
        attributes["Place of Settlement"] = attrs.get("PlaceofSettlement", "")
        
        return attributes
        
    def parse_ir_attributes(self, upi):
        """Parse IR UPI attributes from the UPI record"""
        attributes = {}
        
        # Parse Header information
        header = upi.get("Header", {})
        attributes["Asset Class"] = header.get("AssetClass", "")
        attributes["Instrument Type"] = header.get("InstrumentType", "")
        attributes["Product Type"] = header.get("UseCase", "")
        
        # Parse Attributes information
        attrs = upi.get("Attributes", {})
        attributes["Notional Currency"] = attrs.get("NotionalCurrency", "")
        attributes["Reference Rate"] = attrs.get("ReferenceRate", "")
        attributes["Reference Rate Term Value"] = attrs.get("ReferenceRateTermValue", "")
        attributes["Reference Rate Term Unit"] = attrs.get("ReferenceRateTermUnit", "")
        attributes["Other Leg Reference Rate"] = attrs.get("OtherLegReferenceRate", "")
        attributes["Other Leg Reference Rate Term Value"] = attrs.get("OtherLegReferenceRateTermValue", "")
        attributes["Other Leg Reference Rate Term Unit"] = attrs.get("OtherLegReferenceRateTermUnit", "")
        attributes["Other Notional Currency"] = attrs.get("OtherNotionalCurrency", "")
        attributes["Notional Schedule"] = attrs.get("NotionalSchedule", "")
        attributes["Delivery Type"] = attrs.get("DeliveryType", "")
        
        return attributes
        
    def calculate_upi_score(self, trade, upi, is_reversed=False):
        """Calculate matching score between trade and UPI"""
        score = 0
        max_score = 0
        
        # Parse UPI attributes
        if self.asset_class.get() == "FX":
            upi_attrs = self.parse_fx_attributes(upi)
        else:
            upi_attrs = self.parse_ir_attributes(upi)
            
        # Score each mapped field
        for field, var in self.mapping_vars.items():
            if var.get() == "N/A":
                continue
                
            weight = self.get_field_weight(field)
            max_score += weight
            
            trade_value = trade.get(field)
            upi_value = upi_attrs.get(field)
            
            if trade_value and upi_value:
                # Special handling for currency pair matching
                if field in ["Notional Currency", "Other Notional Currency"]:
                    if self.currencies_match(trade_value, upi_value):
                        if is_reversed:
                            score += weight * 0.9  # Slightly lower score for reversed match
                        else:
                            score += weight
                else:
                    # Standard string matching
                    if str(trade_value).strip().upper() == str(upi_value).strip().upper():
                        score += weight
                    elif str(trade_value).strip().upper() in str(upi_value).strip().upper():
                        score += weight * 0.8
                        
        return (score / max_score * 100) if max_score > 0 else 0
        
    def currencies_match(self, trade_currency, upi_currency):
        """Check if currencies match (handles various formats)"""
        if not trade_currency or not upi_currency:
            return False
            
        trade_curr = str(trade_currency).strip().upper()
        upi_curr = str(upi_currency).strip().upper()
        
        return trade_curr == upi_curr
        
    def get_field_weight(self, field):
        """Get scoring weight for each field"""
        if self.asset_class.get() == "FX":
            weights = {
                "Asset Class": 20,
                "Instrument Type": 15,
                "Product Type": 20,
                "Notional Currency": 15,
                "Other Notional Currency": 15,
                "Settlement Currency": 10,
                "Delivery Type": 10,
                "Option Type": 8,
                "Option Exercise Style": 8,
                "Valuation Method or Trigger": 5,
                "Underlying Asset Type": 5,
                "Return or Payout Trigger": 5,
                "Place of Settlement": 12
            }
        else:
            weights = {
                "Asset Class": 20,
                "Instrument Type": 15,
                "Product Type": 20,
                "Notional Currency": 10,
                "Reference Rate": 15,
                "Reference Rate Term Value": 5,
                "Reference Rate Term Unit": 5,
                "Other Leg Reference Rate": 10,
                "Other Leg Reference Rate Term Value": 3,
                "Other Leg Reference Rate Term Unit": 3,
                "Other Notional Currency": 8,
                "Notional Schedule": 5,
                "Delivery Type": 8
            }
            
        return weights.get(field, 5)  # Default weight
        
    def get_match_details(self, trade, upi):
        """Get detailed match information"""
        if not upi:
            return "No match found"
            
        details = []
        
        # Parse UPI attributes
        if self.asset_class.get() == "FX":
            upi_attrs = self.parse_fx_attributes(upi)
        else:
            upi_attrs = self.parse_ir_attributes(upi)
            
        # Compare key fields
        key_fields = ["Asset Class", "Instrument Type", "Product Type", "Notional Currency", "Other Notional Currency"]
        
        for field in key_fields:
            trade_value = trade.get(field, "")
            upi_value = upi_attrs.get(field, "")
            
            if trade_value and upi_value:
                match_status = "✓" if str(trade_value).strip().upper() == str(upi_value).strip().upper() else "✗"
                details.append(f"{field}: {match_status} ({trade_value} vs {upi_value})")
                
        return "; ".join(details)
        
    def display_results(self):
        """Display search results in the tree view"""
        # Clear existing results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
            
        if not self.results:
            return
            
        # Configure columns
        columns = ["Trade Index", "Best UPI", "Score", "Match Details", "Is Offshore", "Target Products"]
        self.results_tree["columns"] = columns
        self.results_tree["show"] = "headings"
        
        # Configure column headings and widths
        for col in columns:
            self.results_tree.heading(col, text=col)
            if col == "Match Details":
                self.results_tree.column(col, width=400)
            elif col == "Best UPI":
                self.results_tree.column(col, width=150)
            else:
                self.results_tree.column(col, width=100)
                
        # Insert results
        for result in self.results:
            values = [result.get(col, "") for col in columns]
            self.results_tree.insert("", "end", values=values)
            
    def export_results(self):
        """Export results to Excel"""
        if not self.results:
            messagebox.showwarning("Warning", "No results to export")
            return
            
        try:
            # Create DataFrame from results
            df = pd.DataFrame(self.results)
            
            # Ask user for save location
            filename = filedialog.asksaveasfilename(
                title="Save Results",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if filename:
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Results exported to {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting results: {str(e)}")

def main():
    root = tk.Tk()
    app = UPISearchTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()