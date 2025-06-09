
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import json
import os
import re
from tkinter import scrolledtext
import traceback

class UPISearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("UPI Search Automation Tool")
        self.root.geometry("1000x800")
        
        # Initialize variables
        self.upi_data = None
        self.trade_data = None
        self.upi_file_path = tk.StringVar()
        self.trade_file_path = tk.StringVar()
        self.asset_class = tk.StringVar(value="FX")
        self.mapping_dict = {}
        self.results = []
        
        # Create UI
        self.create_ui()
    
    def create_ui(self):
        # Create a notebook (tabbed interface)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.tab1 = ttk.Frame(notebook)
        self.tab2 = ttk.Frame(notebook)
        self.tab3 = ttk.Frame(notebook)
        
        notebook.add(self.tab1, text="Upload Files")
        notebook.add(self.tab2, text="Map Columns")
        notebook.add(self.tab3, text="Results")
        
        # Tab 1 - File Upload
        self.create_upload_tab()
        
        # Tab 2 - Column Mapping
        self.create_mapping_tab()
        
        # Tab 3 - Results
        self.create_results_tab()
    
    def create_upload_tab(self):
        # UPI File Upload
        upi_frame = ttk.LabelFrame(self.tab1, text="UPI Data (JSON)")
        upi_frame.pack(fill='x', expand=True, padx=10, pady=10)
        
        ttk.Label(upi_frame, text="UPI JSON File:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(upi_frame, textvariable=self.upi_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(upi_frame, text="Browse", command=self.browse_upi_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Trade Data Upload
        trade_frame = ttk.LabelFrame(self.tab1, text="Trade Data (Excel)")
        trade_frame.pack(fill='x', expand=True, padx=10, pady=10)
        
        ttk.Label(trade_frame, text="Trade Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(trade_frame, textvariable=self.trade_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(trade_frame, text="Browse", command=self.browse_trade_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Asset Class Selection
        asset_frame = ttk.LabelFrame(self.tab1, text="Asset Class")
        asset_frame.pack(fill='x', expand=True, padx=10, pady=10)
        
        ttk.Radiobutton(asset_frame, text="FX (Foreign Exchange)", variable=self.asset_class, value="FX").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Radiobutton(asset_frame, text="IR (Interest Rate)", variable=self.asset_class, value="IR").grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Load Data Button
        ttk.Button(self.tab1, text="Load Data", command=self.load_data).pack(pady=20)
        
        # Status display
        self.status_upload = tk.StringVar()
        ttk.Label(self.tab1, textvariable=self.status_upload).pack(pady=5)
    
    def create_mapping_tab(self):
        # This will be populated after loading the files
        self.mapping_frame = ttk.Frame(self.tab2)
        self.mapping_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(self.mapping_frame, text="Please load data files first in the 'Upload Files' tab").pack(pady=20)
        
        # Map Button (initially hidden)
        self.map_button = ttk.Button(self.tab2, text="Map Columns & Search UPIs", command=self.search_upis)
        
        # Status display
        self.status_mapping = tk.StringVar()
        self.status_label_mapping = ttk.Label(self.tab2, textvariable=self.status_mapping)
        self.status_label_mapping.pack(pady=5)
    
    def create_results_tab(self):
        # Results display
        results_frame = ttk.Frame(self.tab3)
        results_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create a scrolled text widget for displaying results
        self.results_text = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, width=90, height=30)
        self.results_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Export button
        self.export_button = ttk.Button(self.tab3, text="Export Results to Excel", command=self.export_results)
        
        # Initially display message
        self.results_text.insert(tk.END, "Results will be displayed here after mapping and searching.")
    
    def browse_upi_file(self):
        filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filename:
            self.upi_file_path.set(filename)
    
    def browse_trade_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename:
            self.trade_file_path.set(filename)
    
    def load_data(self):
        try:
            # Check if files are selected
            if not self.upi_file_path.get() or not self.trade_file_path.get():
                messagebox.showerror("Error", "Please select both UPI JSON file and Trade Excel file")
                return
            
            # Load UPI data
            with open(self.upi_file_path.get(), 'r') as f:
                self.upi_data = json.load(f)
            
            # Load trade data
            self.trade_data = pd.read_excel(self.trade_file_path.get())
            
            # Update status
            self.status_upload.set(f"Files loaded successfully. UPI records: {len(self.upi_data.get('upis', []))} | Trade records: {len(self.trade_data)}")
            
            # Create mapping UI
            self.create_mapping_ui()
            
            # Switch to mapping tab
            self.root.focus_force()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading files: {str(e)}")
            self.status_upload.set(f"Error: {str(e)}")
    
    def create_mapping_ui(self):
        # Clear existing widgets in mapping frame
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        
        # Create a notebook for different mapping sections
        mapping_notebook = ttk.Notebook(self.mapping_frame)
        mapping_notebook.pack(fill='both', expand=True)
        
        # Create mapping tabs based on asset class
        if self.asset_class.get() == "FX":
            # Create FX mapping tab
            fx_tab = ttk.Frame(mapping_notebook)
            mapping_notebook.add(fx_tab, text="FX Mapping")
            
            # Create mapping widgets for FX
            self.create_fx_mapping_widgets(fx_tab)
            
        elif self.asset_class.get() == "IR":
            # Create IR mapping tab
            ir_tab = ttk.Frame(mapping_notebook)
            mapping_notebook.add(ir_tab, text="IR Mapping")
            
            # Create mapping widgets for IR
            self.create_ir_mapping_widgets(ir_tab)
        
        # Show map button
        self.map_button.pack(pady=10)
    
    def create_fx_mapping_widgets(self, parent):
        # Display available columns from trade data
        ttk.Label(parent, text="Map your Excel columns to UPI search attributes").grid(row=0, column=0, columnspan=3, pady=10)
        
        # Create mapping fields for FX
        row = 1
        
        # Create all mapping fields
        mapping_fields = [
            ("Asset Class", "assetClass"),
            ("Instrument Type", "instrumentType"),
            ("Product Type", "product"),
            ("Currency Pair", "currencyPair"),
            ("Settlement Currency", "settlementCurrency"),
            ("Option Type", "optionType"),
            ("Option Style", "optionStyle"),
            ("Delivery Type", "deliveryType")
        ]
        
        # Dictionary to store the mapping variables
        self.mapping_vars = {}
        
        for label, field_name in mapping_fields:
            ttk.Label(parent, text=label + ":").grid(row=row, column=0, padx=5, pady=5, sticky='w')
            
            # Create variable and dropdown for mapping
            var = tk.StringVar()
            self.mapping_vars[field_name] = var
            
            # Add a "Not Applicable" option
            columns = list(self.trade_data.columns) + ["N/A"]
            
            # Try to auto-select a matching column
            auto_select = self.find_matching_column(label, columns)
            if auto_select:
                var.set(auto_select)
            else:
                var.set(columns[0])
            
            dropdown = ttk.Combobox(parent, textvariable=var, values=columns, width=30)
            dropdown.grid(row=row, column=1, padx=5, pady=5)
            
            row += 1
    
    def create_ir_mapping_widgets(self, parent):
        # Display available columns from trade data
        ttk.Label(parent, text="Map your Excel columns to UPI search attributes").grid(row=0, column=0, columnspan=3, pady=10)
        
        # Create mapping fields for IR
        row = 1
        
        # Create all mapping fields
        mapping_fields = [
            ("Asset Class", "assetClass"),
            ("Instrument Type", "instrumentType"),
            ("Product Type", "product"),
            ("Reference Rate", "referenceRate"),
            ("Currency", "currency"),
            ("Term", "term"),
            ("Other Leg Reference Rate", "otherLegReferenceRate"),
            ("Other Leg Currency", "otherLegCurrency"),
            ("Other Leg Term", "otherLegTerm"),
            ("Delivery Type", "deliveryType")
        ]
        
        # Dictionary to store the mapping variables
        self.mapping_vars = {}
        
        for label, field_name in mapping_fields:
            ttk.Label(parent, text=label + ":").grid(row=row, column=0, padx=5, pady=5, sticky='w')
            
            # Create variable and dropdown for mapping
            var = tk.StringVar()
            self.mapping_vars[field_name] = var
            
            # Add a "Not Applicable" option
            columns = list(self.trade_data.columns) + ["N/A"]
            
            # Try to auto-select a matching column
            auto_select = self.find_matching_column(label, columns)
            if auto_select:
                var.set(auto_select)
            else:
                var.set(columns[0])
            
            dropdown = ttk.Combobox(parent, textvariable=var, values=columns, width=30)
            dropdown.grid(row=row, column=1, padx=5, pady=5)
            
            row += 1
    
    def find_matching_column(self, label, columns):
        """Try to automatically match a label to a column name"""
        # Remove spaces and convert to lowercase for comparison
        label_simple = label.lower().replace(" ", "")
        
        for col in columns:
            col_simple = col.lower().replace(" ", "")
            
            # Check for exact match or partial match
            if col_simple == label_simple or label_simple in col_simple or col_simple in label_simple:
                return col
            
            # Check for common abbreviations and synonyms
            if label == "Asset Class" and ("asset" in col_simple or "class" in col_simple or col_simple == "assetclass"):
                return col
            elif label == "Instrument Type" and ("instrument" in col_simple or "type" in col_simple or "producttype" in col_simple):
                return col
            elif label == "Product Type" and ("product" in col_simple or "swaptype" in col_simple or "forwardtype" in col_simple):
                return col
            elif label == "Currency Pair" and ("ccy" in col_simple or "pair" in col_simple or "ccypair" in col_simple):
                return col
            elif label == "Reference Rate" and ("ref" in col_simple or "rate" in col_simple or "refrate" in col_simple):
                return col
            elif label == "Currency" and ("currency" in col_simple or "ccy" in col_simple or col_simple == "ccy1" or col_simple == "currency1"):
                return col
            elif label == "Term" and ("term" in col_simple or "tenor" in col_simple or col_simple == "term1"):
                return col
            elif label == "Other Leg Reference Rate" and ("other" in col_simple or "leg" in col_simple or "refrate2" in col_simple):
                return col
            elif label == "Other Leg Currency" and ("other" in col_simple or "leg" in col_simple or "currency2" in col_simple or "ccy2" in col_simple):
                return col
            elif label == "Other Leg Term" and ("other" in col_simple or "leg" in col_simple or "term2" in col_simple):
                return col
            elif label == "Delivery Type" and ("delivery" in col_simple or "settlement" in col_simple or "settlementtype" in col_simple):
                return col
            elif label == "Option Type" and ("option" in col_simple and "type" in col_simple):
                return col
            elif label == "Option Style" and ("option" in col_simple and "style" in col_simple):
                return col
                
        return None
    
    def search_upis(self):
        try:
            # Clear previous results
            self.results = []
            self.results_text.delete(1.0, tk.END)
            
            # Get mapping from UI
            mapping = {field: var.get() for field, var in self.mapping_vars.items()}
            
            # Process each trade
            for index, trade in self.trade_data.iterrows():
                result = self.find_matching_upi(trade, mapping)
                self.results.append(result)
            
            # Display results
            self.display_results()
            
            # Show export button
            self.export_button.pack(pady=10)
            
            # Update status
            self.status_mapping.set(f"UPI search completed. {len(self.results)} trades processed.")
            
            # Switch to results tab
            notebook = self.tab3.master
            notebook.select(2)  # Select the third tab (index 2)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error searching UPIs: {str(e)}\n{traceback.format_exc()}")
            self.status_mapping.set(f"Error: {str(e)}")
    
    def find_matching_upi(self, trade, mapping):
        result = {"TradeDetails": trade.to_dict(), "MatchedUPI": None, "Score": 0, "Message": ""}
        
        # Get UPI data
        upis = self.upi_data.get("upis", [])
        if not upis:
            result["Message"] = "No UPI records found in the UPI data"
            return result
        
        # Initialize asset class specific search criteria
        if self.asset_class.get() == "FX":
            search_result = self.search_fx_upi(trade, mapping, upis)
        else:  # IR
            search_result = self.search_ir_upi(trade, mapping, upis)
        
        # Update result with search results
        result.update(search_result)
        
        return result
    
    def search_fx_upi(self, trade, mapping, upis):
        result = {"MatchedUPI": None, "Score": 0, "Message": ""}
        
        # Extract values from trade data based on mapping
        try:
            # Asset class could be "FX" or "ForeignExchange" in different systems
            asset_class_col = mapping["assetClass"]
            asset_class = trade[asset_class_col] if asset_class_col != "N/A" else None
            if asset_class and "fx" in asset_class.lower():
                asset_class = "ForeignExchange"
            
            instrument_type_col = mapping["instrumentType"]
            instrument_type = trade[instrument_type_col] if instrument_type_col != "N/A" else None
            
            product_col = mapping["product"]
            product = trade[product_col] if product_col != "N/A" else None
            
            ccy_pair_col = mapping["currencyPair"]
            ccy_pair = trade[ccy_pair_col] if ccy_pair_col != "N/A" else None
            
            settlement_ccy_col = mapping["settlementCurrency"]
            settlement_ccy = trade[settlement_ccy_col] if settlement_ccy_col != "N/A" else None
            
            option_type_col = mapping["optionType"]
            option_type = trade[option_type_col] if option_type_col != "N/A" else None
            
            option_style_col = mapping["optionStyle"]
            option_style = trade[option_style_col] if option_style_col != "N/A" else None
            
            delivery_type_col = mapping["deliveryType"]
            delivery_type = trade[delivery_type_col] if delivery_type_col != "N/A" else None
            
            # Match logic for FX UPIs
            best_match = None
            best_score = 0
            
            for upi in upis:
                # Initialize score for this UPI
                score = 0
                
                # Check asset class match
                if asset_class and upi.get("assetClass") == asset_class:
                    score += 20
                elif asset_class and (
                    ("fx" in asset_class.lower() and upi.get("assetClass") == "ForeignExchange") or
                    (asset_class.lower() == "foreignexchange" and "fx" in upi.get("assetClass", "").lower())
                ):
                    score += 15  # Partial match
                
                # Check instrument type match
                if instrument_type and upi.get("instrumentType") == instrument_type:
                    score += 20
                
                # Check product match
                if product and upi.get("product") == product:
                    score += 20
                
                # Check currency pair match
                if ccy_pair and upi.get("underlying", {}).get("currencyPair") == ccy_pair:
                    score += 20
                
                # Check settlement currency match
                if settlement_ccy and upi.get("underlying", {}).get("settlementCurrency") == settlement_ccy:
                    score += 10
                
                # Check option type match for options
                if option_type and upi.get("optionType") == option_type:
                    score += 5
                
                # Check option style match for options
                if option_style and upi.get("optionStyle") == option_style:
                    score += 5
                
                # Check delivery type match
                if delivery_type and upi.get("deliveryType") == delivery_type:
                    score += 10
                elif delivery_type and (
                    (delivery_type.lower() == "cash" and upi.get("deliveryType") == "Cash") or
                    (delivery_type.lower() == "physical" and upi.get("deliveryType") == "Physical")
                ):
                    score += 8  # Case-insensitive match
                
                # Update best match if this UPI has a higher score
                if score > best_score:
                    best_score = score
                    best_match = upi
            
            # Set result based on best match
            if best_match and best_score >= 50:  # Require at least 50% match
                result["MatchedUPI"] = best_match
                result["Score"] = best_score
                result["Message"] = "UPI found with match score: " + str(best_score)
            else:
                result["Message"] = "No matching UPI found with sufficient confidence"
            
        except Exception as e:
            result["Message"] = f"Error during FX UPI search: {str(e)}"
        
        return result
    
    def search_ir_upi(self, trade, mapping, upis):
        result = {"MatchedUPI": None, "Score": 0, "Message": ""}
        
        # Extract values from trade data based on mapping
        try:
            # Asset class could be "IR", "Rates" or "InterestRate" in different systems
            asset_class_col = mapping["assetClass"]
            asset_class = trade[asset_class_col] if asset_class_col != "N/A" else None
            if asset_class and "ir" in asset_class.lower():
                asset_class = "Rates"
            
            instrument_type_col = mapping["instrumentType"]
            instrument_type = trade[instrument_type_col] if instrument_type_col != "N/A" else None
            
            product_col = mapping["product"]
            product = trade[product_col] if product_col != "N/A" else None
            
            ref_rate_col = mapping["referenceRate"]
            ref_rate = trade[ref_rate_col] if ref_rate_col != "N/A" else None
            
            currency_col = mapping["currency"]
            currency = trade[currency_col] if currency_col != "N/A" else None
            
            term_col = mapping["term"]
            term = trade[term_col] if term_col != "N/A" else None
            
            other_ref_rate_col = mapping["otherLegReferenceRate"]
            other_ref_rate = trade[other_ref_rate_col] if other_ref_rate_col != "N/A" else None
            
            other_currency_col = mapping["otherLegCurrency"]
            other_currency = trade[other_currency_col] if other_currency_col != "N/A" else None
            
            other_term_col = mapping["otherLegTerm"]
            other_term = trade[other_term_col] if other_term_col != "N/A" else None
            
            delivery_type_col = mapping["deliveryType"]
            delivery_type = trade[delivery_type_col] if delivery_type_col != "N/A" else None
            
            # Match logic for IR UPIs
            best_match = None
            best_score = 0
            
            for upi in upis:
                # Initialize score for this UPI
                score = 0
                
                # Check asset class match
                if asset_class and upi.get("assetClass") == asset_class:
                    score += 20
                elif asset_class and (
                    ("ir" in asset_class.lower() and upi.get("assetClass") == "Rates") or
                    (asset_class.lower() == "rates" and "ir" in upi.get("assetClass", "").lower())
                ):
                    score += 15  # Partial match
                
                # Check instrument type match
                if instrument_type and upi.get("instrumentType") == instrument_type:
                    score += 20
                
                # Check product match
                if product and upi.get("product") == product:
                    score += 20
                
                # Check reference rate match
                if ref_rate and upi.get("underlying", {}).get("referenceRate") == ref_rate:
                    score += 15
                
                # Check currency match
                if currency and upi.get("underlying", {}).get("currency") == currency:
                    score += 10
                
                # Check term match
                if term and upi.get("underlying", {}).get("term") == term:
                    score += 10
                
                # Check other leg reference rate match
                if other_ref_rate and upi.get("otherLeg", {}).get("referenceRate") == other_ref_rate:
                    score += 10
                
                # Check other leg currency match
                if other_currency and upi.get("otherLeg", {}).get("currency") == other_currency:
                    score += 5
                
                # Check other leg term match
                if other_term and upi.get("otherLeg", {}).get("term") == other_term:
                    score += 5
                
                # Check delivery type match
                if delivery_type and upi.get("deliveryType") == delivery_type:
                    score += 5
                
                # Update best match if this UPI has a higher score
                if score > best_score:
                    best_score = score
                    best_match = upi
            
            # Set result based on best match
            if best_match and best_score >= 50:  # Require at least 50% match
                result["MatchedUPI"] = best_match
                result["Score"] = best_score
                result["Message"] = "UPI found with match score: " + str(best_score)
            else:
                result["Message"] = "No matching UPI found with sufficient confidence"
            
        except Exception as e:
            result["Message"] = f"Error during IR UPI search: {str(e)}"
        
        return result
    
    def display_results(self):
        self.results_text.delete(1.0, tk.END)
        
        # Write header
        self.results_text.insert(tk.END, "UPI Search Results\n")
        self.results_text.insert(tk.END, "=" * 80 + "\n\n")
        
        # Display results for each trade
        for i, result in enumerate(self.results):
            trade_details = result["TradeDetails"]
            matched_upi = result["MatchedUPI"]
            score = result["Score"]
            message = result["Message"]
            
            # Trade details
            self.results_text.insert(tk.END, f"Trade {i+1}:\n")
            
            # Get trade ID if available
            trade_id = next((trade_details[col] for col in trade_details if "id" in col.lower()), f"Trade {i+1}")
            self.results_text.insert(tk.END, f"Trade ID: {trade_id}\n")
            
            # Display key trade details
            self.results_text.insert(tk.END, "Key Trade Details:\n")
            for key, value in trade_details.items():
                if value and pd.notna(value) and not (isinstance(value, float) and pd.isna(value)):
                    self.results_text.insert(tk.END, f"  - {key}: {value}\n")
            
            # UPI match result
            self.results_text.insert(tk.END, f"Match Score: {score}\n")
            self.results_text.insert(tk.END, f"Message: {message}\n")
            
            if matched_upi:
                self.results_text.insert(tk.END, f"Matched UPI Code: {matched_upi.get('upiCode', 'N/A')}\n")
                self.results_text.insert(tk.END, "UPI Details:\n")
                
                # Display UPI details
                self.results_text.insert(tk.END, f"  - Asset Class: {matched_upi.get('assetClass', 'N/A')}\n")
                self.results_text.insert(tk.END, f"  - Instrument Type: {matched_upi.get('instrumentType', 'N/A')}\n")
                self.results_text.insert(tk.END, f"  - Product: {matched_upi.get('product', 'N/A')}\n")
                
                # Asset class specific details
                if matched_upi.get('assetClass') == "ForeignExchange":
                    underlying = matched_upi.get('underlying', {})
                    self.results_text.insert(tk.END, f"  - Currency Pair: {underlying.get('currencyPair', 'N/A')}\n")
                    if 'settlementCurrency' in underlying:
                        self.results_text.insert(tk.END, f"  - Settlement Currency: {underlying.get('settlementCurrency', 'N/A')}\n")
                    if 'optionType' in matched_upi:
                        self.results_text.insert(tk.END, f"  - Option Type: {matched_upi.get('optionType', 'N/A')}\n")
                    if 'optionStyle' in matched_upi:
                        self.results_text.insert(tk.END, f"  - Option Style: {matched_upi.get('optionStyle', 'N/A')}\n")
                
                elif matched_upi.get('assetClass') == "Rates":
                    underlying = matched_upi.get('underlying', {})
                    self.results_text.insert(tk.END, f"  - Reference Rate: {underlying.get('referenceRate', 'N/A')}\n")
                    self.results_text.insert(tk.END, f"  - Currency: {underlying.get('currency', 'N/A')}\n")
                    self.results_text.insert(tk.END, f"  - Term: {underlying.get('term', 'N/A')}\n")
                    
                    # Other leg details if applicable
                    if 'otherLeg' in matched_upi:
                        other_leg = matched_upi.get('otherLeg', {})
                        self.results_text.insert(tk.END, f"  - Other Leg Reference Rate: {other_leg.get('referenceRate', 'N/A')}\n")
                        self.results_text.insert(tk.END, f"  - Other Leg Currency: {other_leg.get('currency', 'N/A')}\n")
                        self.results_text.insert(tk.END, f"  - Other Leg Term: {other_leg.get('term', 'N/A')}\n")
                
                self.results_text.insert(tk.END, f"  - Delivery Type: {matched_upi.get('deliveryType', 'N/A')}\n")
            
            # Separator between trades
            self.results_text.insert(tk.END, "\n" + "-" * 80 + "\n\n")
    
    def export_results(self):
        if not self.results:
            messagebox.showinfo("Export Results", "No results to export.")
            return
        
        try:
            # Ask for save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if not file_path:
                return  # User cancelled
            
            # Prepare data for export
            export_data = []
            
            for result in self.results:
                trade_details = result["TradeDetails"]
                matched_upi = result["MatchedUPI"]
                
                row = {}
                
                # Add trade details
                for key, value in trade_details.items():
                    row[f"Trade_{key}"] = value
                
                # Add UPI match details
                row["Match_Score"] = result["Score"]
                row["Match_Message"] = result["Message"]
                
                if matched_upi:
                    row["UPI_Code"] = matched_upi.get("upiCode", "")
                    row["UPI_AssetClass"] = matched_upi.get("assetClass", "")
                    row["UPI_InstrumentType"] = matched_upi.get("instrumentType", "")
                    row["UPI_Product"] = matched_upi.get("product", "")
                    row["UPI_DeliveryType"] = matched_upi.get("deliveryType", "")
                    
                    # Asset class specific details
                    if matched_upi.get('assetClass') == "ForeignExchange":
                        underlying = matched_upi.get('underlying', {})
                        row["UPI_CurrencyPair"] = underlying.get('currencyPair', "")
                        row["UPI_SettlementCurrency"] = underlying.get('settlementCurrency', "")
                        row["UPI_OptionType"] = matched_upi.get('optionType', "")
                        row["UPI_OptionStyle"] = matched_upi.get('optionStyle', "")
                    
                    elif matched_upi.get('assetClass') == "Rates":
                        underlying = matched_upi.get('underlying', {})
                        row["UPI_ReferenceRate"] = underlying.get('referenceRate', "")
                        row["UPI_Currency"] = underlying.get('currency', "")
                        row["UPI_Term"] = underlying.get('term', "")
                        
                        if 'otherLeg' in matched_upi:
                            other_leg = matched_upi.get('otherLeg', {})
                            row["UPI_OtherLegReferenceRate"] = other_leg.get('referenceRate', "")
                            row["UPI_OtherLegCurrency"] = other_leg.get('currency', "")
                            row["UPI_OtherLegTerm"] = other_leg.get('term', "")
                
                export_data.append(row)
            
            # Create DataFrame and export to Excel
            df = pd.DataFrame(export_data)
            df.to_excel(file_path, index=False)
            
            messagebox.showinfo("Export Results", f"Results exported successfully to {file_path}")
        
        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting results: {str(e)}")

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = UPISearchTool(root)
    root.mainloop()
