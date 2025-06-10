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
        self.root.title("UPI Search Automation Tool - DSB Schema")
        self.root.geometry("1200x900")
        
        # Initialize variables
        self.upi_data = None
        self.trade_data = None
        self.upi_file_path = tk.StringVar()
        self.trade_file_path = tk.StringVar()
        self.asset_class = tk.StringVar(value="FX")
        self.product_type = tk.StringVar()
        self.mapping_dict = {}
        self.results = []
        self.available_products = []
        
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
        self.tab4 = ttk.Frame(notebook)
        
        notebook.add(self.tab1, text="Upload Files")
        notebook.add(self.tab2, text="Select Product")
        notebook.add(self.tab3, text="Map Columns")
        notebook.add(self.tab4, text="Results")
        
        # Tab 1 - File Upload
        self.create_upload_tab()
        
        # Tab 2 - Product Selection
        self.create_product_selection_tab()
        
        # Tab 3 - Column Mapping
        self.create_mapping_tab()
        
        # Tab 4 - Results
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
    
    def create_product_selection_tab(self):
        # Product selection frame
        product_frame = ttk.LabelFrame(self.tab2, text="Select Product Type")
        product_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Instructions
        ttk.Label(product_frame, text="Please load data files first, then select the product type for your trades:").pack(pady=10)
        
        # Product selection dropdown
        self.product_label = ttk.Label(product_frame, text="Product Type:")
        self.product_dropdown = ttk.Combobox(product_frame, textvariable=self.product_type, width=40, state="readonly")
        
        # Continue button
        self.continue_button = ttk.Button(self.tab2, text="Continue to Column Mapping", command=self.proceed_to_mapping)
        
        # Status display
        self.status_product = tk.StringVar()
        ttk.Label(self.tab2, textvariable=self.status_product).pack(pady=5)
    
    def create_mapping_tab(self):
        # This will be populated after product selection
        self.mapping_frame = ttk.Frame(self.tab3)
        self.mapping_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(self.mapping_frame, text="Please select product type first in the 'Select Product' tab").pack(pady=20)
        
        # Map Button (initially hidden)
        self.map_button = ttk.Button(self.tab3, text="Map Columns & Search UPIs", command=self.search_upis)
        
        # Status display
        self.status_mapping = tk.StringVar()
        self.status_label_mapping = ttk.Label(self.tab3, textvariable=self.status_mapping)
        self.status_label_mapping.pack(pady=5)
    
    def create_results_tab(self):
        # Results display
        results_frame = ttk.Frame(self.tab4)
        results_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create a scrolled text widget for displaying results
        self.results_text = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, width=100, height=35)
        self.results_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Export button
        self.export_button = ttk.Button(self.tab4, text="Export Results to Excel", command=self.export_results)
        
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
            
            # Extract available products based on asset class
            self.extract_available_products()
            
            # Update status
            upi_count = len(self.upi_data) if isinstance(self.upi_data, list) else 1
            self.status_upload.set(f"Files loaded successfully. UPI records: {upi_count} | Trade records: {len(self.trade_data)}")
            
            # Setup product selection
            self.setup_product_selection()
            
            # Switch to product selection tab
            notebook = self.tab2.master
            notebook.select(1)  # Select the second tab (index 1)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading files: {str(e)}")
            self.status_upload.set(f"Error: {str(e)}")
    
    def extract_available_products(self):
        """Extract available product types from the loaded UPI data based on asset class"""
        self.available_products = []
        
        try:
            # Handle different UPI data structures
            if isinstance(self.upi_data, dict):
                # Single UPI record
                upi_records = [self.upi_data]
            elif isinstance(self.upi_data, list):
                # Multiple UPI records
                upi_records = self.upi_data
            else:
                raise ValueError("Unexpected UPI data format")
            
            # Extract products based on asset class
            asset_class_filter = "Foreign_Exchange" if self.asset_class.get() == "FX" else "Rates"
            
            products = set()
            for upi in upi_records:
                header = upi.get("Header", {})
                if header.get("AssetClass") == asset_class_filter:
                    use_case = header.get("UseCase")
                    if use_case:
                        products.add(use_case)
            
            self.available_products = sorted(list(products))
            
        except Exception as e:
            messagebox.showerror("Error", f"Error extracting products: {str(e)}")
            self.available_products = []
    
    def setup_product_selection(self):
        """Setup the product selection dropdown"""
        if self.available_products:
            self.product_label.pack(pady=5)
            self.product_dropdown['values'] = self.available_products
            self.product_dropdown.pack(pady=5)
            self.continue_button.pack(pady=20)
            
            # Auto-select first product if only one available
            if len(self.available_products) == 1:
                self.product_type.set(self.available_products[0])
            
            self.status_product.set(f"Found {len(self.available_products)} product types for {self.asset_class.get()}")
        else:
            self.status_product.set("No products found for the selected asset class")
    
    def proceed_to_mapping(self):
        """Proceed to column mapping after product selection"""
        if not self.product_type.get():
            messagebox.showerror("Error", "Please select a product type")
            return
        
        # Create mapping UI based on selected product
        self.create_mapping_ui()
        
        # Switch to mapping tab
        notebook = self.tab3.master
        notebook.select(2)  # Select the third tab (index 2)
    
    def create_mapping_ui(self):
        """Create mapping UI based on selected asset class and product"""
        # Clear existing widgets in mapping frame
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        
        # Create scrollable frame for mapping
        canvas = tk.Canvas(self.mapping_frame)
        scrollbar = ttk.Scrollbar(self.mapping_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Get mapping fields based on asset class and product
        mapping_fields = self.get_mapping_fields()
        
        # Create header
        ttk.Label(scrollable_frame, text=f"Map your Excel columns to UPI search attributes", 
                 font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=3, pady=10)
        ttk.Label(scrollable_frame, text=f"Asset Class: {self.asset_class.get()} | Product: {self.product_type.get()}", 
                 font=("Arial", 10)).grid(row=1, column=0, columnspan=3, pady=5)
        
        # Dictionary to store the mapping variables
        self.mapping_vars = {}
        
        row = 2
        for label, field_name, required in mapping_fields:
            # Label with required indicator
            label_text = label + (" *" if required else "")
            ttk.Label(scrollable_frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky='w')
            
            # Create variable and dropdown for mapping
            var = tk.StringVar()
            self.mapping_vars[field_name] = var
            
            # Add columns and "N/A" option
            columns = list(self.trade_data.columns) + ["N/A"]
            
            # Try to auto-select a matching column
            auto_select = self.find_matching_column(label, columns)
            if auto_select:
                var.set(auto_select)
            else:
                var.set("N/A" if not required else columns[0])
            
            dropdown = ttk.Combobox(scrollable_frame, textvariable=var, values=columns, width=40)
            dropdown.grid(row=row, column=1, padx=5, pady=5)
            
            # Add description
            description = self.get_field_description(field_name)
            ttk.Label(scrollable_frame, text=description, font=("Arial", 8), foreground="gray").grid(row=row, column=2, padx=5, pady=5, sticky='w')
            
            row += 1
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Show map button
        self.map_button.pack(pady=10)
    
    def get_mapping_fields(self):
        """Get mapping fields based on asset class and product type"""
        asset_class = self.asset_class.get()
        product = self.product_type.get()
        
        if asset_class == "FX":
            return self.get_fx_mapping_fields(product)
        else:  # IR
            return self.get_ir_mapping_fields(product)
    
    def get_fx_mapping_fields(self, product):
        """Get FX mapping fields based on product type"""
        # Common FX fields
        fields = [
            ("Notional Currency", "NotionalCurrency", True),
            ("Other Notional Currency", "OtherNotionalCurrency", True),
        ]
        
        # Product-specific fields
        if product == "Forward":
            fields.extend([
                ("Delivery Type", "DeliveryType", True),
            ])
        elif product == "NDF":
            fields.extend([
                ("Settlement Currency", "SettlementCurrency", True),
            ])
        elif product == "Non_Standard":
            fields.extend([
                ("Settlement Currency", "SettlementCurrency", False),
                ("Underlying Asset Type", "UnderlyingAssetType", True),
                ("Return or Payout Trigger", "ReturnorPayoutTrigger", True),
                ("Delivery Type", "DeliveryType", True),
                ("Place of Settlement", "PlaceofSettlement", False),
            ])
        elif product in ["Digital_Option", "Vanilla_Option"]:
            fields.extend([
                ("Option Type", "OptionType", True),
                ("Option Exercise Style", "OptionExerciseStyle", True),
                ("Delivery Type", "DeliveryType", True),
            ])
            if product == "Digital_Option":
                fields.extend([
                    ("Valuation Method or Trigger", "ValuationMethodorTrigger", True),
                    ("Settlement Currency", "SettlementCurrency", True),
                ])
        elif product == "FX_Swap":
            fields.extend([
                ("Delivery Type", "DeliveryType", True),
            ])
        
        return fields
    
    def get_ir_mapping_fields(self, product):
        """Get IR mapping fields based on product type"""
        # Common IR fields
        fields = [
            ("Notional Currency", "NotionalCurrency", True),
            ("Reference Rate", "ReferenceRate", True),
            ("Reference Rate Term Value", "ReferenceRateTermValue", True),
            ("Reference Rate Term Unit", "ReferenceRateTermUnit", True),
            ("Notional Schedule", "NotionalSchedule", True),
            ("Delivery Type", "DeliveryType", True),
        ]
        
        # Product-specific fields
        if product in ["Basis", "Basis_OIS", "Cross_Currency_Basis"]:
            fields.extend([
                ("Other Leg Reference Rate", "OtherLegReferenceRate", True),
                ("Other Leg Reference Rate Term Value", "OtherLegReferenceRateTermValue", True),
                ("Other Leg Reference Rate Term Unit", "OtherLegReferenceRateTermUnit", True),
            ])
        
        if product.startswith("Cross_Currency"):
            fields.extend([
                ("Other Notional Currency", "OtherNotionalCurrency", True),
            ])
        
        return fields
    
    def get_field_description(self, field_name):
        """Get description for mapping fields"""
        descriptions = {
            "NotionalCurrency": "Currency in which the notional is denominated (e.g., USD, EUR)",
            "OtherNotionalCurrency": "Currency for leg 2 in cross-currency contracts",
            "DeliveryType": "CASH or PHYS",
            "SettlementCurrency": "Currency for settlement",
            "UnderlyingAssetType": "Spot, Forward, Options, or Futures",
            "ReturnorPayoutTrigger": "Payout mechanism",
            "PlaceofSettlement": "Country/location of settlement",
            "OptionType": "CALL, PUTO, or OPTL",
            "OptionExerciseStyle": "AMER, BERM, or EURO",
            "ValuationMethodorTrigger": "Digital (Binary) or Digital Barrier",
            "ReferenceRate": "Reference rate identifier (e.g., USD-LIBOR-3M)",
            "ReferenceRateTermValue": "Numeric value for term (e.g., 3 for 3M)",
            "ReferenceRateTermUnit": "DAYS, WEEK, MNTH, or YEAR",
            "NotionalSchedule": "Constant, Accreting, Amortizing, or Custom",
            "OtherLegReferenceRate": "Reference rate for second leg",
            "OtherLegReferenceRateTermValue": "Term value for second leg",
            "OtherLegReferenceRateTermUnit": "Term unit for second leg",
        }
        return descriptions.get(field_name, "")
    
    def find_matching_column(self, label, columns):
        """Try to automatically match a label to a column name"""
        # Remove spaces and convert to lowercase for comparison
        label_simple = label.lower().replace(" ", "").replace("*", "")
        
        for col in columns:
            if col == "N/A":
                continue
                
            col_simple = col.lower().replace(" ", "").replace("_", "")
            
            # Check for exact match or partial match
            if col_simple == label_simple or label_simple in col_simple or col_simple in label_simple:
                return col
            
            # Check for common abbreviations and synonyms
            if "currency" in label.lower() and ("ccy" in col_simple or "currency" in col_simple):
                return col
            elif "delivery" in label.lower() and ("delivery" in col_simple or "settlement" in col_simple):
                return col
            elif "reference" in label.lower() and "rate" in label.lower() and ("ref" in col_simple or "rate" in col_simple):
                return col
            elif "term" in label.lower() and ("term" in col_simple or "tenor" in col_simple):
                return col
            elif "option" in label.lower() and "type" in label.lower() and ("option" in col_simple and "type" in col_simple):
                return col
            elif "option" in label.lower() and "style" in label.lower() and ("style" in col_simple or "exercise" in col_simple):
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
            notebook = self.tab4.master
            notebook.select(3)  # Select the fourth tab (index 3)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error searching UPIs: {str(e)}\n{traceback.format_exc()}")
            self.status_mapping.set(f"Error: {str(e)}")
    
    def find_matching_upi(self, trade, mapping):
        result = {"TradeDetails": trade.to_dict(), "MatchedUPI": None, "Score": 0, "Message": ""}
        
        try:
            # Handle different UPI data structures
            if isinstance(self.upi_data, dict):
                upi_records = [self.upi_data]
            elif isinstance(self.upi_data, list):
                upi_records = self.upi_data
            else:
                result["Message"] = "Invalid UPI data format"
                return result
            
            # Filter UPIs by asset class and product
            asset_class_filter = "Foreign_Exchange" if self.asset_class.get() == "FX" else "Rates"
            product_filter = self.product_type.get()
            
            relevant_upis = []
            for upi in upi_records:
                header = upi.get("Header", {})
                if (header.get("AssetClass") == asset_class_filter and 
                    header.get("UseCase") == product_filter):
                    relevant_upis.append(upi)
            
            if not relevant_upis:
                result["Message"] = f"No UPI records found for {asset_class_filter} {product_filter}"
                return result
            
            # Perform matching
            best_match = None
            best_score = 0
            
            for upi in relevant_upis:
                score = self.calculate_upi_score(trade, mapping, upi)
                if score > best_score:
                    best_score = score
                    best_match = upi
            
            # Set result based on best match
            if best_match and best_score >= 50:  # Require at least 50% match
                result["MatchedUPI"] = best_match
                result["Score"] = best_score
                result["Message"] = f"UPI found with match score: {best_score}"
            else:
                result["Message"] = f"No matching UPI found with sufficient confidence (best score: {best_score})"
            
        except Exception as e:
            result["Message"] = f"Error during UPI search: {str(e)}"
        
        return result
    
    def calculate_upi_score(self, trade, mapping, upi):
        """Calculate matching score between trade and UPI"""
        score = 0
        max_score = 0
        
        # Get UPI attributes
        attributes = upi.get("Attributes", {})
        
        # Score each mapped field
        for field_name, column_name in mapping.items():
            if column_name == "N/A" or column_name not in trade:
                continue
            
            trade_value = trade[column_name]
            if pd.isna(trade_value) or trade_value == "":
                continue
            
            # Get UPI value for this field
            upi_value = attributes.get(field_name)
            if upi_value is None:
                continue
            
            # Calculate field score
            field_score = self.calculate_field_score(field_name, trade_value, upi_value)
            score += field_score
            max_score += self.get_field_weight(field_name)
        
        # Return percentage score
        return int((score / max_score * 100)) if max_score > 0 else 0
    
    def calculate_field_score(self, field_name, trade_value, upi_value):
        """Calculate score for a specific field match"""
        weight = self.get_field_weight(field_name)
        
        # Convert to strings for comparison
        trade_str = str(trade_value).strip().upper()
        upi_str = str(upi_value).strip().upper()
        
        # Exact match
        if trade_str == upi_str:
            return weight
        
        # Partial matches for specific fields
        if field_name in ["DeliveryType"]:
            if ("CASH" in trade_str and "CASH" in upi_str) or ("PHYS" in trade_str and "PHYS" in upi_str):
                return weight * 0.8
        
        # Currency code matches (handle different formats)
        if "Currency" in field_name:
            if len(trade_str) == 3 and len(upi_str) == 3 and trade_str == upi_str:
                return weight
        
        # Reference rate partial matches
        if "ReferenceRate" in field_name:
            if trade_str in upi_str or upi_str in trade_str:
                return weight * 0.7
        
        return 0
    
    def get_field_weight(self, field_name):
        """Get weight for different fields"""
        weights = {
            "NotionalCurrency": 25,
            "OtherNotionalCurrency": 20,
            "ReferenceRate": 20,
            "ReferenceRateTermValue": 15,
            "ReferenceRateTermUnit": 10,
            "OtherLegReferenceRate": 15,
            "OtherLegReferenceRateTermValue": 10,
            "OtherLegReferenceRateTermUnit": 5,
            "DeliveryType": 15,
            "SettlementCurrency": 10,
            "OptionType": 15,
            "OptionExerciseStyle": 10,
            "ValuationMethodorTrigger": 10,
            "NotionalSchedule": 10,
            "UnderlyingAssetType": 10,
            "ReturnorPayoutTrigger": 10,
            "PlaceofSettlement": 5,
        }
        return weights.get(field_name, 5)
    
    def display_results(self):
        self.results_text.delete(1.0, tk.END)
        
        # Write header
        self.results_text.insert(tk.END, "UPI Search Results\n")
        self.results_text.insert(tk.END, "=" * 100 + "\n\n")
        self.results_text.insert(tk.END, f"Asset Class: {self.asset_class.get()} | Product: {self.product_type.get()}\n")
        self.results_text.insert(tk.END, f"Total Trades Processed: {len(self.results)}\n\n")
        
        # Summary statistics
        matched_count = sum(1 for r in self.results if r["MatchedUPI"] is not None)
        avg_score = sum(r["Score"] for r in self.results) / len(self.results) if self.results else 0
        
        self.results_text.insert(tk.END, f"Matches Found: {matched_count}/{len(self.results)}\n")
        self.results_text.insert(tk.END, f"Average Match Score: {avg_score:.1f}\n")
        self.results_text.insert(tk.END, "=" * 100 + "\n\n")
        
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
            
            # Display key trade details (only non-empty values)
            self.results_text.insert(tk.END, "Key Trade Details:\n")
            for key, value in trade_details.items():
                if value and pd.notna(value) and str(value).strip():
                    self.results_text.insert(tk.END, f"  - {key}: {value}\n")
            
            # UPI match result
            self.results_text.insert(tk.END, f"Match Score: {score}%\n")
            self.results_text.insert(tk.END, f"Status: {message}\n")
            
            if matched_upi:
                identifier = matched_upi.get("Identifier", {})
                attributes = matched_upi.get("Attributes", {})
                derived = matched_upi.get("Derived", {})
                
                self.results_text.insert(tk.END, f"Matched UPI Code: {identifier.get('UPI', 'N/A')}\n")
                self.results_text.insert(tk.END, "UPI Details:\n")
                
                # Display key UPI attributes
                for key, value in attributes.items():
                    if value:
                        self.results_text.insert(tk.END, f"  - {key}: {value}\n")
                
                # Display some derived attributes
                if derived.get("ShortName"):
                    self.results_text.insert(tk.END, f"  - Short Name: {derived.get('ShortName')}\n")
                if derived.get("UnderlierName"):
                    self.results_text.insert(tk.END, f"  - Underlier Name: {derived.get('UnderlierName')}\n")
            
            # Separator between trades
            self.results_text.insert(tk.END, "\n" + "-" * 100 + "\n\n")
    
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
            
            for i, result in enumerate(self.results):
                trade_details = result["TradeDetails"]
                matched_upi = result["MatchedUPI"]
                
                row = {}
                
                # Add trade details
                for key, value in trade_details.items():
                    row[f"Trade_{key}"] = value
                
                # Add UPI match details
                row["Match_Score"] = result["Score"]
                row["Match_Message"] = result["Message"]
                row["Asset_Class"] = self.asset_class.get()
                row["Product_Type"] = self.product_type.get()
                
                if matched_upi:
                    identifier = matched_upi.get("Identifier", {})
                    attributes = matched_upi.get("Attributes", {})
                    derived = matched_upi.get("Derived", {})
                    
                    row["UPI_Code"] = identifier.get("UPI", "")
                    row["UPI_Status"] = identifier.get("Status", "")
                    row["UPI_LastUpdate"] = identifier.get("LastUpdateDateTime", "")
                    
                    # Add all attributes
                    for key, value in attributes.items():
                        row[f"UPI_{key}"] = value
                    
                    # Add key derived fields
                    row["UPI_ShortName"] = derived.get("ShortName", "")
                    row["UPI_UnderlierName"] = derived.get("UnderlierName", "")
                    row["UPI_ClassificationType"] = derived.get("ClassificationType", "")
                
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