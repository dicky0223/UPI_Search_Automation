import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import json
import os
import re
from tkinter import scrolledtext
import traceback
import time

class UPISearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("UPI Search Automation Tool - DSB RECORDS Format")
        self.root.geometry("1400x1000")  # Increased size for new columns
        
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
        
        # Load UPI schemas for attribute definitions
        self.upi_schemas = {}
        self._load_upi_schemas()
        
        # Create UI
        self.create_ui()
    
    def _load_upi_schemas(self):
        """Load UPI schema files to get attribute definitions and allowable values"""
        try:
            # Define schema file mapping based on asset class and product
            schema_files = {
                ("Foreign_Exchange", "Forward"): "Foreign_Exchange.Forward.Forward.UPI.V1.json",
                ("Foreign_Exchange", "NDF"): "Foreign_Exchange.Forward.NDF.UPI.V1.json",
                ("Foreign_Exchange", "Non_Standard"): "Foreign_Exchange.Forward.Non_Standard.UPI.V1.json",
                ("Foreign_Exchange", "Digital_Option"): "Foreign_Exchange.Option.Digital_Option.UPI.V1.json",
                ("Foreign_Exchange", "Vanilla_Option"): "Foreign_Exchange.Option.Vanilla_Option.UPI.V1.json",
                ("Foreign_Exchange", "FX_Swap"): "Foreign_Exchange.Swap.FX_Swap.UPI.V1.json",
                ("Rates", "Basis"): "Rates.Swap.Basis.UPI.V1.json",
                ("Rates", "Basis_OIS"): "Rates.Swap.Basis_OIS.UPI.V1.json",
                ("Rates", "Cross_Currency_Basis"): "Rates.Swap.Cross_Currency_Basis.UPI.V1.json",
                ("Rates", "Cross_Currency_Fixed_Fixed"): "Rates.Swap.Cross_Currency_Fixed_Fixed.UPI.V1.json",
                ("Rates", "Cross_Currency_Fixed_Float"): "Rates.Swap.Cross_Currency_Fixed_Float.UPI.V1.json",
            }
            
            # Load each schema file
            for key, filename in schema_files.items():
                if os.path.exists(filename):
                    try:
                        with open(filename, 'r', encoding='utf-8') as f:
                            schema = json.load(f)
                            self.upi_schemas[key] = schema
                    except Exception as e:
                        print(f"Error loading schema {filename}: {e}")
                        
        except Exception as e:
            print(f"Error loading UPI schemas: {e}")
    
    def get_upi_attribute_details(self, field_name):
        """Get attribute details (description, enum values) from UPI schema"""
        try:
            asset_class = "Foreign_Exchange" if self.asset_class.get() == "FX" else "Rates"
            product = self.product_type.get()
            
            schema_key = (asset_class, product)
            schema = self.upi_schemas.get(schema_key)
            
            if not schema:
                return {"description": "", "enum": [], "elaboration": {}}
            
            # Look for the field in Attributes section
            attributes = schema.get("properties", {}).get("Attributes", {}).get("properties", {})
            field_def = attributes.get(field_name, {})
            
            # Extract details
            description = field_def.get("description", "")
            enum_values = field_def.get("enum", [])
            elaboration = field_def.get("elaboration", {})
            
            # If not found in Attributes, check Derived section
            if not description and not enum_values:
                derived = schema.get("properties", {}).get("Derived", {}).get("properties", {})
                field_def = derived.get(field_name, {})
                description = field_def.get("description", "")
                enum_values = field_def.get("enum", [])
                elaboration = field_def.get("elaboration", {})
            
            return {
                "description": description,
                "enum": enum_values,
                "elaboration": elaboration
            }
            
        except Exception as e:
            print(f"Error getting attribute details for {field_name}: {e}")
            return {"description": "", "enum": [], "elaboration": {}}
    
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
        upi_frame = ttk.LabelFrame(self.tab1, text="UPI Data (RECORDS File)")
        upi_frame.pack(fill='x', expand=True, padx=10, pady=10)
        
        ttk.Label(upi_frame, text="UPI RECORDS File:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(upi_frame, textvariable=self.upi_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(upi_frame, text="Browse", command=self.browse_upi_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Add note about RECORDS format
        ttk.Label(upi_frame, text="Note: Supports RECORDS files downloaded from DSB website (JSON line format)", 
                 font=("Arial", 8), foreground="gray").grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky='w')
        
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
        
        # Progress bar for UPI search
        self.progress_frame = ttk.Frame(self.tab3)
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_label = ttk.Label(self.progress_frame, text="")
        
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
        filename = filedialog.askopenfilename(
            filetypes=[("RECORDS files", "*.RECORDS"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.upi_file_path.set(filename)
    
    def browse_trade_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename:
            self.trade_file_path.set(filename)
    
    def parse_records_file(self, file_path):
        """Parse RECORDS file format from DSB - JSON line format"""
        try:
            upi_records = []
            
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # Process each line as a JSON object
            for line_num, line in enumerate(lines, 1):
                line = line.strip()
                
                # Skip empty lines and comments
                if not line or line.startswith('#'):
                    continue
                
                try:
                    # Parse each line as JSON
                    record = json.loads(line)
                    
                    # Validate that it's a UPI record with required structure
                    if not self.is_valid_upi_record(record):
                        continue
                    
                    # Filter by asset class
                    asset_class_filter = "Foreign_Exchange" if self.asset_class.get() == "FX" else "Rates"
                    header = record.get("Header", {})
                    
                    if header.get("AssetClass") == asset_class_filter:
                        upi_records.append(record)
                        
                except json.JSONDecodeError as e:
                    print(f"Error parsing JSON on line {line_num}: {e}")
                    continue
                except Exception as e:
                    print(f"Error processing line {line_num}: {e}")
                    continue
            
            if not upi_records:
                raise ValueError(f"No valid UPI records found for asset class: {self.asset_class.get()}")
            
            return upi_records
            
        except Exception as e:
            raise Exception(f"Error parsing RECORDS file: {str(e)}")
    
    def is_valid_upi_record(self, record):
        """Validate that the record has the expected UPI structure"""
        try:
            # Check for required top-level keys
            required_keys = ["Header", "Identifier", "Derived", "Attributes"]
            if not all(key in record for key in required_keys):
                return False
            
            # Check Header structure
            header = record.get("Header", {})
            if not all(key in header for key in ["AssetClass", "InstrumentType", "UseCase", "Level"]):
                return False
            
            # Check Identifier structure
            identifier = record.get("Identifier", {})
            if not identifier.get("UPI"):
                return False
            
            return True
            
        except Exception:
            return False
    
    def load_data(self):
        try:
            # Check if files are selected
            if not self.upi_file_path.get() or not self.trade_file_path.get():
                messagebox.showerror("Error", "Please select both UPI RECORDS file and Trade Excel file")
                return
            
            # Show loading status
            self.status_upload.set("Loading UPI data...")
            self.root.update_idletasks()
            
            # Load UPI data from RECORDS file
            self.upi_data = self.parse_records_file(self.upi_file_path.get())
            
            # Update status
            self.status_upload.set("Loading trade data...")
            self.root.update_idletasks()
            
            # Load trade data
            self.trade_data = pd.read_excel(self.trade_file_path.get())
            
            # Extract available products based on asset class
            self.extract_available_products()
            
            # Update status
            upi_count = len(self.upi_data)
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
            # Extract products based on asset class
            asset_class_filter = "Foreign_Exchange" if self.asset_class.get() == "FX" else "Rates"
            
            products = set()
            for upi in self.upi_data:
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
                 font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=5, pady=10)
        ttk.Label(scrollable_frame, text=f"Asset Class: {self.asset_class.get()} | Product: {self.product_type.get()}", 
                 font=("Arial", 10)).grid(row=1, column=0, columnspan=5, pady=5)
        
        # Column headers
        ttk.Label(scrollable_frame, text="Attribute", font=("Arial", 9, "bold")).grid(row=2, column=0, padx=5, pady=5, sticky='w')
        ttk.Label(scrollable_frame, text="Input Method", font=("Arial", 9, "bold")).grid(row=2, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(scrollable_frame, text="Value/Column", font=("Arial", 9, "bold")).grid(row=2, column=2, padx=5, pady=5, sticky='w')
        ttk.Label(scrollable_frame, text="Allowable Values", font=("Arial", 9, "bold")).grid(row=2, column=3, padx=5, pady=5, sticky='w')
        ttk.Label(scrollable_frame, text="Description", font=("Arial", 9, "bold")).grid(row=2, column=4, padx=5, pady=5, sticky='w')
        
        # Dictionary to store the mapping variables
        self.mapping_vars = {}
        
        row = 3
        for label, field_name, required in mapping_fields:
            # Label with required indicator
            label_text = label + (" *" if required else "")
            ttk.Label(scrollable_frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky='nw')
            
            # Create variables for this field
            input_method_var = tk.StringVar(value="column")  # "column" or "manual"
            value_var = tk.StringVar()
            
            self.mapping_vars[field_name] = {
                "method": input_method_var,
                "value": value_var
            }
            
            # Input method selection (Column or Manual)
            method_frame = ttk.Frame(scrollable_frame)
            method_frame.grid(row=row, column=1, padx=5, pady=5, sticky='nw')
            
            ttk.Radiobutton(method_frame, text="Column", variable=input_method_var, value="column",
                           command=lambda fn=field_name: self.update_input_method(fn)).pack(anchor='w')
            ttk.Radiobutton(method_frame, text="Manual", variable=input_method_var, value="manual",
                           command=lambda fn=field_name: self.update_input_method(fn)).pack(anchor='w')
            
            # Value/Column selection frame
            value_frame = ttk.Frame(scrollable_frame)
            value_frame.grid(row=row, column=2, padx=5, pady=5, sticky='nw')
            
            # Column dropdown (initially visible)
            columns = list(self.trade_data.columns) + ["N/A"]
            auto_select = self.find_matching_column(label, columns)
            if auto_select:
                value_var.set(auto_select)
            else:
                value_var.set("N/A" if not required else columns[0])
            
            column_dropdown = ttk.Combobox(value_frame, textvariable=value_var, values=columns, width=25)
            column_dropdown.pack()
            
            # Manual input entry (initially hidden)
            manual_entry = ttk.Entry(value_frame, textvariable=value_var, width=25)
            
            # Store widgets for later access
            self.mapping_vars[field_name]["column_widget"] = column_dropdown
            self.mapping_vars[field_name]["manual_widget"] = manual_entry
            
            # Get attribute details from schema
            attr_details = self.get_upi_attribute_details(field_name)
            
            # Allowable values display
            allowable_frame = ttk.Frame(scrollable_frame)
            allowable_frame.grid(row=row, column=3, padx=5, pady=5, sticky='nw')
            
            if attr_details["enum"]:
                # Create scrollable text for enum values
                enum_text = scrolledtext.ScrolledText(allowable_frame, width=20, height=4, wrap=tk.WORD)
                enum_text.pack()
                
                # Add enum values with elaboration
                for enum_val in attr_details["enum"]:
                    enum_text.insert(tk.END, f"• {enum_val}")
                    if enum_val in attr_details["elaboration"]:
                        elaboration = attr_details["elaboration"][enum_val]
                        if elaboration and elaboration != enum_val:
                            enum_text.insert(tk.END, f": {elaboration[:50]}...")
                    enum_text.insert(tk.END, "\n")
                
                enum_text.config(state='disabled')
            else:
                ttk.Label(allowable_frame, text="Any value", font=("Arial", 8), foreground="gray").pack()
            
            # Description display
            desc_frame = ttk.Frame(scrollable_frame)
            desc_frame.grid(row=row, column=4, padx=5, pady=5, sticky='nw')
            
            description = attr_details["description"] or self.get_field_description(field_name)
            if description:
                desc_text = tk.Text(desc_frame, width=30, height=3, wrap=tk.WORD, font=("Arial", 8))
                desc_text.insert(tk.END, description)
                desc_text.config(state='disabled')
                desc_text.pack()
            else:
                ttk.Label(desc_frame, text="No description available", font=("Arial", 8), foreground="gray").pack()
            
            row += 1
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Show map button
        self.map_button.pack(pady=10)
    
    def update_input_method(self, field_name):
        """Update the input widget based on selected method"""
        mapping_info = self.mapping_vars[field_name]
        method = mapping_info["method"].get()
        
        # Hide both widgets first
        mapping_info["column_widget"].pack_forget()
        mapping_info["manual_widget"].pack_forget()
        
        # Show the appropriate widget
        if method == "column":
            mapping_info["column_widget"].pack()
            # Reset to column selection if switching from manual
            columns = list(self.trade_data.columns) + ["N/A"]
            if mapping_info["value"].get() not in columns:
                mapping_info["value"].set("N/A")
        else:  # manual
            mapping_info["manual_widget"].pack()
            # Clear value when switching to manual
            mapping_info["value"].set("")
    
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
            
            # Show progress bar
            self.progress_frame.pack(pady=10)
            self.progress_label.pack(pady=5)
            self.progress_bar.pack(fill='x', padx=20, pady=5)
            
            # Get mapping from UI
            mapping = {}
            for field_name, mapping_info in self.mapping_vars.items():
                method = mapping_info["method"].get()
                value = mapping_info["value"].get()
                
                mapping[field_name] = {
                    "method": method,
                    "value": value
                }
            
            # Initialize progress
            total_trades = len(self.trade_data)
            self.progress_bar['maximum'] = total_trades
            self.progress_bar['value'] = 0
            
            # Process each trade with progress updates
            for index, trade in self.trade_data.iterrows():
                # Update progress
                current_trade = index + 1
                self.progress_bar['value'] = current_trade
                self.progress_label.config(text=f"Processing trade {current_trade} of {total_trades}...")
                self.status_mapping.set(f"Searching UPIs... {current_trade}/{total_trades}")
                
                # Force GUI update to show progress
                self.root.update_idletasks()
                
                # Find matching UPI for this trade
                result = self.find_matching_upi(trade, mapping)
                self.results.append(result)
                
                # Small delay to make progress visible (remove for production)
                time.sleep(0.01)
            
            # Hide progress bar
            self.progress_frame.pack_forget()
            
            # Display results
            self.display_results()
            
            # Show export button
            self.export_button.pack(pady=10)
            
            # Update status
            matched_count = sum(1 for r in self.results if r["MatchedUPI"] is not None)
            self.status_mapping.set(f"UPI search completed. {matched_count}/{len(self.results)} trades matched.")
            
            # Switch to results tab
            notebook = self.tab4.master
            notebook.select(3)  # Select the fourth tab (index 3)
            
        except Exception as e:
            # Hide progress bar on error
            self.progress_frame.pack_forget()
            messagebox.showerror("Error", f"Error searching UPIs: {str(e)}\n{traceback.format_exc()}")
            self.status_mapping.set(f"Error: {str(e)}")
    
    def find_matching_upi(self, trade, mapping):
        result = {"TradeDetails": trade.to_dict(), "MatchedUPI": None, "Score": 0, "Message": "", "AllMatches": []}
        
        try:
            # Filter UPIs by asset class and product
            asset_class_filter = "Foreign_Exchange" if self.asset_class.get() == "FX" else "Rates"
            product_filter = self.product_type.get()
            
            relevant_upis = []
            for upi in self.upi_data:
                header = upi.get("Header", {})
                if (header.get("AssetClass") == asset_class_filter and 
                    header.get("UseCase") == product_filter):
                    relevant_upis.append(upi)
            
            if not relevant_upis:
                result["Message"] = f"No UPI records found for {asset_class_filter} {product_filter}"
                return result
            
            # Perform matching and collect all scores
            all_matches = []
            for upi in relevant_upis:
                score = self.calculate_upi_score(trade, mapping, upi)
                if score > 0:  # Only include UPIs with some match
                    all_matches.append({
                        "upi": upi,
                        "score": score
                    })
            
            # Sort by score (highest first)
            all_matches.sort(key=lambda x: x["score"], reverse=True)
            result["AllMatches"] = all_matches
            
            # Set result based on best match
            threshold_score = 50  # Adjustable threshold
            if all_matches:
                best_match = all_matches[0]
                best_score = best_match["score"]
                
                if best_score >= threshold_score:
                    result["MatchedUPI"] = best_match["upi"]
                    result["Score"] = best_score
                    
                    # Check for multiple high-scoring matches
                    high_score_matches = [m for m in all_matches if m["score"] >= threshold_score]
                    if len(high_score_matches) > 1:
                        result["Message"] = f"Multiple UPIs found with high scores. Best match: {best_score}% (Total candidates: {len(high_score_matches)})"
                    else:
                        result["Message"] = f"UPI found with match score: {best_score}%"
                else:
                    result["Message"] = f"No matching UPI found with sufficient confidence (best score: {best_score}%, threshold: {threshold_score}%)"
            else:
                result["Message"] = "No UPI matches found based on provided trade attributes"
            
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
        for field_name, mapping_info in mapping.items():
            method = mapping_info["method"]
            value = mapping_info["value"]
            
            if method == "manual":
                # Use manual input value
                if not value or value.strip() == "":
                    continue
                trade_value = value.strip()
            else:  # column mapping
                column_name = value
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
        multiple_matches_count = sum(1 for r in self.results if len(r.get("AllMatches", [])) > 1)
        
        self.results_text.insert(tk.END, f"Matches Found: {matched_count}/{len(self.results)}\n")
        self.results_text.insert(tk.END, f"Average Match Score: {avg_score:.1f}%\n")
        self.results_text.insert(tk.END, f"Trades with Multiple Candidate UPIs: {multiple_matches_count}\n")
        self.results_text.insert(tk.END, "=" * 100 + "\n\n")
        
        # Display results for each trade
        for i, result in enumerate(self.results):
            trade_details = result["TradeDetails"]
            matched_upi = result["MatchedUPI"]
            score = result["Score"]
            message = result["Message"]
            all_matches = result.get("AllMatches", [])
            
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
            
            # Show multiple matches if available
            if len(all_matches) > 1:
                self.results_text.insert(tk.END, f"Alternative UPI Candidates ({len(all_matches)} total):\n")
                for j, match in enumerate(all_matches[:3]):  # Show top 3 matches
                    upi_code = match["upi"].get("Identifier", {}).get("UPI", "N/A")
                    match_score = match["score"]
                    self.results_text.insert(tk.END, f"  {j+1}. UPI: {upi_code} (Score: {match_score}%)\n")
                if len(all_matches) > 3:
                    self.results_text.insert(tk.END, f"  ... and {len(all_matches) - 3} more candidates\n")
            
            if matched_upi:
                identifier = matched_upi.get("Identifier", {})
                attributes = matched_upi.get("Attributes", {})
                derived = matched_upi.get("Derived", {})
                
                self.results_text.insert(tk.END, f"Selected UPI Code: {identifier.get('UPI', 'N/A')}\n")
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
                all_matches = result.get("AllMatches", [])
                
                row = {}
                
                # Add trade details
                for key, value in trade_details.items():
                    row[f"Trade_{key}"] = value
                
                # Add UPI match details
                row["Match_Score"] = result["Score"]
                row["Match_Message"] = result["Message"]
                row["Asset_Class"] = self.asset_class.get()
                row["Product_Type"] = self.product_type.get()
                row["Total_Candidate_UPIs"] = len(all_matches)
                
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
                
                # Add alternative UPI candidates
                for j, match in enumerate(all_matches[:5]):  # Export top 5 alternatives
                    alt_upi = match["upi"]
                    alt_identifier = alt_upi.get("Identifier", {})
                    row[f"Alternative_UPI_{j+1}_Code"] = alt_identifier.get("UPI", "")
                    row[f"Alternative_UPI_{j+1}_Score"] = match["score"]
                
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