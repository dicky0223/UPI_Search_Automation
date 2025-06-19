import argparse
import pandas as pd
import json
from datetime import datetime
import sys
import os

class UPISearchBatch:
    def __init__(self):
        self.upi_data = None
        self.trade_data = None
        self.results = None
        self.column_mappings = {}
    
    def load_upi_data(self, upi_file_path):
        """Load UPI data from JSON file"""
        try:
            with open(upi_file_path, 'r') as f:
                self.upi_data = json.load(f)
            print(f"Loaded {len(self.upi_data.get('upis', []))} UPI records")
            return True
        except Exception as e:
            print(f"Error loading UPI data: {str(e)}")
            return False
    
    def load_trade_data(self, trade_file_path):
        """Load trade data from Excel file"""
        try:
            self.trade_data = pd.read_excel(trade_file_path)
            print(f"Loaded {len(self.trade_data)} trade records")
            return True
        except Exception as e:
            print(f"Error loading trade data: {str(e)}")
            return False
    
    def apply_cnh_handling(self):
        """Apply CNH-specific handling logic to trade data"""
        print("Applying CNH handling logic...")
        
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
            print(f"Applied CNH handling to {cnh_trades_count} trades")
            print("CNH trades will use:")
            print("- Currency: CNY (normalized from CNH)")
            print("- PlaceofSettlement: Hong Kong")
            print("- UseCase: Non_Deliverable_FX_Swap (Swaps) or Non_Standard (Forwards/Options)")
        else:
            print("No CNH trades detected")
    
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
    
    def auto_map_columns(self, asset_class):
        """Automatically map columns based on common naming patterns"""
        trade_columns = list(self.trade_data.columns)
        
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
        
        print(f"Auto-mapped {len(self.column_mappings)} columns:")
        for attr, col in self.column_mappings.items():
            print(f"  {attr} -> {col}")
    
    def search_upis(self, asset_class):
        """Perform UPI search"""
        print("Starting UPI search...")
        results = []
        
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
        print(f"UPI search completed. Processed {len(results)} trades.")
        
        # Print summary
        matched_trades = sum(1 for r in results if r['Match_Score'] >= 50)
        high_confidence = sum(1 for r in results if r['Match_Score'] >= 80)
        
        print(f"Summary:")
        print(f"  Total trades: {len(results)}")
        print(f"  Matched trades (score ≥ 50): {matched_trades}")
        print(f"  High confidence matches (score ≥ 80): {high_confidence}")
        print(f"  Match rate: {(matched_trades/len(results))*100:.1f}%")
        
        return results
    
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
    
    def export_results(self, output_file):
        """Export results to Excel file"""
        if not self.results:
            print("No results to export.")
            return False
        
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
            df.to_excel(output_file, index=False)
            print(f"Results exported to {output_file}")
            return True
        
        except Exception as e:
            print(f"Error exporting results: {str(e)}")
            return False

def main():
    parser = argparse.ArgumentParser(description='UPI Search Automation Tool - Batch Processing')
    parser.add_argument('--upi', required=True, help='Path to UPI JSON file')
    parser.add_argument('--trade', required=True, help='Path to trade Excel file')
    parser.add_argument('--asset-class', choices=['FX', 'IR'], default='FX', help='Asset class (FX or IR)')
    parser.add_argument('--output', help='Output Excel file path')
    
    args = parser.parse_args()
    
    # Set default output filename if not provided
    if not args.output:
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        args.output = f'upi_search_results_{timestamp}.xlsx'
    
    # Initialize batch processor
    processor = UPISearchBatch()
    
    # Load data
    if not processor.load_upi_data(args.upi):
        sys.exit(1)
    
    if not processor.load_trade_data(args.trade):
        sys.exit(1)
    
    # Apply CNH handling
    processor.apply_cnh_handling()
    
    # Auto-map columns
    processor.auto_map_columns(args.asset_class)
    
    # Search UPIs
    processor.search_upis(args.asset_class)
    
    # Export results
    if processor.export_results(args.output):
        print(f"Process completed successfully. Results saved to {args.output}")
    else:
        print("Failed to export results.")
        sys.exit(1)

if __name__ == "__main__":
    main()