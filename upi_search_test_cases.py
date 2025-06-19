import unittest
import pandas as pd
import json
import tempfile
import os
from upi_search_batch import UPISearchBatch

class TestCNHHandling(unittest.TestCase):
    def setUp(self):
        """Set up test data"""
        # Create sample UPI data with CNH-specific UPIs
        self.sample_upi_data = {
            "upis": [
                {
                    "upiCode": "CNH_SWAP_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Swap",
                    "product": "Non_Deliverable_FX_Swap",
                    "underlying": {
                        "currencyPair": "USD/CNY",
                        "settlementCurrency": "CNY"
                    },
                    "deliveryType": "Cash",
                    "placeOfSettlement": "Hong Kong"
                },
                {
                    "upiCode": "CNH_FWD_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Forward",
                    "product": "Non_Standard",
                    "underlying": {
                        "currencyPair": "USD/CNY",
                        "settlementCurrency": "CNY"
                    },
                    "deliveryType": "Cash",
                    "placeOfSettlement": "Hong Kong"
                },
                {
                    "upiCode": "CNH_OPT_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Option",
                    "product": "Non_Standard",
                    "underlying": {
                        "currencyPair": "USD/CNY",
                        "settlementCurrency": "CNY"
                    },
                    "optionType": "Call",
                    "optionStyle": "European",
                    "deliveryType": "Cash",
                    "placeOfSettlement": "Hong Kong"
                },
                {
                    "upiCode": "REGULAR_FWD_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Forward",
                    "product": "Vanilla",
                    "underlying": {
                        "currencyPair": "EUR/USD"
                    },
                    "deliveryType": "Physical"
                }
            ]
        }
        
        # Create sample trade data with CNH trades
        self.sample_trade_data = pd.DataFrame([
            {
                'TradeID': 'T001',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Swap',
                'CcyPair': 'USD/CNH',
                'SettlementCcy': 'CNH',
                'DeliveryType': 'Cash'
            },
            {
                'TradeID': 'T002',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Forward',
                'CcyPair': 'USD/CNH',
                'SettlementCcy': 'CNH',
                'DeliveryType': 'Cash'
            },
            {
                'TradeID': 'T003',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Option',
                'CcyPair': 'USD/CNH',
                'SettlementCcy': 'CNH',
                'OptionType': 'Call',
                'OptionStyle': 'European',
                'DeliveryType': 'Cash'
            },
            {
                'TradeID': 'T004',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Forward',
                'CcyPair': 'EUR/USD',
                'DeliveryType': 'Physical'
            }
        ])
        
        # Create temporary files
        self.upi_file = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
        json.dump(self.sample_upi_data, self.upi_file)
        self.upi_file.close()
        
        self.trade_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.xlsx', delete=False)
        self.sample_trade_data.to_excel(self.trade_file.name, index=False)
        self.trade_file.close()
        
        # Initialize processor
        self.processor = UPISearchBatch()
        self.processor.load_upi_data(self.upi_file.name)
        self.processor.load_trade_data(self.trade_file.name)
    
    def tearDown(self):
        """Clean up temporary files"""
        os.unlink(self.upi_file.name)
        os.unlink(self.trade_file.name)
    
    def test_cnh_detection_and_normalization(self):
        """Test that CNH trades are detected and currency is normalized to CNY"""
        self.processor.apply_cnh_handling()
        
        # Check that CNH trades have been processed
        cnh_trades = self.processor.trade_data[
            self.processor.trade_data['ProcessedCurrency'] == 'CNY'
        ]
        
        # Should have 3 CNH trades (T001, T002, T003)
        self.assertEqual(len(cnh_trades), 3)
        
        # Check that all CNH trades have Hong Kong as place of settlement
        for idx, trade in cnh_trades.iterrows():
            self.assertEqual(trade['ProcessedPlaceofSettlement'], 'Hong Kong')
    
    def test_cnh_usecase_assignment(self):
        """Test that CNH trades get correct UseCase assignments"""
        self.processor.apply_cnh_handling()
        
        # Check Swap trade (T001)
        swap_trade = self.processor.trade_data[
            self.processor.trade_data['TradeID'] == 'T001'
        ].iloc[0]
        self.assertEqual(swap_trade['ProcessedUseCase'], 'Non_Deliverable_FX_Swap')
        
        # Check Forward trade (T002)
        forward_trade = self.processor.trade_data[
            self.processor.trade_data['TradeID'] == 'T002'
        ].iloc[0]
        self.assertEqual(forward_trade['ProcessedUseCase'], 'Non_Standard')
        
        # Check Option trade (T003)
        option_trade = self.processor.trade_data[
            self.processor.trade_data['TradeID'] == 'T003'
        ].iloc[0]
        self.assertEqual(option_trade['ProcessedUseCase'], 'Non_Standard')
        
        # Check regular trade (T004) - should not have processed UseCase
        regular_trade = self.processor.trade_data[
            self.processor.trade_data['TradeID'] == 'T004'
        ].iloc[0]
        self.assertEqual(regular_trade['ProcessedUseCase'], '')
    
    def test_cnh_upi_matching(self):
        """Test that CNH trades match with appropriate UPIs"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        results = self.processor.search_upis('FX')
        
        # Check that CNH trades match with CNH-specific UPIs
        swap_result = next(r for r in results if r['Original_TradeID'] == 'T001')
        self.assertEqual(swap_result['Best_UPI'], 'CNH_SWAP_001')
        self.assertGreater(swap_result['Match_Score'], 80)  # Should be high confidence
        
        forward_result = next(r for r in results if r['Original_TradeID'] == 'T002')
        self.assertEqual(forward_result['Best_UPI'], 'CNH_FWD_001')
        self.assertGreater(forward_result['Match_Score'], 80)
        
        option_result = next(r for r in results if r['Original_TradeID'] == 'T003')
        self.assertEqual(option_result['Best_UPI'], 'CNH_OPT_001')
        self.assertGreater(option_result['Match_Score'], 80)
        
        # Check that regular trade matches with regular UPI
        regular_result = next(r for r in results if r['Original_TradeID'] == 'T004')
        self.assertEqual(regular_result['Best_UPI'], 'REGULAR_FWD_001')
    
    def test_column_mapping_with_processed_fields(self):
        """Test that column mapping includes processed fields for CNH handling"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        
        # Check that processed fields are mapped
        self.assertIn('Place of Settlement', self.processor.column_mappings)
        self.assertEqual(
            self.processor.column_mappings['Place of Settlement'],
            'ProcessedPlaceofSettlement'
        )
        
        # Check that settlement currency mapping includes processed currency
        if 'Settlement Currency' in self.processor.column_mappings:
            self.assertIn('ProcessedCurrency', self.processor.column_mappings['Settlement Currency'])
    
    def test_extract_trade_attributes_with_cnh_overrides(self):
        """Test that trade attribute extraction applies CNH overrides"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        
        # Get a CNH trade
        cnh_trade = self.processor.trade_data[
            self.processor.trade_data['TradeID'] == 'T001'
        ].iloc[0]
        
        attrs = self.processor.extract_trade_attributes(cnh_trade)
        
        # Check that CNH-specific attributes are applied
        self.assertEqual(attrs.get('Product Type'), 'Non_Deliverable_FX_Swap')
        self.assertEqual(attrs.get('Place of Settlement'), 'Hong Kong')
        
        # Check that currency is normalized
        if 'Settlement Currency' in attrs:
            self.assertEqual(attrs['Settlement Currency'], 'CNY')

class TestCNHIntegration(unittest.TestCase):
    """Integration tests for CNH handling"""
    
    def test_end_to_end_cnh_processing(self):
        """Test complete CNH processing workflow"""
        # Create test data
        upi_data = {
            "upis": [
                {
                    "upiCode": "CNH_TEST_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Swap",
                    "product": "Non_Deliverable_FX_Swap",
                    "underlying": {
                        "currencyPair": "USD/CNY",
                        "settlementCurrency": "CNY"
                    },
                    "deliveryType": "Cash",
                    "placeOfSettlement": "Hong Kong"
                }
            ]
        }
        
        trade_data = pd.DataFrame([
            {
                'TradeID': 'CNH_TEST',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Swap',
                'CcyPair': 'USD/CNH',
                'SettlementCcy': 'CNH',
                'DeliveryType': 'Cash'
            }
        ])
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as upi_file:
            json.dump(upi_data, upi_file)
            upi_file_path = upi_file.name
        
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.xlsx', delete=False) as trade_file:
            trade_data.to_excel(trade_file.name, index=False)
            trade_file_path = trade_file.name
        
        try:
            # Process with batch processor
            processor = UPISearchBatch()
            processor.load_upi_data(upi_file_path)
            processor.load_trade_data(trade_file_path)
            processor.apply_cnh_handling()
            processor.auto_map_columns('FX')
            results = processor.search_upis('FX')
            
            # Verify results
            self.assertEqual(len(results), 1)
            result = results[0]
            self.assertEqual(result['Best_UPI'], 'CNH_TEST_001')
            self.assertGreater(result['Match_Score'], 80)
            
            # Verify trade attributes include CNH-specific values
            trade_attrs = result['Trade_Attributes']
            self.assertEqual(trade_attrs.get('Product Type'), 'Non_Deliverable_FX_Swap')
            self.assertEqual(trade_attrs.get('Place of Settlement'), 'Hong Kong')
            
        finally:
            # Clean up
            os.unlink(upi_file_path)
            os.unlink(trade_file_path)

def run_cnh_tests():
    """Run CNH-specific tests"""
    print("Running CNH Handling Tests...")
    print("=" * 50)
    
    # Create test suite
    suite = unittest.TestSuite()
    
    # Add CNH handling tests
    suite.addTest(unittest.makeSuite(TestCNHHandling))
    suite.addTest(unittest.makeSuite(TestCNHIntegration))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print("\n" + "=" * 50)
    print("CNH Handling Test Summary:")
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    
    if result.failures:
        print("\nFailures:")
        for test, traceback in result.failures:
            print(f"- {test}: {traceback}")
    
    if result.errors:
        print("\nErrors:")
        for test, traceback in result.errors:
            print(f"- {test}: {traceback}")
    
    if result.wasSuccessful():
        print("\nAll CNH handling tests passed! ✅")
    else:
        print("\nSome tests failed. Please review the implementation. ❌")
    
    return result.wasSuccessful()

if __name__ == "__main__":
    # Run CNH-specific tests
    success = run_cnh_tests()
    
    if not success:
        exit(1)
    
    print("\nCNH handling implementation is working correctly!")
    print("\nKey features implemented:")
    print("✅ CNH currency detection and normalization to CNY")
    print("✅ Automatic UseCase assignment based on instrument type:")
    print("   - Swaps: Non_Deliverable_FX_Swap")
    print("   - Forwards/Options: Non_Standard")
    print("✅ Place of Settlement set to 'Hong Kong' for CNH trades")
    print("✅ Enhanced UPI matching with CNH-specific attributes")
    print("✅ Integration with both GUI and batch processing versions")