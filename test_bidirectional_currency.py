import unittest
import pandas as pd
import json
import tempfile
import os
from upi_search_batch import UPISearchBatch

class TestBidirectionalCurrencyMatching(unittest.TestCase):
    def setUp(self):
        """Set up test data for bidirectional currency matching"""
        # Create sample UPI data with different currency pair orders
        self.sample_upi_data = {
            "upis": [
                {
                    "upiCode": "USD_EUR_FWD_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Forward",
                    "product": "Vanilla",
                    "underlying": {
                        "currencyPair": "USD/EUR"
                    },
                    "deliveryType": "Physical"
                },
                {
                    "upiCode": "EUR_USD_FWD_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Forward",
                    "product": "Vanilla",
                    "underlying": {
                        "currencyPair": "EUR/USD"
                    },
                    "deliveryType": "Physical"
                },
                {
                    "upiCode": "GBP_JPY_FWD_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Forward",
                    "product": "Vanilla",
                    "underlying": {
                        "currencyPair": "GBP/JPY"
                    },
                    "deliveryType": "Physical"
                },
                {
                    "upiCode": "CNY_USD_NDF_001",
                    "assetClass": "ForeignExchange",
                    "instrumentType": "Forward",
                    "product": "Non_Standard",
                    "underlying": {
                        "currencyPair": "CNY/USD",
                        "settlementCurrency": "CNY"
                    },
                    "deliveryType": "Cash",
                    "placeOfSettlement": "Hong Kong"
                }
            ]
        }
        
        # Create sample trade data with different currency pair orders
        self.sample_trade_data = pd.DataFrame([
            {
                'TradeID': 'T001',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Forward',
                'CcyPair': 'EUR/USD',  # Reversed order from first UPI
                'DeliveryType': 'Physical'
            },
            {
                'TradeID': 'T002',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Forward',
                'CcyPair': 'USD/EUR',  # Same order as first UPI
                'DeliveryType': 'Physical'
            },
            {
                'TradeID': 'T003',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Forward',
                'CcyPair': 'JPY/GBP',  # Reversed order from third UPI
                'DeliveryType': 'Physical'
            },
            {
                'TradeID': 'T004',
                'AssetClass': 'ForeignExchange',
                'InstrumentType': 'Forward',
                'CcyPair': 'USD/CNH',  # CNH should be normalized to CNY and matched
                'DeliveryType': 'Cash'
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
    
    def test_currency_extraction_from_pair(self):
        """Test that individual currencies are correctly extracted from currency pairs"""
        # Test normal currency pair
        ccy1 = self.processor.extract_currency_from_pair("USD/EUR", 0)
        ccy2 = self.processor.extract_currency_from_pair("USD/EUR", 1)
        self.assertEqual(ccy1, "USD")
        self.assertEqual(ccy2, "EUR")
        
        # Test with spaces
        ccy1 = self.processor.extract_currency_from_pair("GBP / JPY", 0)
        ccy2 = self.processor.extract_currency_from_pair("GBP / JPY", 1)
        self.assertEqual(ccy1, "GBP")
        self.assertEqual(ccy2, "JPY")
        
        # Test invalid format
        ccy1 = self.processor.extract_currency_from_pair("INVALID", 0)
        self.assertEqual(ccy1, "")
    
    def test_trade_attribute_extraction_with_currencies(self):
        """Test that trade attributes include individual currencies"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        
        # Get a trade with currency pair
        trade = self.processor.trade_data.iloc[0]
        attrs = self.processor.extract_trade_attributes(trade)
        
        # Check that individual currencies are extracted
        self.assertIn('TradeNotionalCurrency', attrs)
        self.assertIn('TradeOtherNotionalCurrency', attrs)
        self.assertEqual(attrs['TradeNotionalCurrency'], 'EUR')
        self.assertEqual(attrs['TradeOtherNotionalCurrency'], 'USD')
    
    def test_bidirectional_currency_matching(self):
        """Test that currencies match in both directions"""
        # Create test trade attributes
        trade_attrs = {
            'TradeNotionalCurrency': 'EUR',
            'TradeOtherNotionalCurrency': 'USD'
        }
        
        # Test UPI with same order
        upi_same_order = {
            'underlying': {'currencyPair': 'EUR/USD'}
        }
        
        # Test UPI with reversed order
        upi_reversed_order = {
            'underlying': {'currencyPair': 'USD/EUR'}
        }
        
        # Test UPI with different currencies
        upi_different = {
            'underlying': {'currencyPair': 'GBP/JPY'}
        }
        
        # Both same and reversed order should match
        self.assertTrue(self.processor.match_currencies_bidirectional(trade_attrs, upi_same_order))
        self.assertTrue(self.processor.match_currencies_bidirectional(trade_attrs, upi_reversed_order))
        
        # Different currencies should not match
        self.assertFalse(self.processor.match_currencies_bidirectional(trade_attrs, upi_different))
    
    def test_cnh_currency_normalization_in_bidirectional_matching(self):
        """Test that CNH is properly normalized to CNY in bidirectional matching"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        
        # Get CNH trade
        cnh_trade = self.processor.trade_data[
            self.processor.trade_data['TradeID'] == 'T004'
        ].iloc[0]
        
        attrs = self.processor.extract_trade_attributes(cnh_trade)
        
        # Check that CNH was normalized to CNY
        self.assertEqual(attrs.get('TradeNotionalCurrency'), 'USD')
        self.assertEqual(attrs.get('TradeOtherNotionalCurrency'), 'CNY')  # Should be normalized from CNH
    
    def test_end_to_end_bidirectional_matching(self):
        """Test complete bidirectional matching workflow"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        results = self.processor.search_upis('FX')
        
        # Check that trades match with UPIs regardless of currency order
        
        # T001: EUR/USD trade should match with USD/EUR UPI
        t001_result = next(r for r in results if r['Original_TradeID'] == 'T001')
        self.assertEqual(t001_result['Best_UPI'], 'USD_EUR_FWD_001')
        self.assertGreater(t001_result['Match_Score'], 80)
        
        # T002: USD/EUR trade should match with USD/EUR UPI
        t002_result = next(r for r in results if r['Original_TradeID'] == 'T002')
        self.assertEqual(t002_result['Best_UPI'], 'USD_EUR_FWD_001')
        self.assertGreater(t002_result['Match_Score'], 80)
        
        # T003: JPY/GBP trade should match with GBP/JPY UPI
        t003_result = next(r for r in results if r['Original_TradeID'] == 'T003')
        self.assertEqual(t003_result['Best_UPI'], 'GBP_JPY_FWD_001')
        self.assertGreater(t003_result['Match_Score'], 80)
        
        # T004: USD/CNH trade should match with CNY/USD UPI (after CNH normalization)
        t004_result = next(r for r in results if r['Original_TradeID'] == 'T004')
        self.assertEqual(t004_result['Best_UPI'], 'CNY_USD_NDF_001')
        self.assertGreater(t004_result['Match_Score'], 80)
    
    def test_scoring_weights_adjustment(self):
        """Test that scoring weights are properly adjusted for bidirectional matching"""
        self.processor.apply_cnh_handling()
        self.processor.auto_map_columns('FX')
        
        # Get a trade
        trade = self.processor.trade_data.iloc[0]
        trade_attrs = self.processor.extract_trade_attributes(trade)
        
        # Get a matching UPI
        upi = self.sample_upi_data['upis'][0]
        
        # Calculate score
        score = self.processor.calculate_match_score(trade_attrs, upi, 'FX')
        
        # Score should include currency matching points
        # Asset Class (20) + Instrument Type (20) + Product Type (20) + Currencies (20) + Delivery Type (10) = 90
        self.assertGreaterEqual(score, 80)  # Allow for some variation in matching

def run_bidirectional_currency_tests():
    """Run bidirectional currency matching tests"""
    print("Running Bidirectional Currency Matching Tests...")
    print("=" * 60)
    
    # Create test suite
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestBidirectionalCurrencyMatching))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print("\n" + "=" * 60)
    print("Bidirectional Currency Matching Test Summary:")
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
        print("\nAll bidirectional currency matching tests passed! ✅")
        print("\nKey features implemented:")
        print("✅ Currency pair extraction from trade data")
        print("✅ Bidirectional currency matching (USD/EUR matches EUR/USD)")
        print("✅ CNH normalization to CNY in currency matching")
        print("✅ Updated scoring system for individual currency matching")
        print("✅ Integration with existing CNH handling logic")
    else:
        print("\nSome tests failed. Please review the implementation. ❌")
    
    return result.wasSuccessful()

if __name__ == "__main__":
    success = run_bidirectional_currency_tests()
    
    if not success:
        exit(1)
    
    print("\nBidirectional currency matching implementation is working correctly!")