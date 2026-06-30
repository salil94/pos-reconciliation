#!/usr/bin/env python3
"""
Deliveroo POS-System Reconciliation Script with Duplicate Flagging
Similar to Careem script but shows duplicates instead of removing them
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
from pathlib import Path
import logging
from datetime import datetime
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def add_occurrence_column(df: pd.DataFrame, id_col: str = 'order_id',
                          sort_cols=('received_at', 'item_price', 'location', 'datetime', 'subtotal')) -> pd.DataFrame:
    """
    Sort by id + any of the provided stable columns that exist, then add a 0-based
    'occurrence' within each id. This makes each row uniquely matchable even if
    order_id is duplicated.
    """
    cols = [id_col] + [c for c in sort_cols if c in df.columns]
    if cols:
        df = df.sort_values(cols, kind='stable').copy()
    else:
        df = df.copy()
    df['occurrence'] = df.groupby(id_col).cumcount()
    return df

class DeliverooReconciliation:
    """Reconciliation tool for Deliveroo POS and System data with duplicate flagging"""

    # Restaurant configurations
    RESTAURANTS = {
        'brand_a': {
            'name': 'BrandA',
            'brand_filter': 'brand_a',
            'description': 'BrandA - Restaurant (Deliveroo)',
            'locations': ['Location1', 'Location2', 'Location3', 'Location4']
        },
        'brand_b': {
            'name': 'BrandB',
            'brand_filter': 'brand_b',
            'description': 'BrandB - Restaurant (Deliveroo)',
            'locations': ['Location1', 'Location2', 'Location4']
        },
        'brand_c': {
            'name': 'BrandC',
            'brand_filter': 'brand_c',
            'description': 'BrandC - Restaurant (Deliveroo)',
            'locations': ['Location1', 'Location2', 'Location4']
        },
        'brand_d': {
            'name': 'BrandD',
            'brand_filter': 'brand_d',
            'description': 'BrandD - Restaurant (Deliveroo)',
            'locations': ['Location1', 'Location2', 'Location4']
        },
        'brand_e': {
            'name': 'BrandE',
            'brand_filter': 'brand_e',
            'description': 'BrandE - Restaurant (Deliveroo)',
            'locations': ['Location1', 'Location2', 'Location4']
        },
        'brand_f': {
            'name': 'BrandF',
            'brand_filter': 'brand_f',
            'description': 'BrandF - Restaurant (Deliveroo)',
            'locations': ['Location1', 'Location3']
        }
    }

    def __init__(self, restaurant_key='brand_a'):
        self.restaurant_key = restaurant_key
        self.restaurant_config = self.RESTAURANTS.get(restaurant_key)
        self.pos_data = pd.DataFrame()
        self.system_data = pd.DataFrame()
        self.reconciliation_results = {
            'matched': [],
            'price_discrepancies': [],
            'missing_in_pos': [],
            'missing_in_system': [],
            'duplicates_in_pos': [],
            'duplicates_in_system': [],
            'summary': {}
        }

        if not self.restaurant_config:
            raise ValueError(f"Unknown restaurant: {restaurant_key}. Available: {list(self.RESTAURANTS.keys())}")

    def flag_duplicates(self, df, id_column='order_id', source='pos'):
        """Flag duplicate orders in dataframe without removing them"""
        df['is_duplicate'] = False
        df['duplicate_group'] = 0

        # Find duplicates based on order_id
        duplicates = df[df.duplicated(subset=[id_column], keep=False)]

        if len(duplicates) > 0:
            logger.info(f"🔄 Found {len(duplicates)} duplicate records in {source.upper()}")

            # Mark duplicates and assign group numbers
            group_num = 1
            for order_id in duplicates[id_column].unique():
                mask = df[id_column] == order_id
                df.loc[mask, 'is_duplicate'] = True
                df.loc[mask, 'duplicate_group'] = group_num

                # Log duplicate details
                dup_records = df[mask]
                logger.info(f"   Duplicate Group {group_num}: Order ID {order_id} appears {len(dup_records)} times")
                group_num += 1

            # Store duplicate information for reporting
            if source == 'pos':
                self.reconciliation_results['duplicates_in_pos'] = duplicates.to_dict('records')
            else:
                self.reconciliation_results['duplicates_in_system'] = duplicates.to_dict('records')

        return df

    def load_pos_data(self, filepath='data/pos_files/Grubtech POS.xlsx'):
        """Load and process POS data from Excel file"""
        try:
            restaurant_name = self.restaurant_config['name']
            brand_filter = self.restaurant_config['brand_filter']

            logger.info(f"📊 Loading POS data for {restaurant_name} (Deliveroo)...")

            # Check if file exists
            if not Path(filepath).exists():
                # Try alternative paths
                alt_paths = ['Grubtech POS.xlsx', 'data/Grubtech POS.xlsx']
                for alt_path in alt_paths:
                    if Path(alt_path).exists():
                        filepath = alt_path
                        break
                else:
                    raise FileNotFoundError(f"POS file not found: {filepath}")

            # Read the Excel file
            df = pd.read_excel(filepath, sheet_name='Sheet0')
            logger.info(f"📄 Raw POS data: {len(df)} records")

            # Filter for Deliveroo orders
            deliveroo_df = df[df['Channel'].str.lower() == 'deliveroo'].copy()
            logger.info(f"📊 Deliveroo orders: {len(deliveroo_df)} records")

            # Show brand distribution
            if len(deliveroo_df) > 0:
                logger.info("🏪 Deliveroo brands in POS:")
                brands = deliveroo_df['Brand'].value_counts()
                for brand, count in brands.head(10).items():
                    logger.info(f"   {brand}: {count}")

            # Filter for selected restaurant
            filtered_df = deliveroo_df[
                deliveroo_df['Brand'].str.lower().str.contains(brand_filter, na=False)
            ].copy()
            logger.info(f"📊 After '{brand_filter}' filter: {len(filtered_df)} records")

            # Structure data
            self.pos_data = pd.DataFrame({
                'brand': filtered_df['Brand'],
                'channel': filtered_df['Channel'],
                'location': filtered_df['Location'],
                'order_id': pd.to_numeric(filtered_df['Order ID'], errors='coerce'),
                'received_at': pd.to_datetime(filtered_df['Received At'], errors='coerce'),
                'customer_name': filtered_df['Customer Name'],
                'item_price': pd.to_numeric(filtered_df['Item Price'], errors='coerce'),
                'net_sales': pd.to_numeric(filtered_df['Net Sales'], errors='coerce'),
                'total': pd.to_numeric(filtered_df['Total(Receipt Total)'], errors='coerce'),
                'payment_method': filtered_df['Payment Method']
            })

            # Clean invalid order IDs
            before_cleanup = len(self.pos_data)
            self.pos_data = self.pos_data.dropna(subset=['order_id'])
            if len(self.pos_data) > 0:
                self.pos_data['order_id'] = self.pos_data['order_id'].astype(int)
            after_cleanup = len(self.pos_data)

            logger.info(f"📊 After removing invalid Order IDs: {after_cleanup} records (removed {before_cleanup - after_cleanup})")

            # Flag duplicates but keep them
            self.pos_data = self.flag_duplicates(self.pos_data, 'order_id', 'pos')

            logger.info(f"✅ Loaded {len(self.pos_data)} {restaurant_name} orders (including duplicates)")
            return self.pos_data

        except Exception as e:
            logger.error(f"❌ Error loading POS data: {e}")
            raise

    def load_system_data(self, csv_files=None):
        """Load and process Deliveroo system data from CSV files"""
        try:
            restaurant_name = self.restaurant_config['name']
            brand_filter = self.restaurant_config['brand_filter']

            logger.info(f"📊 Loading Deliveroo system data for {restaurant_name}...")

            # Find CSV files if not provided
            if csv_files is None:
                csv_paths = ['data/system_files/', 'data/', '.']
                deliveroo_files = []

                for path in csv_paths:
                    search_path = Path(path)
                    if search_path.exists():
                        files = list(search_path.glob("*rs-orders-report*.csv"))
                        deliveroo_files.extend(files)

                if not deliveroo_files:
                    raise FileNotFoundError("No Deliveroo CSV files found")

                csv_files = [str(f) for f in deliveroo_files]

            # Load all CSV files
            all_data = []
            for file_path in csv_files:
                logger.info(f"📄 Loading {Path(file_path).name}...")
                df = pd.read_csv(file_path)
                df['source_file'] = Path(file_path).name
                all_data.append(df)
                logger.info(f"   {len(df)} records")

            combined_df = pd.concat(all_data, ignore_index=True)
            logger.info(f"📊 Combined data: {len(combined_df)} records")

            # Structure system data
            self.system_data = pd.DataFrame({
                'order_id': pd.to_numeric(combined_df['Order number'], errors='coerce'),
                'datetime': pd.to_datetime(
                    combined_df['Date submitted'] + ' ' + combined_df['Time submitted'],
                    errors='coerce'
                ),
                'subtotal': pd.to_numeric(combined_df['Subtotal'], errors='coerce'),
                'commission': pd.to_numeric(combined_df['Deliveroo commission'], errors='coerce'),
                'restaurant_name': combined_df['Restaurant name'],
                'order_status': combined_df['Order status'],
                'source_file': combined_df['source_file']
            })

            # Filter for selected restaurant
            restaurant_filter = (
                self.system_data['restaurant_name'].str.lower().str.contains(brand_filter, na=False)
            )

            self.system_data = self.system_data[restaurant_filter].copy()
            logger.info(f"📊 After '{brand_filter}' filter: {len(self.system_data)} records")

            # Clean invalid order IDs
            before_cleanup = len(self.system_data)
            self.system_data = self.system_data.dropna(subset=['order_id'])
            if len(self.system_data) > 0:
                self.system_data['order_id'] = self.system_data['order_id'].astype(int)
            after_cleanup = len(self.system_data)

            logger.info(f"📊 After cleanup: {after_cleanup} records (removed {before_cleanup - after_cleanup})")

            # Flag duplicates but keep them
            self.system_data = self.flag_duplicates(self.system_data, 'order_id', 'system')

            logger.info(f"✅ Loaded {len(self.system_data)} {restaurant_name} orders (including duplicates)")
            return self.system_data

        except Exception as e:
            logger.error(f"❌ Error loading system data: {e}")
            raise

    def reconcile(self):
        """Perform reconciliation between POS and System data"""
        restaurant_name = self.restaurant_config['name']
        logger.info(f"\n🔄 Starting Deliveroo Reconciliation - {restaurant_name}")
        logger.info("=" * 60)

        # Tag duplicates with occurrence so each row is uniquely identifiable
        self.pos_data = add_occurrence_column(
            self.pos_data, id_col='order_id',
            sort_cols=('received_at', 'item_price', 'location')
        )
        self.system_data = add_occurrence_column(
            self.system_data, id_col='order_id',
            sort_cols=('datetime', 'subtotal', 'restaurant_name')
        )

        # Create lookup dictionaries with composite keys (order_id, occurrence)
        pos_dict = self.pos_data.set_index(['order_id', 'occurrence']).to_dict('index')
        system_dict = self.system_data.set_index(['order_id', 'occurrence']).to_dict('index')

        # Count duplicates
        pos_duplicates = int((self.pos_data['is_duplicate'] == True).sum())
        system_duplicates = int((self.system_data['is_duplicate'] == True).sum())

        logger.info(f"\n📊 DATA SUMMARY:")
        logger.info(f"   POS rows: {len(self.pos_data)} (including {pos_duplicates} duplicates)")
        logger.info(f"   System rows: {len(self.system_data)} (including {system_duplicates} duplicates)")

        # Find common composite keys
        pos_keys = set(pos_dict.keys())
        system_keys = set(system_dict.keys())
        common_keys = pos_keys.intersection(system_keys)

        logger.info(f"   Common pairs: {len(common_keys)}")
        logger.info(f"   Only in POS: {len(pos_keys - system_keys)}")
        logger.info(f"   Only in System: {len(system_keys - pos_keys)}")

        processed_keys = set()

        # Process POS rows by composite key
        for (order_id, occ), pos_order in pos_dict.items():
            processed_keys.add((order_id, occ))
            system_order = system_dict.get((order_id, occ))

            if system_order is None:
                # Missing in system
                self.reconciliation_results['missing_in_system'].append({
                    'order_id': order_id,
                    'occurrence': occ,
                    'pos_data': pos_order,
                    'is_duplicate': pos_order.get('is_duplicate', False),
                    'duplicate_group': pos_order.get('duplicate_group', 0)
                })
            else:
                # Order exists in both - check price match
                pos_price = pos_order.get('item_price', 0) or 0
                system_price = system_order.get('subtotal', 0) or 0
                price_difference = abs(pos_price - system_price)

                match_record = {
                    'order_id': order_id,
                    'occurrence': occ,
                    'pos_data': pos_order,
                    'system_data': system_order,
                    'pos_price': pos_price,
                    'system_price': system_price,
                    'price_difference': price_difference,
                    'price_match': price_difference < 0.01,
                    'pos_is_duplicate': pos_order.get('is_duplicate', False),
                    'system_is_duplicate': system_order.get('is_duplicate', False)
                }

                if match_record['price_match']:
                    self.reconciliation_results['matched'].append(match_record)
                else:
                    self.reconciliation_results['price_discrepancies'].append(match_record)

        # Find rows missing in POS
        for (order_id, occ), system_order in system_dict.items():
            if (order_id, occ) not in processed_keys:
                self.reconciliation_results['missing_in_pos'].append({
                    'order_id': order_id,
                    'occurrence': occ,
                    'system_data': system_order,
                    'is_duplicate': system_order.get('is_duplicate', False),
                    'duplicate_group': system_order.get('duplicate_group', 0)
                })

        # Generate summary
        total_orders = max(len(self.pos_data), len(self.system_data))
        self.reconciliation_results['summary'] = {
            'restaurant': restaurant_name,
            'channel': 'Deliveroo',
            'total_pos_orders': len(self.pos_data),
            'total_system_orders': len(self.system_data),
            'pos_duplicates': pos_duplicates,
            'system_duplicates': system_duplicates,
            'perfect_matches': len(self.reconciliation_results['matched']),
            'price_discrepancies': len(self.reconciliation_results['price_discrepancies']),
            'missing_in_pos': len(self.reconciliation_results['missing_in_pos']),
            'missing_in_system': len(self.reconciliation_results['missing_in_system']),
            'reconciliation_rate': (len(self.reconciliation_results['matched']) / total_orders * 100) if total_orders > 0 else 0
        }

        logger.info("✅ Deliveroo Reconciliation Complete!")
        return self.reconciliation_results

    def generate_report(self):
        """Generate reconciliation report"""
        restaurant_name = self.restaurant_config['name']
        logger.info(f"\n📋 DELIVEROO RECONCILIATION REPORT - {restaurant_name}")
        logger.info("=" * 70)

        summary = self.reconciliation_results['summary']

        logger.info("\n📊 SUMMARY STATISTICS:")
        logger.info(f"  Restaurant: {summary['restaurant']}")
        logger.info(f"  Channel: {summary['channel']}")
        logger.info(f"  Total POS Orders: {summary['total_pos_orders']}")
        logger.info(f"  Total System Orders: {summary['total_system_orders']}")
        logger.info(f"  POS Duplicates: {summary['pos_duplicates']} 🔄")
        logger.info(f"  System Duplicates: {summary['system_duplicates']} 🔄")
        logger.info(f"  Perfect Matches: {summary['perfect_matches']} ✅")
        logger.info(f"  Price Discrepancies: {summary['price_discrepancies']} ⚠️")
        logger.info(f"  Missing in POS: {summary['missing_in_pos']} ❌")
        logger.info(f"  Missing in System: {summary['missing_in_system']} ❌")
        logger.info(f"  Reconciliation Rate: {summary['reconciliation_rate']:.2f}%")

        return self.reconciliation_results

    def export_results(self, output_dir="deliveroo_output"):
        """Export reconciliation results to Excel"""
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        restaurant_name = self.restaurant_config['name'].replace(' ', '_')
        filename = f"deliveroo_reconciliation_{restaurant_name}_{timestamp}.xlsx"

        with pd.ExcelWriter(output_path / filename, engine='openpyxl') as writer:
            # Summary sheet
            summary_df = pd.DataFrame([self.reconciliation_results['summary']])
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Duplicates in POS
            if self.reconciliation_results['duplicates_in_pos']:
                dup_pos_df = pd.DataFrame(self.reconciliation_results['duplicates_in_pos'])
                dup_pos_df.to_excel(writer, sheet_name='Duplicates_POS', index=False)

            # Duplicates in System
            if self.reconciliation_results['duplicates_in_system']:
                dup_sys_df = pd.DataFrame(self.reconciliation_results['duplicates_in_system'])
                dup_sys_df.to_excel(writer, sheet_name='Duplicates_System', index=False)

            # Perfect matches
            if self.reconciliation_results['matched']:
                matches_data = []
                for match in self.reconciliation_results['matched']:
                    matches_data.append({
                        'Order_ID': match['order_id'],
                        'Occurrence': match.get('occurrence', 0),
                        'POS_Price': match['pos_price'],
                        'System_Price': match['system_price'],
                        'POS_Date': match['pos_data'].get('received_at'),
                        'System_Date': match['system_data'].get('datetime'),
                        'Customer': match['pos_data'].get('customer_name'),
                        'POS_Duplicate': match['pos_is_duplicate'],
                        'System_Duplicate': match['system_is_duplicate']
                    })
                matches_df = pd.DataFrame(matches_data)
                matches_df.to_excel(writer, sheet_name='Perfect_Matches', index=False)

            # Price discrepancies
            if self.reconciliation_results['price_discrepancies']:
                disc_data = []
                for disc in self.reconciliation_results['price_discrepancies']:
                    disc_data.append({
                        'Order_ID': disc['order_id'],
                        'Occurrence': disc.get('occurrence', 0),
                        'POS_Price': disc['pos_price'],
                        'System_Price': disc['system_price'],
                        'Difference': disc['price_difference'],
                        'POS_Date': disc['pos_data'].get('received_at'),
                        'System_Date': disc['system_data'].get('datetime'),
                        'POS_Duplicate': disc['pos_is_duplicate'],
                        'System_Duplicate': disc['system_is_duplicate']
                    })
                disc_df = pd.DataFrame(disc_data)
                disc_df.to_excel(writer, sheet_name='Price_Discrepancies', index=False)

            # Missing in System
            if self.reconciliation_results['missing_in_system']:
                missing_sys_data = []
                for missing in self.reconciliation_results['missing_in_system']:
                    missing_sys_data.append({
                        'Order_ID': missing['order_id'],
                        'Occurrence': missing.get('occurrence', 0),
                        'POS_Price': missing['pos_data'].get('item_price'),
                        'POS_Date': missing['pos_data'].get('received_at'),
                        'Customer': missing['pos_data'].get('customer_name'),
                        'Location': missing['pos_data'].get('location'),
                        'Is_Duplicate': missing['is_duplicate'],
                        'Duplicate_Group': missing['duplicate_group']
                    })
                missing_sys_df = pd.DataFrame(missing_sys_data)
                missing_sys_df.to_excel(writer, sheet_name='Missing_in_System', index=False)

            # Missing in POS
            if self.reconciliation_results['missing_in_pos']:
                missing_pos_data = []
                for missing in self.reconciliation_results['missing_in_pos']:
                    missing_pos_data.append({
                        'Order_ID': missing['order_id'],
                        'Occurrence': missing.get('occurrence', 0),
                        'System_Price': missing['system_data'].get('subtotal'),
                        'System_Date': missing['system_data'].get('datetime'),
                        'Restaurant': missing['system_data'].get('restaurant_name'),
                        'Status': missing['system_data'].get('order_status'),
                        'Is_Duplicate': missing['is_duplicate'],
                        'Duplicate_Group': missing['duplicate_group']
                    })
                missing_pos_df = pd.DataFrame(missing_pos_data)
                missing_pos_df.to_excel(writer, sheet_name='Missing_in_POS', index=False)

        output_file = output_path / filename
        logger.info(f"💾 Results exported to: {output_file}")
        return str(output_file)

def main(restaurant_key=None):
    """Main execution function"""
    try:
        logger.info("🚀 Starting Deliveroo POS-System Reconciliation")
        logger.info("=" * 60)

        # Use default restaurant if none specified
        if not restaurant_key:
            restaurant_key = 'brand_a'
            logger.info(f"🎯 Using default restaurant: {restaurant_key}")

        # Initialize reconciler
        reconciler = DeliverooReconciliation(restaurant_key)

        # Load data
        reconciler.load_pos_data()
        reconciler.load_system_data()

        # Perform reconciliation
        results = reconciler.reconcile()

        # Generate report
        report = reconciler.generate_report()

        # Export results
        output_file = reconciler.export_results()

        logger.info("\n🎉 Deliveroo reconciliation completed!")
        logger.info(f"📊 Results exported to: {output_file}")

        # Print final summary
        summary = results['summary']
        logger.info(f"\n📈 FINAL SUMMARY - {summary['restaurant']}:")
        logger.info(f"  Match Rate: {summary['reconciliation_rate']:.2f}%")
        logger.info(f"  Duplicates: POS={summary['pos_duplicates']}, System={summary['system_duplicates']}")
        logger.info(f"  Total Issues: {summary['price_discrepancies'] + summary['missing_in_pos'] + summary['missing_in_system']}")

        return reconciler, results

    except Exception as e:
        logger.error(f"❌ Reconciliation failed: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def process_all_restaurants():
    """Process all restaurants"""
    logger.info("🏪 PROCESSING ALL DELIVEROO RESTAURANTS")
    logger.info("=" * 50)

    results_summary = []

    for restaurant_key in DeliverooReconciliation.RESTAURANTS.keys():
        try:
            logger.info(f"\n🔄 Processing {restaurant_key}...")
            reconciler, results = main(restaurant_key)

            if results:
                results_summary.append({
                    'restaurant': restaurant_key,
                    'summary': results['summary']
                })

        except Exception as e:
            logger.error(f"❌ Failed to process {restaurant_key}: {e}")
            continue

    # Generate combined summary
    if results_summary:
        logger.info("\n📊 COMBINED DELIVEROO RECONCILIATION SUMMARY")
        logger.info("=" * 60)

        total_pos = sum(r['summary']['total_pos_orders'] for r in results_summary)
        total_system = sum(r['summary']['total_system_orders'] for r in results_summary)
        total_matches = sum(r['summary']['perfect_matches'] for r in results_summary)

        logger.info(f"📈 OVERALL STATISTICS:")
        logger.info(f"  Total POS Orders: {total_pos}")
        logger.info(f"  Total System Orders: {total_system}")
        logger.info(f"  Total Perfect Matches: {total_matches}")
        logger.info(f"  Overall Match Rate: {(total_matches / max(total_pos, total_system) * 100) if max(total_pos, total_system) > 0 else 0:.2f}%")

        logger.info(f"\n📋 BY RESTAURANT:")
        for result in results_summary:
            summary = result['summary']
            logger.info(f"  {result['restaurant']}: {summary['reconciliation_rate']:.2f}% match ({summary['perfect_matches']}/{summary['total_pos_orders']})")

    return results_summary

if __name__ == "__main__":
    print("🚀 Deliveroo Reconciliation Script")
    print("=" * 50)

    try:
        # Check command line arguments
        restaurant_key = None
        if len(sys.argv) > 1:
            if sys.argv[1].lower() == 'all':
                print("🏪 Processing ALL restaurants...")
                process_all_restaurants()
            else:
                restaurant_key = sys.argv[1].lower()
                reconciler, results = main(restaurant_key)
        else:
            reconciler, results = main(restaurant_key)

        if results:
            print("✅ Script completed successfully!")
        else:
            print("❌ Script completed with errors.")

    except Exception as e:
        print(f"❌ Error occurred: {e}")
        import traceback
        traceback.print_exc()
