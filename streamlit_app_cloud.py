#!/usr/bin/env python3
"""
T&T Purchase Order Processor - Version with Embedded Store Names
Upload purchase_orders.xlsx and product_variants.xlsx to convert to Odoo format.
Store names are now embedded in the code.
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import logging
from typing import Dict, List, Tuple, Optional
from io import BytesIO, StringIO
import subprocess
import sys
import re

# Embedded Store Names - No need to upload store names file anymore!
EMBEDDED_STORE_NAMES = {
    1: "T&T Supermarket Inc., Metrotown  - 001",
    3: "T&T Supermarket Inc., Chinatown  - 003",
    4: "T&T Supermarket Inc., First Avenue  - 004",
    5: "T&T Supermarket Inc., Osaka  - 005",
    6: "T&T Supermarket Inc., Surrey  - 006",
    7: "T&T Supermarket Inc., Calgary  - 007",
    8: "T&T Supermarket Inc., Coquitlam  - 008",
    9: "T&T Supermarket Inc., Promenade Store - 009",
    10: "T&T Supermarket Inc., Edmonton  - 010",
    11: "T&T Supermarket Inc., Warden&Steels Store - 011",
    13: "T&T Supermarket Inc., Central City  - 013",
    14: "T&T Supermarket Inc., Harvest Hills  - 014",
    15: "T&T Supermarket Inc., Central Parkway Store - 015",
    17: "T&T Supermarket Inc., Northtown Edmonton  - 017",
    18: "T&T Supermarket Inc., Ottawa Store - 018",
    19: "T&T Supermarket Inc., Park Royal  - 019",
    20: "T&T Supermarket Inc., Weldrick Store - 020",
    21: "T&T Supermarket Inc., Woodbine Store - 021",
    22: "T&T Supermarket Inc., Unionville Store - 022",
    23: "T&T Supermarket Inc., Ora  - 023",
    24: "T&T Supermarket Inc., SE Edmonton  - 024",
    25: "T&T Supermarket Inc., Marine Gateway  - 025",
    26: "T&T Supermarket Inc., Lansdowne  - 026",
    27: "T&T Supermarket Inc., Aurora Store - 027",
    28: "T&T Supermarket Inc., Waterloo Store - 028",
    29: "T&T Supermarket Inc., Kingsway  - 029",
    30: "T&T Supermarket Inc., Deerfoot  - 030",
    31: "T&T Supermarket Inc., Langley  - 031",
    32: "T&T Supermarket Inc., College Store - 032",
    33: "T&T Supermarket Inc., Sage Hill  - 033",
    34: "T&T Supermarket Inc., St.Croix Store - 034",
    35: "T&T Supermarket Inc., Fairview Mall Store - 035",
    36: "T&T Supermarket Inc., Lougheed  - 036",
    37: "T&T Supermarket Inc., London Store - 037",
    38: "T&T Supermarket Inc., Downtown Store - 038",
    39: "T&T Supermarket Inc., Kanata Store - 039",
    40: "T&T Supermarket Inc., Brossard Store - 040"
}

def get_embedded_store_names() -> pd.DataFrame:
    """Return the embedded store names as a DataFrame"""
    store_data = [
        {'Store ID': store_id, 'Store Official Name': official_name}
        for store_id, official_name in EMBEDDED_STORE_NAMES.items()
    ]
    return pd.DataFrame(store_data)

# Force install required packages if not available
def install_package(package):
    """Install a package if not available"""
    try:
        __import__(package)
        return True
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            return True
        except:
            return False

# Try to install and import Excel libraries
if not install_package("openpyxl"):
    st.error("Failed to install openpyxl. Please check your requirements.txt file.")

# Try to import Excel reading libraries
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.error("openpyxl is not available. Please ensure it's in your requirements.txt file.")

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="T&T Purchase Order Processor",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        background: linear-gradient(90deg, #1f77b4, #ff7f0e);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 1rem;
        padding: 0.5rem;
        background: linear-gradient(135deg, #ecf0f1, #d5dbdb);
        border-radius: 8px;
        border-left: 4px solid #3498db;
    }
    .success-box {
        background: linear-gradient(135deg, #d4edda, #c3e6cb);
        border: 1px solid #c3e6cb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background: linear-gradient(135deg, #fff3cd, #ffeaa7);
        border: 1px solid #ffeaa7;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background: linear-gradient(135deg, #f8d7da, #f5c6cb);
        border: 1px solid #f5c6cb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .embedded-info {
        background: linear-gradient(135deg, #e8f4fd, #d1ecf1);
        border: 1px solid #bee5eb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
        border-left: 4px solid #17a2b8;
    }
</style>
""", unsafe_allow_html=True)

def read_excel_file(file) -> pd.DataFrame:
    """Read Excel file with fallback options for different engines"""
    try:
        # Try with openpyxl first (for .xlsx files)
        if OPENPYXL_AVAILABLE:
            try:
                return pd.read_excel(file, engine='openpyxl')
            except Exception as e:
                logger.warning(f"openpyxl failed: {e}")
        
        # Try with default engine
        try:
            return pd.read_excel(file)
        except Exception as e:
            logger.warning(f"default engine failed: {e}")
        
        # If all else fails, try with specific engine based on file extension
        file_name = file.name.lower()
        if file_name.endswith('.xlsx'):
            if OPENPYXL_AVAILABLE:
                return pd.read_excel(file, engine='openpyxl')
            else:
                raise Exception("openpyxl is required for .xlsx files but not available")
        else:
            raise Exception("Unsupported file format")
            
    except Exception as e:
        raise Exception(f"Failed to read Excel file: {e}")

def read_csv_file(file) -> pd.DataFrame:
    """Read CSV file as an alternative to Excel"""
    try:
        # Try different encodings
        encodings = ['utf-8', 'latin-1', 'cp1252']
        for encoding in encodings:
            try:
                file.seek(0)  # Reset file pointer
                return pd.read_csv(file, encoding=encoding)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                logger.warning(f"CSV reading failed with {encoding}: {e}")
                continue
        
        raise Exception("Failed to read CSV file with any encoding")
    except Exception as e:
        raise Exception(f"Failed to read CSV file: {e}")

def validate_and_reorder_columns(df: pd.DataFrame, expected_columns: List[str]) -> pd.DataFrame:
    """Validate and reorder columns, handling missing columns gracefully"""
    existing_columns = df.columns.tolist()
    missing_columns = [col for col in expected_columns if col not in existing_columns]
    
    if missing_columns:
        st.warning(f"⚠️ Some expected columns are missing: {missing_columns}")
        st.info("📋 Available columns: " + ", ".join(existing_columns))
        
        # Handle common column mapping
        if 'Internal Reference' not in existing_columns and 'Item#' in existing_columns:
            df['Internal Reference'] = df['Item#']
            st.info("✅ Mapped 'Item#' to 'Internal Reference'")
        
        if '# of Order' not in existing_columns and 'Ordered Qty' in existing_columns:
            df['# of Order'] = df['Ordered Qty']
            st.info("✅ Mapped 'Ordered Qty' to '# of Order'")
        
        # Re-check for missing columns after mapping
        existing_columns = df.columns.tolist()
        missing_columns = [col for col in expected_columns if col not in existing_columns]
        
        if missing_columns:
            st.error(f"❌ Still missing required columns: {missing_columns}")
            st.error("Please ensure your data contains all required columns or use the correct file format.")
            return df
    
    # Reorder columns, but only use existing ones
    available_columns = [col for col in expected_columns if col in df.columns]
    if available_columns:
        df = df[available_columns]
    
    return df

class OdooConverter:
    """Odoo conversion functionality with embedded store names"""
    
    def __init__(self, purchase_orders: pd.DataFrame, product_variants: pd.DataFrame):
        self.purchase_orders = purchase_orders
        self.product_variants = product_variants
        self.store_names = get_embedded_store_names()  # Use embedded store names
        self.order_summaries = None
        self.order_line_details = None
        self.expanded_orders = None
        
        # Data preparation
        self._prepare_data()
        
    def _prepare_data(self):
        """Prepare data for processing"""
        # Clean column names and handle specific column name issues
        self.purchase_orders.columns = self.purchase_orders.columns.str.strip()
        self.product_variants.columns = self.product_variants.columns.str.strip()
        
        # Convert Internal Reference to string for consistent matching
        self.purchase_orders['Internal Reference'] = self.purchase_orders['Internal Reference'].astype(str)
        self.product_variants['Internal Reference'] = self.product_variants['Internal Reference'].astype(str)
        
        st.write("🔍 Debug: Data preparation completed")
        st.write(f"Purchase Orders columns: {list(self.purchase_orders.columns)}")
        st.write(f"Product Variants columns: {list(self.product_variants.columns)}")
        st.success(f"✅ Using embedded store names ({len(self.store_names)} stores)")
        
    def extract_store_id_from_official_name(self, official_name: str) -> str:
        """Extract store ID from official store name"""
        # Pattern: "T&T Supermarket Inc., Store Name - XXX"
        match = re.search(r'- (\d+)$', official_name)
        if match:
            return match.group(1)
        return None
        
    def match_store_names(self) -> List[str]:
        """Match store names with official names using embedded store data"""
        errors = []
        
        # Create a mapping from store ID to official name using embedded data
        store_mapping = dict(zip(self.store_names['Store ID'], self.store_names['Store Official Name']))
        
        # Add official store name to purchase orders
        self.purchase_orders['Store Official Name'] = self.purchase_orders['Store ID'].map(store_mapping)
        
        # Log unmatched stores
        unmatched_stores = self.purchase_orders[self.purchase_orders['Store Official Name'].isna()]['Store ID'].unique()
        if len(unmatched_stores) > 0:
            error_msg = f"Unmatched store IDs: {unmatched_stores}"
            errors.append(error_msg)
            st.warning(f"⚠️ {error_msg}")
            
            # Show available store IDs for reference
            available_stores = sorted(self.store_names['Store ID'].unique())
            st.info(f"📋 Available store IDs in embedded data: {available_stores}")
        else:
            st.success("✅ All store IDs successfully matched with embedded store names")
        
        # Log successful matches
        matched_count = self.purchase_orders['Store Official Name'].notna().sum()
        st.info(f"📊 Successfully matched {matched_count} out of {len(self.purchase_orders)} order lines")
        
        return errors
    
    def create_order_summaries(self):
        """Create order summaries by store"""
        # Group by store and aggregate data
        summaries = []
        order_ref_counter = 6  # Start with OATS000006
        
        for store_id in sorted(self.purchase_orders['Store ID'].unique()):
            store_data = self.purchase_orders[self.purchase_orders['Store ID'] == store_id]
            
            # Get store information
            store_name = store_data['Store Name'].iloc[0]
            official_name = store_data['Store Official Name'].iloc[0]
            
            # Get all PO numbers for this store
            po_numbers = sorted(store_data['PO No.'].unique())
            po_numbers_str = ', '.join(map(str, po_numbers))
            
            # Get earliest order and delivery dates
            earliest_order_date = store_data['Order Date'].min()
            earliest_delivery_date = store_data['Delivery Date'].min()
            
            # Create order reference
            order_ref = f"OATS{order_ref_counter:06d}"
            order_ref_counter += 1
            
            summaries.append({
                'Order Reference': order_ref,
                'Customer Official Name': official_name if pd.notna(official_name) else f"Store {store_id} - {store_name}",
                'Store ID': store_id,
                'Store Name': store_name,
                'Order Date': earliest_order_date,
                'Delivery Date': earliest_delivery_date,
                'PO Numbers': po_numbers_str,
                'Total PO Count': len(po_numbers)
            })
        
        self.order_summaries = pd.DataFrame(summaries)
        st.success(f"✅ Created {len(self.order_summaries)} order summaries")
        
    def handle_multi_product_references(self) -> List[str]:
        """Handle internal references that link to multiple products"""
        errors = []
        
        # Check if required columns exist
        required_po_columns = ['Internal Reference', '# of Order', 'Price']
        required_pv_columns = ['Internal Reference', 'Barcode', 'Name', 'Units Per Order']
        
        missing_po_cols = [col for col in required_po_columns if col not in self.purchase_orders.columns]
        missing_pv_cols = [col for col in required_pv_columns if col not in self.product_variants.columns]
        
        if missing_po_cols:
            error_msg = f"Missing columns in purchase orders: {missing_po_cols}"
            errors.append(error_msg)
            st.error(f"❌ {error_msg}")
        
        if missing_pv_cols:
            error_msg = f"Missing columns in product variants: {missing_pv_cols}"
            errors.append(error_msg)
            st.error(f"❌ {error_msg}")
        
        if missing_po_cols or missing_pv_cols:
            self.expanded_orders = pd.DataFrame()
            return errors
        
        # Find internal references with multiple products
        ref_counts = self.product_variants['Internal Reference'].value_counts()
        multi_product_refs = ref_counts[ref_counts > 1].index.tolist()
        
        st.info(f"📊 Found {len(multi_product_refs)} internal references with multiple products")
        
        # Check for matching references
        po_internal_refs = set(self.purchase_orders['Internal Reference'].unique())
        pv_internal_refs = set(self.product_variants['Internal Reference'].unique())
        
        matching_refs = po_internal_refs & pv_internal_refs
        unmatched_refs = po_internal_refs - pv_internal_refs
        
        st.info(f"📊 Found {len(matching_refs)} matching internal references")
        
        if unmatched_refs:
            error_msg = f"Unmatched internal references: {list(unmatched_refs)[:10]}..."
            errors.append(error_msg)
            st.warning(f"⚠️ {error_msg}")
        
        if len(matching_refs) == 0:
            error_msg = "No matching internal references found between purchase orders and product variants"
            errors.append(error_msg)
            st.error(f"❌ {error_msg}")
            self.expanded_orders = pd.DataFrame()
            return errors
        
        # Create expanded purchase orders for multi-product references
        expanded_orders = []
        
        for _, row in self.purchase_orders.iterrows():
            internal_ref = row['Internal Reference']
            
            if internal_ref in multi_product_refs:
                # Get all products for this internal reference
                products = self.product_variants[self.product_variants['Internal Reference'] == internal_ref]
                
                if len(products) == 0:
                    error_msg = f"No products found for multi-product reference: {internal_ref}"
                    errors.append(error_msg)
                    continue
                
                # Calculate units per product (distribute equally)
                try:
                    total_units = row['# of Order'] * products.iloc[0]['Units Per Order']
                    units_per_product = total_units / len(products)
                    
                    # Create a line for each product
                    for i, (_, product) in enumerate(products.iterrows()):
                        # Distribute units as evenly as possible
                        if i == 0:
                            # First product gets the remainder
                            actual_units = int(units_per_product) + (total_units % len(products))
                        else:
                            actual_units = int(units_per_product)
                        
                        # Calculate unit price - IMPROVED CALCULATION
                        unit_price = row['Price'] / product['Units Per Order']
                        
                        expanded_orders.append({
                            'Store ID': row['Store ID'],
                            'Store Name': row['Store Name'],
                            'Store Official Name': row['Store Official Name'],
                            'PO No.': row['PO No.'],
                'Order Date': row['Order Date'],
                'Delivery Date': row['Delivery Date']
            })
        
        if line_details:
            self.order_line_details = pd.DataFrame(line_details)
            st.success(f"✅ Created {len(self.order_line_details)} order line details")
        else:
            st.error("❌ No order line details were created")
            self.order_line_details = pd.DataFrame()
    
    def validate_data(self) -> List[str]:
        """Validate the processed data"""
        errors = []
        
        # Check for missing data
        if self.purchase_orders is not None:
            missing_official_names = self.purchase_orders['Store Official Name'].isna().sum()
            if missing_official_names > 0:
                error_msg = f"Missing official store names: {missing_official_names}"
                errors.append(error_msg)
                st.warning(f"⚠️ {error_msg}")
        
        # Check for unmatched internal references
        if self.purchase_orders is not None and self.product_variants is not None:
            po_refs = set(self.purchase_orders['Internal Reference'].unique())
            pv_refs = set(self.product_variants['Internal Reference'].unique())
            unmatched_refs = po_refs - pv_refs
            if unmatched_refs:
                error_msg = f"Unmatched internal references: {list(unmatched_refs)[:10]}..."
                errors.append(error_msg)
                st.warning(f"⚠️ {error_msg}")
        
        # Validate calculations
        if (self.order_line_details is not None and not self.order_line_details.empty and 
            self.purchase_orders is not None and 'Total Price' in self.order_line_details.columns):
            
            total_price_check = abs(self.order_line_details['Total Price'].sum() - 
                                   self.purchase_orders['Price'].sum())
            if total_price_check > 0.01:  # Allow small rounding differences
                error_msg = f"Price mismatch: ${total_price_check:.2f}"
                errors.append(error_msg)
                st.warning(f"⚠️ {error_msg}")
        
        if not errors:
            st.success("✅ Data validation completed successfully")
        
        return errors
    
    def generate_summary_report(self):
        """Generate a summary report of the conversion"""
        if self.order_summaries is None or self.order_line_details is None:
            st.error("❌ Cannot generate summary report - missing data")
            return
        
        report = {
            'Total Stores': len(self.order_summaries),
            'Total PO Numbers': self.purchase_orders['PO No.'].nunique(),
            'Total Order Lines (Original)': len(self.purchase_orders),
            'Total Order Lines (Expanded)': len(self.order_line_details),
            'Multi-Product References': len(self.expanded_orders[self.expanded_orders['Is Multi Product'] == True]['Internal Reference'].unique()) if self.expanded_orders is not None else 0,
            'Total Value': self.order_line_details['Total Price'].sum(),
            'Average Order Value': self.order_line_details['Total Price'].mean(),
            'Embedded Stores Available': len(EMBEDDED_STORE_NAMES)
        }
        
        st.markdown("### 📊 Conversion Summary Report")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Stores", report['Total Stores'])
            st.metric("Total PO Numbers", report['Total PO Numbers'])
        
        with col2:
            st.metric("Original Order Lines", report['Total Order Lines (Original)'])
            st.metric("Expanded Order Lines", report['Total Order Lines (Expanded)'])
        
        with col3:
            st.metric("Multi-Product References", report['Multi-Product References'])
            st.metric("Total Value", f"${report['Total Value']:,.2f}")
        
        with col4:
            st.metric("Average Order Value", f"${report['Average Order Value']:,.2f}")
            st.metric("Embedded Stores", report['Embedded Stores Available'])
    
    def process_all(self) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
        """Run the complete conversion process"""
        errors = []
        
        # Match store names (now using embedded data)
        store_errors = self.match_store_names()
        errors.extend(store_errors)
        
        # Create order summaries
        self.create_order_summaries()
        
        # Handle multi-product references
        ref_errors = self.handle_multi_product_references()
        errors.extend(ref_errors)
        
        # Create order line details
        self.create_order_line_details()
        
        # Validate data
        validation_errors = self.validate_data()
        errors.extend(validation_errors)
        
        # Generate summary report
        self.generate_summary_report()
        
        return self.order_summaries, self.order_line_details, errors

def main():
    """Main Streamlit application with embedded store names"""
    
    # Header
    st.markdown('<h1 class="main-header">🛒 T&T Purchase Order Processor</h1>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Information about embedded store names
    st.markdown("""
    <div class="embedded-info">
        <h4>✨ Enhanced Version with Embedded Store Names</h4>
        <p>🎉 <strong>No need to upload store names file anymore!</strong> All 37 T&T store names are now embedded in the application.</p>
        <p>📋 Simply upload your <strong>Purchase Orders</strong> and <strong>Product Variants</strong> files to get started.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # File uploads in columns (only 2 columns now since store names are embedded)
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<h3 class="section-header">📦 Product Variants</h3>', unsafe_allow_html=True)
        
        # File type selection
        product_file_type = st.radio(
            "Select file type:",
            ["Excel (.xlsx)", "CSV (.csv)"],
            key="product_variants_type"
        )
        
        if product_file_type == "Excel (.xlsx)":
            product_variants_file = st.file_uploader(
                "Upload Product Variants file",
                type=['xlsx'],
                key="product_variants"
            )
        else:
            product_variants_file = st.file_uploader(
                "Upload Product Variants file",
                type=['csv'],
                key="product_variants"
            )
        
        product_variants = None
        if product_variants_file:
            try:
                if product_file_type == "Excel (.xlsx)":
                    product_variants = read_excel_file(product_variants_file)
                else:
                    product_variants = read_csv_file(product_variants_file)
                
                if not product_variants.empty:
                    st.success(f"✅ Loaded {len(product_variants)} products")
                    st.dataframe(product_variants.head(3), use_container_width=True)
                else:
                    st.error("❌ File is empty")
            except Exception as e:
                st.error(f"❌ Error loading file: {e}")
    
    with col2:
        st.markdown('<h3 class="section-header">🛒 Purchase Orders</h3>', unsafe_allow_html=True)
        
        # File type selection
        orders_file_type = st.radio(
            "Select file type:",
            ["Excel (.xlsx)", "CSV (.csv)"],
            key="purchase_orders_type"
        )
        
        if orders_file_type == "Excel (.xlsx)":
            purchase_orders_file = st.file_uploader(
                "Upload Purchase Orders file",
                type=['xlsx'],
                key="purchase_orders"
            )
        else:
            purchase_orders_file = st.file_uploader(
                "Upload Purchase Orders file",
                type=['csv'],
                key="purchase_orders"
            )
        
        purchase_orders = None
        if purchase_orders_file:
            try:
                if orders_file_type == "Excel (.xlsx)":
                    purchase_orders = read_excel_file(purchase_orders_file)
                else:
                    purchase_orders = read_csv_file(purchase_orders_file)
                
                if not purchase_orders.empty:
                    # Clean column names
                    purchase_orders.columns = purchase_orders.columns.str.strip()
                    if '# of Order ' in purchase_orders.columns:
                        purchase_orders = purchase_orders.rename(columns={'# of Order ': '# of Order'})
                    
                    # Convert to numeric for proper sorting
                    purchase_orders['Store ID'] = pd.to_numeric(purchase_orders['Store ID'], errors='coerce')
                    purchase_orders['PO No.'] = pd.to_numeric(purchase_orders['PO No.'], errors='coerce')
                    
                    # Sort by Store ID and PO No.
                    purchase_orders = purchase_orders.sort_values(by=['Store ID', 'PO No.'], ascending=[True, True])
                    
                    # Validate and reorder columns
                    expected_columns = ['Store ID', 'Store Name', 'PO No.', 'Order Date', 'Delivery Date',
                                      'Internal Reference', '# of Order', 'Price']
                    purchase_orders = validate_and_reorder_columns(purchase_orders, expected_columns)
                    
                    st.success(f"✅ Loaded {len(purchase_orders)} orders")
                    st.dataframe(purchase_orders.head(3), use_container_width=True)
                else:
                    st.error("❌ File is empty")
            except Exception as e:
                st.error(f"❌ Error loading file: {e}")
    
    # Show embedded store names information
    st.markdown("---")
    with st.expander("📋 View Embedded Store Names", expanded=False):
        embedded_stores = get_embedded_store_names()
        st.dataframe(embedded_stores, use_container_width=True)
        st.info(f"📊 Total embedded stores: {len(embedded_stores)}")
    
    # Processing section
    st.markdown("---")
    st.markdown('<h2 class="section-header">🔄 Process & Convert</h2>', unsafe_allow_html=True)
    
    # Check if required files are uploaded (only 2 files needed now)
    if product_variants is not None and purchase_orders is not None:
        if st.button("🚀 Convert to Odoo Format", type="primary", use_container_width=True):
            with st.spinner("Converting to Odoo format using embedded store names..."):
                try:
                    # Initialize converter (no need to pass store_names anymore)
                    converter = OdooConverter(purchase_orders, product_variants)
                    
                    # Process conversion
                    order_summaries, order_line_details, errors = converter.process_all()
                    
                    # Display results
                    st.success("✅ Conversion completed successfully!")
                    
                    # Show order summaries
                    if order_summaries is not None and not order_summaries.empty:
                        with st.expander("📋 Order Summaries", expanded=True):
                            st.dataframe(order_summaries, use_container_width=True)
                    else:
                        st.error("❌ No order summaries were generated")
                    
                    # Show order line details
                    if order_line_details is not None and not order_line_details.empty:
                        with st.expander("📊 Order Line Details (First 20 rows)", expanded=False):
                            st.dataframe(order_line_details.head(20), use_container_width=True)
                    else:
                        st.error("❌ No order line details were generated")
                    
                    # Show errors if any
                    if errors:
                        with st.expander("⚠️ Conversion Warnings", expanded=False):
                            for error in errors[:10]:
                                st.warning(error)
                            if len(errors) > 10:
                                st.info(f"... and {len(errors) - 10} more warnings")
                    
                    # Create Excel file for download only if we have data
                    if order_summaries is not None and order_line_details is not None and not order_line_details.empty:
                        with st.spinner("Preparing download file..."):
                            excel_buffer = BytesIO()
                            
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                # Save order summaries
                                order_summaries.to_excel(writer, sheet_name='Order Summaries', index=False)
                                
                                # Save order line details
                                order_line_details.to_excel(writer, sheet_name='Order Line Details', index=False)
                                
                                # Save original data for reference
                                purchase_orders.to_excel(writer, sheet_name='Original Purchase Orders', index=False)
                                product_variants.to_excel(writer, sheet_name='Product Variants', index=False)
                                
                                # Save embedded store names for reference
                                get_embedded_store_names().to_excel(writer, sheet_name='Embedded Store Names', index=False)
                            
                            excel_buffer.seek(0)
                        
                        # Download section
                        st.markdown("---")
                        st.markdown('<h2 class="section-header">📥 Download Results</h2>', unsafe_allow_html=True)
                        
                        # Download button
                        st.download_button(
                            label="📥 Download Odoo_Import_Ready.xlsx",
                            data=excel_buffer.getvalue(),
                            file_name="Odoo_Import_Ready.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                        
                        st.info("📋 The downloaded file contains:")
                        st.markdown("""
                        - **Order Summaries**: Summary of orders by store
                        - **Order Line Details**: Detailed product lines for Odoo import
                        - **Original Purchase Orders**: Raw extracted data
                        - **Product Variants**: Reference product data
                        - **Embedded Store Names**: Reference store data (embedded in app)
                        """)
                    else:
                        st.error("❌ Cannot generate download file - no valid data was created")
                
                except Exception as e:
                    st.error(f"❌ Error during conversion: {e}")
                    st.error("Please check your file formats and try again.")
                    
                    # Additional debugging info
                    st.write("📊 Debug Info:")
                    st.write(f"Purchase Orders shape: {purchase_orders.shape}")
                    st.write(f"Product Variants shape: {product_variants.shape}")
                    st.write(f"Purchase Orders columns: {list(purchase_orders.columns)}")
                    st.write(f"Product Variants columns: {list(product_variants.columns)}")
                    
                    # Show first few rows of each dataset
                    st.write("Purchase Orders sample:")
                    st.dataframe(purchase_orders.head(3))
                    st.write("Product Variants sample:")
                    st.dataframe(product_variants.head(3))
    
    else:
        st.info("📋 Please upload both required files to proceed with conversion")
        
        missing_files = []
        if product_variants is None:
            missing_files.append("Product Variants")
        if purchase_orders is None:
            missing_files.append("Purchase Orders")
        
        if missing_files:
            st.warning(f"⚠️ Missing files: {', '.join(missing_files)}")
        
        st.success("✅ Store Names: Using embedded data (no upload required)")
    
    # Help section
    with st.sidebar:
        st.markdown("## ℹ️ Help & Instructions")
        st.markdown("""
        **✨ Enhanced Version Features:**
        - 🎉 **No Store Names Upload Required!** All 37 T&T store names are embedded
        - 🚀 Faster processing with only 2 file uploads needed
        - 📊 Enhanced validation and error reporting
        
        **How to use this tool:**
        
        1. **Upload Files**: Upload only 2 files now:
           - Product Variants (Excel/CSV)
           - Purchase Orders (Excel/CSV)
        
        2. **Convert**: Click "Convert to Odoo Format" button
        
        3. **Download**: Download the "Odoo_Import_Ready.xlsx" file
        
        **Required File Formats:**
        - **Product Variants**: Must contain columns like 'Internal Reference', 'Barcode', 'Name', 'Units Per Order'
        - **Purchase Orders**: Must contain 'Store ID', 'Store Name', 'PO No.', 'Order Date', 'Delivery Date', 'Internal Reference', '# of Order', 'Price'
        
        **Embedded Store Names:**
        - ✅ All 37 T&T store locations included
        - 🔄 Automatic matching by Store ID
        - 📋 View complete list in the expandable section above
        
        **Features:**
        - Automatic product mapping
        - Multi-product reference handling
        - Comprehensive error reporting
        - Odoo-compatible output format
        """)
        
        # Show quick stats about embedded stores
        st.markdown("### 📊 Embedded Store Stats")
        embedded_stores = get_embedded_store_names()
        st.metric("Total Stores", len(embedded_stores))
        st.metric("Store ID Range", f"{embedded_stores['Store ID'].min()}-{embedded_stores['Store ID'].max()}")
        
        # Show some example store names
        st.markdown("**Example Stores:**")
        sample_stores = embedded_stores.sample(min(3, len(embedded_stores)))
        for _, store in sample_stores.iterrows():
            st.write(f"• Store {store['Store ID']}: {store['Store Official Name'].split(' - ')[0].split(', ')[-1]}")

if __name__ == "__main__":
    main()PO No.'],
                            'Order Date': row['Order Date'],
                            'Delivery Date': row['Delivery Date'],
                            'Internal Reference': internal_ref,
                            'Barcode': product['Barcode'],
                            'Product Name': product['Name'],
                            'Units Per Order': product['Units Per Order'],
                            'Original Order Quantity': row['# of Order'],
                            'Total Units': actual_units,
                            'Unit Price': unit_price,
                            'Total Price': actual_units * unit_price,
                            'Is Multi Product': True
                        })
                
                except Exception as e:
                    error_msg = f"Error processing multi-product reference {internal_ref}: {e}"
                    errors.append(error_msg)
                    
            else:
                # Single product reference - keep as is
                product = self.product_variants[self.product_variants['Internal Reference'] == internal_ref]
                if len(product) > 0:
                    product = product.iloc[0]
                    try:
                        total_units = row['# of Order'] * product['Units Per Order']
                        # Calculate unit price - IMPROVED CALCULATION
                        unit_price = row['Price'] / product['Units Per Order']
                        
                        expanded_orders.append({
                            'Store ID': row['Store ID'],
                            'Store Name': row['Store Name'],
                            'Store Official Name': row['Store Official Name'],
                            'PO No.': row['PO No.'],
                            'Order Date': row['Order Date'],
                            'Delivery Date': row['Delivery Date'],
                            'Internal Reference': internal_ref,
                            'Barcode': product['Barcode'],
                            'Product Name': product['Name'],
                            'Units Per Order': product['Units Per Order'],
                            'Original Order Quantity': row['# of Order'],
                            'Total Units': total_units,
                            'Unit Price': unit_price,
                            'Total Price': total_units * unit_price,
                            'Is Multi Product': False
                        })
                    except Exception as e:
                        error_msg = f"Error processing single product reference {internal_ref}: {e}"
                        errors.append(error_msg)
                else:
                    error_msg = f"No product found for internal reference: {internal_ref}"
                    errors.append(error_msg)
        
        if not expanded_orders:
            error_msg = "No expanded orders were created - check product variants matching"
            errors.append(error_msg)
            st.error(f"❌ {error_msg}")
            self.expanded_orders = pd.DataFrame()
        else:
            self.expanded_orders = pd.DataFrame(expanded_orders)
            st.success(f"✅ Expanded to {len(self.expanded_orders)} order lines")
            
        return errors
    
    def create_order_line_details(self):
        """Create detailed order line items for Odoo import"""
        if self.expanded_orders is None or self.expanded_orders.empty:
            st.error("❌ No expanded orders available for creating order line details")
            self.order_line_details = pd.DataFrame()
            return
        
        if self.order_summaries is None or self.order_summaries.empty:
            st.error("❌ No order summaries available for creating order line details")
            self.order_line_details = pd.DataFrame()
            return
        
        # Create order reference mapping
        order_ref_mapping = {}
        for _, summary in self.order_summaries.iterrows():
            store_id = summary['Store ID']
            order_ref_mapping[store_id] = summary['Order Reference']
        
        # Create order line details
        line_details = []
        
        for _, row in self.expanded_orders.iterrows():
            store_id = row['Store ID']
            order_ref = order_ref_mapping.get(store_id, f"OATS{store_id:06d}")
            
            # Determine product identifier
            if row['Is Multi Product']:
                # For multi-product references, use barcode
                product_identifier = row['Barcode']
            else:
                # For single product references, use internal reference
                product_identifier = row['Internal Reference']
            
            line_details.append({
                'Order Reference': order_ref,
                'Store ID': store_id,
                'Store Name': row['Store Name'],
                'Internal Reference': row['Internal Reference'],
                'Barcode': row['Barcode'],
                'Product Identifier': product_identifier,
                'Product Name': row['Product Name'],
                'Original Order Quantity': row['Original Order Quantity'],
                'Units Per Order': row['Units Per Order'],
                'Total Units': row['Total Units'],
                'Unit Price': row['Unit Price'],
                'Total Price': row['Total Price'],
                'PO No.': row['
