#!/usr/bin/env python3
"""
T&T Purchase Order Processor - Simplified Version
Upload purchase_orders.xlsx and convert to Odoo format.
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

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 1rem;
        padding: 0.5rem;
        background-color: #ecf0f1;
        border-radius: 5px;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
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
    """Odoo conversion functionality"""
    
    def __init__(self, purchase_orders: pd.DataFrame, product_variants: pd.DataFrame, store_names: pd.DataFrame):
        self.purchase_orders = purchase_orders
        self.product_variants = product_variants
        self.store_names = store_names
        self.order_summaries = None
        self.order_line_details = None
        
    def match_store_names(self) -> List[str]:
        """Match store names with official names using direct Store ID mapping"""
        errors = []
        
        # Create a mapping from store ID to official name using the Store ID column
        store_mapping = {}
        for _, row in self.store_names.iterrows():
            store_id = row['Store ID']
            official_name = row['Store Official Name']
            store_mapping[store_id] = official_name
        
        # Add official store name to purchase orders
        self.purchase_orders['Store Official Name'] = self.purchase_orders['Store ID'].map(store_mapping)
        
        # Log unmatched stores
        unmatched_stores = self.purchase_orders[self.purchase_orders['Store Official Name'].isna()]['Store ID'].unique()
        if len(unmatched_stores) > 0:
            errors.append(f"Unmatched store IDs: {unmatched_stores}")
        
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
    
    def handle_multi_product_references(self) -> List[str]:
        """Handle internal references that link to multiple products"""
        errors = []
        
        # Find internal references with multiple products
        ref_counts = self.product_variants['Internal Reference'].value_counts()
        multi_product_refs = ref_counts[ref_counts > 1].index.tolist()
        
        # Create expanded purchase orders for multi-product references
        expanded_orders = []
        
        for _, row in self.purchase_orders.iterrows():
            internal_ref = row['Internal Reference']
            
            if internal_ref in multi_product_refs:
                # Get all products for this internal reference
                products = self.product_variants[self.product_variants['Internal Reference'] == internal_ref]
                
                # Calculate units per product (distribute equally)
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
                    
                    # Calculate unit price
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
                        'Total Units': actual_units,
                        'Unit Price': unit_price,
                        'Total Price': actual_units * unit_price,
                        'Is Multi Product': True
                    })
            else:
                # Single product reference - keep as is
                product = self.product_variants[self.product_variants['Internal Reference'] == internal_ref]
                if len(product) > 0:
                    product = product.iloc[0]
                    total_units = row['# of Order'] * product['Units Per Order']
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
                        'Total Price': row['Price'],
                        'Is Multi Product': False
                    })
                else:
                    errors.append(f"No product found for internal reference: {internal_ref}")
        
        self.expanded_orders = pd.DataFrame(expanded_orders)
        return errors
    
    def create_order_line_details(self):
        """Create detailed order line items for Odoo import"""
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
                'PO No.': row['PO No.'],
                'Order Date': row['Order Date'],
                'Delivery Date': row['Delivery Date']
            })
        
        self.order_line_details = pd.DataFrame(line_details)
    
    def process_all(self) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
        """Run the complete conversion process"""
        errors = []
        
        # Match store names
        store_errors = self.match_store_names()
        errors.extend(store_errors)
        
        # Create order summaries
        self.create_order_summaries()
        
        # Handle multi-product references
        ref_errors = self.handle_multi_product_references()
        errors.extend(ref_errors)
        
        # Create order line details
        self.create_order_line_details()
        
        return self.order_summaries, self.order_line_details, errors

def main():
    """Main Streamlit application"""
    
    # Header
    st.markdown('<h1 class="main-header">🛒 T&T Purchase Order Processor</h1>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.info("📋 Upload your files to convert purchase orders to Odoo format")
    
    # File uploads in columns
    col1, col2, col3 = st.columns(3)
    
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
        st.markdown('<h3 class="section-header">🏪 Store Names</h3>', unsafe_allow_html=True)
        
        # File type selection
        store_file_type = st.radio(
            "Select file type:",
            ["Excel (.xlsx)", "CSV (.csv)"],
            key="store_names_type"
        )
        
        if store_file_type == "Excel (.xlsx)":
            store_names_file = st.file_uploader(
                "Upload Store Names file",
                type=['xlsx'],
                key="store_names"
            )
        else:
            store_names_file = st.file_uploader(
                "Upload Store Names file",
                type=['csv'],
                key="store_names"
            )
        
        store_names = None
        if store_names_file:
            try:
                if store_file_type == "Excel (.xlsx)":
                    store_names = read_excel_file(store_names_file)
                else:
                    store_names = read_csv_file(store_names_file)
                
                if not store_names.empty:
                    st.success(f"✅ Loaded {len(store_names)} stores")
                    st.dataframe(store_names.head(3), use_container_width=True)
                else:
                    st.error("❌ File is empty")
            except Exception as e:
                st.error(f"❌ Error loading file: {e}")
    
    with col3:
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
    
    # Processing section
    st.markdown("---")
    st.markdown('<h2 class="section-header">🔄 Process & Convert</h2>', unsafe_allow_html=True)
    
    # Check if all files are uploaded
    if product_variants is not None and store_names is not None and purchase_orders is not None:
        if st.button("🚀 Convert to Odoo Format", type="primary", use_container_width=True):
            with st.spinner("Converting to Odoo format..."):
                try:
                    # Initialize converter
                    converter = OdooConverter(purchase_orders, product_variants, store_names)
                    
                    # Process conversion
                    order_summaries, order_line_details, errors = converter.process_all()
                    
                    # Display results
                    st.success("✅ Conversion completed successfully!")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Stores", len(order_summaries))
                    with col2:
                        st.metric("Total Order Lines", len(order_line_details))
                    with col3:
                        st.metric("Total Value", f"${order_line_details['Total Price'].sum():,.2f}")
                    with col4:
                        st.metric("Average Order Value", f"${order_line_details['Total Price'].mean():,.2f}")
                    
                    # Show order summaries
                    with st.expander("📋 Order Summaries", expanded=True):
                        st.dataframe(order_summaries, use_container_width=True)
                    
                    # Show order line details
                    with st.expander("📊 Order Line Details (First 20 rows)", expanded=False):
                        st.dataframe(order_line_details.head(20), use_container_width=True)
                    
                    # Show errors if any
                    if errors:
                        with st.expander("⚠️ Conversion Warnings", expanded=False):
                            for error in errors[:10]:
                                st.warning(error)
                            if len(errors) > 10:
                                st.info(f"... and {len(errors) - 10} more warnings")
                    
                    # Create Excel file for download
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
                            store_names.to_excel(writer, sheet_name='Store Names', index=False)
                        
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
                    - **Store Names**: Reference store data
                    """)
                
                except Exception as e:
                    st.error(f"❌ Error during conversion: {e}")
                    st.error("Please check your file formats and try again.")
    
    else:
        st.info("📋 Please upload all three files to proceed with conversion")
        
        missing_files = []
        if product_variants is None:
            missing_files.append("Product Variants")
        if store_names is None:
            missing_files.append("Store Names")
        if purchase_orders is None:
            missing_files.append("Purchase Orders")
        
        if missing_files:
            st.warning(f"⚠️ Missing files: {', '.join(missing_files)}")
    
    # Help section
    with st.sidebar:
        st.markdown("## ℹ️ Help & Instructions")
        st.markdown("""
        **How to use this tool:**
        
        1. **Upload Files**: Upload all three required files
           - Product Variants (Excel/CSV)
           - Store Names (Excel/CSV)
           - Purchase Orders (Excel/CSV)
        
        2. **Convert**: Click "Convert to Odoo Format" button
        
        3. **Download**: Download the "Odoo_Import_Ready.xlsx" file
        
        **Required File Formats:**
        - **Product Variants**: Must contain columns like 'Internal Reference', 'Barcode', 'Name', 'Units Per Order'
        - **Store Names**: Must contain 'Store ID' and 'Store Official Name'
        - **Purchase Orders**: Must contain 'Store ID', 'Store Name', 'PO No.', 'Order Date', 'Delivery Date', 'Internal Reference', '# of Order', 'Price'
        
        **Features:**
        - Automatic product mapping
        - Multi-product reference handling
        - Comprehensive error reporting
        - Odoo-compatible output format
        """)

if __name__ == "__main__":
    main() 
