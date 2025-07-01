#!/usr/bin/env python3
"""
T&T Purchase Order Processor - Streamlit Cloud Compatible Version
Combines PDF extraction and Odoo conversion into a single online application.
Optimized for Streamlit Cloud deployment.
"""

import streamlit as st
import pandas as pd
import re
import numpy as np
from datetime import datetime
import logging
from typing import Dict, List, Tuple, Optional
from io import BytesIO, StringIO
import tempfile
import os
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
    st.error("Failed to install openpyxl. Please check your requirements file.")

if not install_package("xlrd"):
    st.error("Failed to install xlrd. Please check your requirements file.")

# Try to import pdfplumber, fallback to alternative if not available
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

# Try to import Excel reading libraries
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.error("openpyxl is not available. Please ensure it's in your requirements.txt file.")

try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False
    st.error("xlrd is not available. Please ensure it's in your requirements.txt file.")

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
    .step-header {
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
        
        # Try with xlrd (for .xls files)
        if XLRD_AVAILABLE:
            try:
                return pd.read_excel(file, engine='xlrd')
            except Exception as e:
                logger.warning(f"xlrd failed: {e}")
        
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
        elif file_name.endswith('.xls'):
            if XLRD_AVAILABLE:
                return pd.read_excel(file, engine='xlrd')
            else:
                raise Exception("xlrd is required for .xls files but not available")
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

class PDFExtractor:
    """PDF extraction functionality with fallback methods"""
    
    @staticmethod
    def validate_date(date_str: str) -> bool:
        """Validate date format MM/DD/YYYY"""
        try:
            datetime.strptime(date_str, '%m/%d/%Y')
            return True
        except ValueError:
            return False

    @staticmethod
    def validate_numeric(value: str, min_val: float = 0, max_val: float = float('inf')) -> bool:
        """Validate numeric value within range"""
        try:
            num_val = float(value)
            return min_val <= num_val <= max_val
        except ValueError:
            return False

    @staticmethod
    def extract_po_data(pdf_file) -> Tuple[List[Dict], List[str]]:
        """Extract purchase order data from PDF with improved error handling and validation."""
        data = []
        current_po = {}
        errors = []
        
        try:
            if PDFPLUMBER_AVAILABLE:
                # Use pdfplumber if available
                with pdfplumber.open(pdf_file) as pdf:
                    for page_num, page in enumerate(pdf.pages, 1):
                        try:
                            text = page.extract_text()
                            if not text:
                                logger.warning(f"No text extracted from page {page_num}")
                                continue
                                
                            lines = text.split('\n')
                            data, errors = PDFExtractor._process_lines(lines, current_po, data, errors, page_num)
                                
                        except Exception as e:
                            errors.append(f"Error processing page {page_num}: {e}")
                            continue
            else:
                # Fallback: Try to extract text using alternative method
                st.error("❌ PDF processing is not available in this environment. Please use the standalone converter for PDF processing.")
                return [], ["PDF processing not available in Streamlit Cloud environment"]
                        
        except Exception as e:
            raise Exception(f"Failed to open or process PDF: {e}")
        
        return data, errors

    @staticmethod
    def _process_lines(lines: List[str], current_po: Dict, data: List[Dict], errors: List[str], page_num: int) -> Tuple[List[Dict], List[str]]:
        """Process lines from PDF text"""
        for line_num, line in enumerate(lines, 1):
            try:
                # Extract PO Number with improved pattern
                if po_match := re.search(r'PO\s*No\.?\s*:?\s*(\d+)', line, re.IGNORECASE):
                    current_po['PO No.'] = po_match.group(1)
                
                # Extract Store Name and Store ID with more flexible pattern
                store_patterns = [
                    r'Store\s*:\s*(.*?)\s*-\s*(\d{3})\b',
                    r'Store\s*:\s*(.*?)\s*\((\d{3})\)',
                    r'Store\s*ID\s*:\s*(\d{3})\s*-\s*(.*?)\b'
                ]
                
                for pattern in store_patterns:
                    if store_match := re.search(pattern, line, re.IGNORECASE):
                        if pattern == r'Store\s*ID\s*:\s*(\d{3})\s*-\s*(.*?)\b':
                            current_po['Store ID'] = store_match.group(1)
                            current_po['Store Name'] = store_match.group(2).strip()
                        else:
                            current_po['Store Name'] = store_match.group(1).strip()
                            current_po['Store ID'] = store_match.group(2)
                        break
                
                # Extract Dates with validation
                date_patterns = [
                    (r'Order\s*Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', 'Order Date'),
                    (r'Delivery\s*Date\s*\(on\s*or\s*before\)\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', 'Delivery Date'),
                    (r'Delivery\s*Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', 'Delivery Date')
                ]
                
                for pattern, field_name in date_patterns:
                    if date_match := re.search(pattern, line, re.IGNORECASE):
                        date_str = date_match.group(1)
                        if PDFExtractor.validate_date(date_str):
                            current_po[field_name] = date_str
                        else:
                            errors.append(f"Invalid date format: {date_str} on line {line_num}")
                        break
                
                # Parse item lines with improved validation
                if 'PO No.' in current_po and re.match(r'^\d{6}\b', line.strip()):
                    parts = line.strip().split()
                    if len(parts) < 4:
                        continue
                        
                    # Find numeric values (quantities and prices)
                    numeric_values = [part for part in parts if re.match(r'^\d+\.\d{2}$', part)]
                    
                    if len(numeric_values) >= 2:
                        try:
                            ordered_qty = float(numeric_values[-2])
                            price = float(numeric_values[-1])
                            
                            # Validate quantities and prices
                            if not PDFExtractor.validate_numeric(str(ordered_qty), 0, 1000000):
                                errors.append(f"Invalid quantity: {ordered_qty} on line {line_num}")
                                continue
                                
                            if not PDFExtractor.validate_numeric(str(price), 0, 1000000):
                                errors.append(f"Invalid price: {price} on line {line_num}")
                                continue
                            
                            data.append({
                                'PO No.': current_po['PO No.'],
                                'Store ID': current_po.get('Store ID', ''),
                                'Store Name': current_po.get('Store Name', ''),
                                'Order Date': current_po.get('Order Date', ''),
                                'Delivery Date': current_po.get('Delivery Date', ''),
                                'Item#': parts[0],
                                'Ordered Qty': ordered_qty,
                                'Price': price
                            })
                        except (ValueError, IndexError) as e:
                            errors.append(f"Error parsing numeric values on line {line_num}: {e}")
                            
            except Exception as e:
                errors.append(f"Error processing line {line_num}: {e}")
                continue
        
        return data, errors

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
                    unit_price = row['Price'] / total_units
                    
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
                    unit_price = row['Price'] / total_units
                    
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
    
    # Show environment info
    if not PDFPLUMBER_AVAILABLE:
        st.warning("""
        ⚠️ **Streamlit Cloud Environment Detected**
        
        PDF processing is limited in this environment. For full PDF processing capabilities, please:
        1. Use the local version of this application, or
        2. Process PDFs locally and upload the extracted data as Excel files
        """)
    
    # Show Excel reading capabilities
    if not OPENPYXL_AVAILABLE and not XLRD_AVAILABLE:
        st.error("""
        ❌ **Excel Reading Libraries Not Available**
        
        Required libraries for reading Excel files are not installed. Please ensure the following are available:
        - openpyxl (for .xlsx files)
        - xlrd (for .xls files)
        
        **Alternative: You can convert your Excel files to CSV format and upload them instead.**
        """)
        
        # Show CSV upload option as alternative
        st.info("""
        📄 **CSV Upload Alternative**
        
        If Excel files are not working, you can:
        1. Open your Excel files in a spreadsheet application
        2. Save them as CSV files (File → Save As → CSV)
        3. Upload the CSV files instead
        """)
    
    # Initialize session state
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'purchase_orders' not in st.session_state:
        st.session_state.purchase_orders = None
    if 'product_variants' not in st.session_state:
        st.session_state.product_variants = None
    if 'store_names' not in st.session_state:
        st.session_state.store_names = None
    if 'extraction_errors' not in st.session_state:
        st.session_state.extraction_errors = []
    if 'conversion_errors' not in st.session_state:
        st.session_state.conversion_errors = []
    
    # Sidebar for navigation
    st.sidebar.title("📋 Processing Steps")
    
    # Step indicators
    steps = ["1. Upload Reference Data", "2. Upload Data", "3. Process & Convert", "4. Download Results"]
    
    for i, step_name in enumerate(steps, 1):
        if i == st.session_state.step:
            st.sidebar.markdown(f"**{step_name}** ✅")
        elif i < st.session_state.step:
            st.sidebar.markdown(f"~~{step_name}~~ ✅")
        else:
            st.sidebar.markdown(f"{step_name}")
    
    st.sidebar.markdown("---")
    
    # Step 1: Upload Reference Data
    if st.session_state.step == 1:
        st.markdown('<h2 class="step-header">Step 1: Upload Reference Data</h2>', unsafe_allow_html=True)
        
        st.info("📁 Please upload the required reference files before processing data.")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📦 Product Variants")
            
            # File type selection
            file_type = st.radio(
                "Select file type:",
                ["Excel (.xlsx/.xls)", "CSV (.csv)"],
                key="product_variants_type"
            )
            
            if file_type == "Excel (.xlsx/.xls)":
                product_variants_file = st.file_uploader(
                    "Upload Product Variant Excel file",
                    type=['xlsx', 'xls']
                )
            else:
                product_variants_file = st.file_uploader(
                    "Upload Product Variant CSV file",
                    type=['csv']
                )
            
            if product_variants_file:
                try:
                    if file_type == "Excel (.xlsx/.xls)":
                        df = read_excel_file(product_variants_file)
                    else:
                        df = read_csv_file(product_variants_file)
                    
                    st.session_state.product_variants = df
                    st.success(f"✅ Product Variants loaded: {len(df)} products")
                    st.dataframe(df.head(), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ Error loading Product Variants: {e}")
        
        with col2:
            st.subheader("🏪 T&T Store Names")
            
            # File type selection
            file_type2 = st.radio(
                "Select file type:",
                ["Excel (.xlsx/.xls)", "CSV (.csv)"],
                key="store_names_type"
            )
            
            if file_type2 == "Excel (.xlsx/.xls)":
                store_names_file = st.file_uploader(
                    "Upload T&T Store Names Excel file",
                    type=['xlsx', 'xls']
                )
            else:
                store_names_file = st.file_uploader(
                    "Upload T&T Store Names CSV file",
                    type=['csv']
                )
            
            if store_names_file:
                try:
                    if file_type2 == "Excel (.xlsx/.xls)":
                        df = read_excel_file(store_names_file)
                    else:
                        df = read_csv_file(store_names_file)
                    
                    st.session_state.store_names = df
                    st.success(f"✅ Store Names loaded: {len(df)} stores")
                    st.dataframe(df.head(), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ Error loading Store Names: {e}")
        
        # Navigation buttons
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.session_state.product_variants is not None and st.session_state.store_names is not None:
                if st.button("Next Step →", type="primary"):
                    st.session_state.step = 2
                    st.rerun()
            else:
                st.button("Next Step →", disabled=True)
    
    # Step 2: Upload Data
    elif st.session_state.step == 2:
        st.markdown('<h2 class="step-header">Step 2: Upload Data</h2>', unsafe_allow_html=True)
        
        if not PDFPLUMBER_AVAILABLE:
            st.info("📄 Upload extracted purchase order data as Excel or CSV files (PDF processing not available in this environment).")
            
            # File type selection
            file_type = st.radio(
                "Select file type:",
                ["Excel (.xlsx/.xls)", "CSV (.csv)"],
                key="data_upload_type"
            )
            
            if file_type == "Excel (.xlsx/.xls)":
                uploaded_files = st.file_uploader(
                    "Upload Purchase Order Excel files",
                    type=['xlsx', 'xls'],
                    accept_multiple_files=True,
                    help="Upload Excel files containing extracted purchase order data"
                )
            else:
                uploaded_files = st.file_uploader(
                    "Upload Purchase Order CSV files",
                    type=['csv'],
                    accept_multiple_files=True,
                    help="Upload CSV files containing extracted purchase order data"
                )
            
            if uploaded_files:
                st.success(f"📁 {len(uploaded_files)} file(s) uploaded")
                
                # Process files
                if st.button("Process Files", type="primary"):
                    with st.spinner("Processing files..."):
                        all_data = []
                        all_errors = []
                        
                        for i, uploaded_file in enumerate(uploaded_files):
                            st.info(f"Processing file {i+1}/{len(uploaded_files)}: {uploaded_file.name}")
                            
                            try:
                                if file_type == "Excel (.xlsx/.xls)":
                                    df = read_excel_file(uploaded_file)
                                else:
                                    df = read_csv_file(uploaded_file)
                                
                                all_data.append(df)
                                
                            except Exception as e:
                                all_errors.append(f"{uploaded_file.name}: {e}")
                        
                        if all_data:
                            # Combine all dataframes
                            combined_df = pd.concat(all_data, ignore_index=True)
                            
                            # Clean column names
                            combined_df.columns = combined_df.columns.str.strip()
                            if '# of Order ' in combined_df.columns:
                                combined_df = combined_df.rename(columns={'# of Order ': '# of Order'})
                            
                            # Convert to numeric for proper sorting
                            combined_df['Store ID'] = pd.to_numeric(combined_df['Store ID'], errors='coerce')
                            combined_df['PO No.'] = pd.to_numeric(combined_df['PO No.'], errors='coerce')
                            
                            # Sort by Store ID and PO No.
                            combined_df = combined_df.sort_values(by=['Store ID', 'PO No.'], ascending=[True, True])
                            
                            # Reorder columns
                            combined_df = combined_df[['Store ID', 'Store Name', 'PO No.', 'Order Date', 'Delivery Date',
                                    'Internal Reference', '# of Order', 'Price']]
                            
                            st.session_state.purchase_orders = combined_df
                            st.session_state.extraction_errors = all_errors
                            
                            st.success(f"✅ Successfully loaded {len(combined_df)} order lines from {len(uploaded_files)} file(s)")
                            
                            # Show preview
                            with st.expander("📊 Preview of Loaded Data", expanded=True):
                                st.dataframe(combined_df.head(20), use_container_width=True)
                            
                            # Show errors if any
                            if all_errors:
                                with st.expander("⚠️ Processing Warnings", expanded=False):
                                    for error in all_errors[:10]:
                                        st.warning(error)
                                    if len(all_errors) > 10:
                                        st.info(f"... and {len(all_errors) - 10} more warnings")
                            
                            # Navigation
                            col1, col2, col3 = st.columns([1, 1, 1])
                            with col2:
                                if st.button("Next Step →", type="primary"):
                                    st.session_state.step = 3
                                    st.rerun()
                        else:
                            st.error("❌ No valid data found in the uploaded files.")
        else:
            st.info("📄 Upload T&T Purchase Order PDF files for processing.")
            
            uploaded_files = st.file_uploader(
                "Upload T&T PO PDF files",
                type="pdf",
                accept_multiple_files=True,
                help="Select one or more PDF files containing T&T purchase order data"
            )
            
            if uploaded_files:
                st.success(f"📁 {len(uploaded_files)} PDF file(s) uploaded")
                
                # Show file details
                file_details = []
                for i, file in enumerate(uploaded_files, 1):
                    file_details.append({
                        "File": f"{i}. {file.name}",
                        "Size": f"{file.size / 1024:.1f} KB"
                    })
                
                st.dataframe(pd.DataFrame(file_details), use_container_width=True)
                
                # Process PDFs
                if st.button("Process PDF Files", type="primary"):
                    with st.spinner("Processing PDF files..."):
                        all_data = []
                        all_errors = []
                        
                        for i, uploaded_file in enumerate(uploaded_files):
                            st.info(f"Processing file {i+1}/{len(uploaded_files)}: {uploaded_file.name}")
                            
                            try:
                                data, errors = PDFExtractor.extract_po_data(BytesIO(uploaded_file.getvalue()))
                                all_data.extend(data)
                                all_errors.extend([f"{uploaded_file.name}: {error}" for error in errors])
                                
                            except Exception as e:
                                all_errors.append(f"{uploaded_file.name}: {e}")
                        
                        if all_data:
                            # Create DataFrame and process
                            df = pd.DataFrame(all_data)
                            
                            # Clean column names
                            df.columns = df.columns.str.strip()
                            if '# of Order ' in df.columns:
                                df = df.rename(columns={'# of Order ': '# of Order'})
                            
                            # Convert to numeric for proper sorting
                            df['Store ID'] = pd.to_numeric(df['Store ID'], errors='coerce')
                            df['PO No.'] = pd.to_numeric(df['PO No.'], errors='coerce')
                            
                            # Sort by Store ID and PO No.
                            df = df.sort_values(by=['Store ID', 'PO No.'], ascending=[True, True])
                            
                            # Reorder columns
                            df = df[['Store ID', 'Store Name', 'PO No.', 'Order Date', 'Delivery Date',
                                    'Internal Reference', '# of Order', 'Price']]
                            
                            st.session_state.purchase_orders = df
                            st.session_state.extraction_errors = all_errors
                            
                            st.success(f"✅ Successfully extracted {len(df)} order lines from {len(uploaded_files)} file(s)")
                            
                            # Show preview
                            with st.expander("📊 Preview of Extracted Data", expanded=True):
                                st.dataframe(df.head(20), use_container_width=True)
                            
                            # Show errors if any
                            if all_errors:
                                with st.expander("⚠️ Processing Warnings", expanded=False):
                                    for error in all_errors[:10]:
                                        st.warning(error)
                                    if len(all_errors) > 10:
                                        st.info(f"... and {len(all_errors) - 10} more warnings")
                            
                            # Navigation
                            col1, col2, col3 = st.columns([1, 1, 1])
                            with col2:
                                if st.button("Next Step →", type="primary"):
                                    st.session_state.step = 3
                                    st.rerun()
                        else:
                            st.error("❌ No valid purchase order data found in the uploaded files.")
        
        # Back button
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("← Back"):
                st.session_state.step = 1
                st.rerun()
    
    # Step 3: Process & Convert
    elif st.session_state.step == 3:
        st.markdown('<h2 class="step-header">Step 3: Process & Convert to Odoo Format</h2>', unsafe_allow_html=True)
        
        if st.session_state.purchase_orders is not None:
            st.info("🔄 Converting data to Odoo-compatible format...")
            
            if st.button("Start Conversion", type="primary"):
                with st.spinner("Converting to Odoo format..."):
                    try:
                        # Initialize converter
                        converter = OdooConverter(
                            st.session_state.purchase_orders,
                            st.session_state.product_variants,
                            st.session_state.store_names
                        )
                        
                        # Process conversion
                        order_summaries, order_line_details, errors = converter.process_all()
                        
                        st.session_state.order_summaries = order_summaries
                        st.session_state.order_line_details = order_line_details
                        st.session_state.conversion_errors = errors
                        
                        # Display results
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
                        
                        st.success("✅ Conversion completed successfully!")
                        
                        # Navigation
                        col1, col2, col3 = st.columns([1, 1, 1])
                        with col2:
                            if st.button("Next Step →", type="primary"):
                                st.session_state.step = 4
                                st.rerun()
                    
                    except Exception as e:
                        st.error(f"❌ Error during conversion: {e}")
        
        # Back button
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("← Back"):
                st.session_state.step = 2
                st.rerun()
    
    # Step 4: Download Results
    elif st.session_state.step == 4:
        st.markdown('<h2 class="step-header">Step 4: Download Results</h2>', unsafe_allow_html=True)
        
        if st.session_state.order_summaries is not None and st.session_state.order_line_details is not None:
            st.success("🎉 Processing completed! Download your Odoo-ready file below.")
            
            # Create Excel file
            with st.spinner("Preparing download file..."):
                excel_buffer = BytesIO()
                
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    # Save order summaries
                    st.session_state.order_summaries.to_excel(writer, sheet_name='Order Summaries', index=False)
                    
                    # Save order line details
                    st.session_state.order_line_details.to_excel(writer, sheet_name='Order Line Details', index=False)
                    
                    # Save original data for reference
                    st.session_state.purchase_orders.to_excel(writer, sheet_name='Original Purchase Orders', index=False)
                    st.session_state.product_variants.to_excel(writer, sheet_name='Product Variants', index=False)
                    st.session_state.store_names.to_excel(writer, sheet_name='Store Names', index=False)
                
                excel_buffer.seek(0)
            
            # Download button
            st.download_button(
                label="📥 Download Odoo Import Ready File",
                data=excel_buffer.getvalue(),
                file_name="Odoo_Import_Ready.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.info("📋 The downloaded file contains:")
            st.markdown("""
            - **Order Summaries**: Summary of orders by store
            - **Order Line Details**: Detailed product lines for Odoo import
            - **Original Purchase Orders**: Raw extracted data
            - **Product Variants**: Reference product data
            - **Store Names**: Reference store data
            """)
            
            # Start over button
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                if st.button("🔄 Start Over", type="secondary"):
                    # Reset session state
                    for key in ['step', 'purchase_orders', 'product_variants', 'store_names', 
                               'order_summaries', 'order_line_details', 'extraction_errors', 'conversion_errors']:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.rerun()
        
        # Back button
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("← Back"):
                st.session_state.step = 3
                st.rerun()
    
    # Help section
    with st.sidebar.expander("ℹ️ Help & Instructions"):
        st.markdown("""
        **How to use this tool:**
        
        1. **Upload Reference Data**: Upload Product Variants and Store Names files
        2. **Upload Data**: Upload PDF files or Excel/CSV files with extracted data
        3. **Process & Convert**: Convert to Odoo-compatible format
        4. **Download Results**: Get your Odoo Import Ready file
        
        **Required Files:**
        - Product Variant Excel/CSV file
        - T&T Store Names Excel/CSV file
        - T&T Purchase Order PDF files (or Excel/CSV files with extracted data)
        
        **Features:**
        - Multi-file processing
        - Automatic product mapping
        - Multi-product reference handling
        - Comprehensive error reporting
        - Odoo-compatible output format
        """)

if __name__ == "__main__":
    main() 