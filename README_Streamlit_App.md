# T&T Purchase Order Processor - Streamlit Application

A comprehensive online application that combines PDF extraction and Odoo conversion into a single, user-friendly interface.

## 🚀 Features

- **PDF Extraction**: Extract purchase order data from T&T PDF files
- **Multi-file Processing**: Process multiple PDF files simultaneously
- **Reference Data Management**: Upload and manage Product Variants and Store Names
- **Odoo Conversion**: Convert extracted data to Odoo-compatible format
- **Multi-product Handling**: Automatically handle internal references with multiple products
- **Error Reporting**: Comprehensive error handling and validation
- **Download Ready**: Generate Excel files ready for Odoo import

## 📋 Requirements

### System Requirements
- Python 3.8 or higher
- Internet connection for Streamlit hosting

### Required Files
1. **Product Variant Excel file** - Contains product information with columns:
   - Internal Reference
   - Lark ID
   - Name
   - Variant Values
   - Quantity On Hand
   - Barcode
   - Units Per Order

2. **T&T Store Names Excel file** - Contains store information with columns:
   - Store Official Name
   - Store ID

3. **T&T Purchase Order PDF files** - PDF files containing purchase order data

## 🛠️ Installation

### Option 1: Local Installation

1. **Clone or download the application files**
2. **Install dependencies**:
   ```bash
   pip install -r requirements_streamlit_app.txt
   ```

3. **Run the application**:
   ```bash
   streamlit run streamlit_app.py
   ```

4. **Access the application**:
   - Open your web browser
   - Go to `http://localhost:8501`

### Option 2: Streamlit Cloud Deployment

1. **Create a GitHub repository** with the following files:
   - `streamlit_app.py`
   - `requirements_streamlit_app.txt`
   - `README_Streamlit_App.md`

2. **Deploy to Streamlit Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Connect your GitHub repository
   - Deploy the application

3. **Access your online application**:
   - Use the provided Streamlit Cloud URL

## 📖 How to Use

### Step 1: Upload Reference Data
1. Upload your **Product Variant Excel file**
2. Upload your **T&T Store Names Excel file**
3. Click "Next Step" when both files are loaded

### Step 2: Upload PDF Files
1. Upload one or more **T&T Purchase Order PDF files**
2. Click "Process PDF Files" to extract data
3. Review the extracted data preview
4. Click "Next Step" to proceed

### Step 3: Process & Convert
1. Click "Start Conversion" to convert to Odoo format
2. Review the conversion results and metrics
3. Check any warnings or errors
4. Click "Next Step" to download

### Step 4: Download Results
1. Click "Download Odoo Import Ready File"
2. The Excel file contains:
   - **Order Summaries**: Summary of orders by store
   - **Order Line Details**: Detailed product lines for Odoo import
   - **Original Purchase Orders**: Raw extracted data
   - **Product Variants**: Reference product data
   - **Store Names**: Reference store data

## 🔧 Configuration

### Customizing Order Reference Prefix
To change the order reference prefix (default: "OATS"), modify line 280 in `streamlit_app.py`:
```python
order_ref = f"YOUR_PREFIX{order_ref_counter:06d}"
```

### Adjusting Validation Rules
Modify validation functions in the `PDFExtractor` class:
- `validate_date()`: Date format validation
- `validate_numeric()`: Numeric value validation

## 🐛 Troubleshooting

### Common Issues

1. **PDF Processing Errors**:
   - Ensure PDF files are not password-protected
   - Check that PDFs contain text (not scanned images)
   - Verify PDF format matches T&T purchase order structure

2. **Reference Data Mismatches**:
   - Ensure Store IDs match between PDFs and Store Names file
   - Verify Internal References exist in Product Variants file
   - Check for typos in reference data

3. **Memory Issues**:
   - Process fewer PDF files at once
   - Close other applications to free memory
   - Use smaller reference files if possible

### Error Messages

- **"No text extracted from page"**: PDF may be scanned or corrupted
- **"Unmatched store IDs"**: Store IDs in PDFs don't match reference data
- **"No product found for internal reference"**: Product not in reference file
- **"Invalid date format"**: Date format doesn't match MM/DD/YYYY

## 📊 Output Format

### Order Summaries Sheet
| Column | Description |
|--------|-------------|
| Order Reference | Unique order identifier (OATS000006, etc.) |
| Customer Official Name | Official store name from reference data |
| Store ID | Store identifier |
| Store Name | Store name from PDF |
| Order Date | Earliest order date for the store |
| Delivery Date | Earliest delivery date for the store |
| PO Numbers | All PO numbers for the store (comma-separated) |
| Total PO Count | Number of unique PO numbers |

### Order Line Details Sheet
| Column | Description |
|--------|-------------|
| Order Reference | Links to Order Summaries |
| Store ID | Store identifier |
| Store Name | Store name |
| Internal Reference | Product internal reference |
| Barcode | Product barcode |
| Product Identifier | Barcode (multi-product) or Internal Reference (single) |
| Product Name | Product name from reference data |
| Original Order Quantity | Quantity from PDF |
| Units Per Order | Units per order from reference data |
| Total Units | Calculated total units |
| Unit Price | Calculated unit price |
| Total Price | Calculated total price |
| PO No. | Purchase order number |
| Order Date | Order date |
| Delivery Date | Delivery date |

## 🔒 Security & Privacy

- **Data Processing**: All processing happens in your browser/session
- **File Storage**: Files are not permanently stored on the server
- **Session Data**: Data is cleared when you close the browser or start over
- **No External Sharing**: Your data is not shared with third parties

## 📞 Support

For technical support or feature requests:
1. Check the troubleshooting section above
2. Review error messages in the application
3. Ensure all required files are properly formatted
4. Contact your system administrator for deployment issues

## 🔄 Updates

To update the application:
1. Download the latest version
2. Replace existing files
3. Restart the Streamlit application
4. Clear browser cache if needed

## 📝 License

This application is provided as-is for internal use. Please ensure compliance with your organization's data handling policies. 