"""
Invoice Generator Web App
=========================
Web-based interface for generating PDF invoices from CSV data.

Run with: streamlit run web_app.py
"""

import os
import io
import json
import zipfile
import pandas as pd
import streamlit as st
from datetime import datetime

# Import PDF generation from existing module
from invoice_generator import (
    create_invoice_pdf, 
    sanitize_filename, 
    get_column_value,
    clean_column_names,
    HARGA_SATUAN,
    COMPANY_NAME,
    COMPANY_TAGLINE,
    COMPANY_ADDRESS,
    COMPANY_PHONE,
    COMPANY_EMAIL,
    BANK_NAME,
    BANK_ACCOUNT,
    BANK_HOLDER
)

# ==============================================================================
# CONFIGURATION FILE
# ==============================================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

def load_config():
    """Load configuration from JSON file"""
    default_config = {
        "harga_satuan": HARGA_SATUAN,
        "company_name": COMPANY_NAME,
        "company_tagline": COMPANY_TAGLINE,
        "company_address": COMPANY_ADDRESS,
        "company_phone": COMPANY_PHONE,
        "company_email": COMPANY_EMAIL,
        "bank_name": BANK_NAME,
        "bank_account": BANK_ACCOUNT,
        "bank_holder": BANK_HOLDER
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved_config = json.load(f)
                # Merge with defaults (in case new fields are added)
                default_config.update(saved_config)
        except Exception as e:
            st.warning(f"⚠️ Gagal memuat config: {e}")
    
    return default_config

def save_config(config):
    """Save configuration to JSON file"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        st.error(f"❌ Gagal menyimpan config: {e}")
        return False

# Load configuration at startup
config = load_config()

# ==============================================================================
# PAGE CONFIG
# ==============================================================================
st.set_page_config(
    page_title="Invoice Generator",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# CUSTOM CSS
# ==============================================================================
st.markdown("""
<style>
    /* Main header */
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #2C3E50;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #7f8c8d;
        text-align: center;
        margin-bottom: 2rem;
    }
    
    /* Stats cards */
    .stat-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 1rem;
        color: white;
        text-align: center;
    }
    .stat-number {
        font-size: 2rem;
        font-weight: 700;
    }
    .stat-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    
    /* Success message */
    .success-box {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        padding: 1rem 1.5rem;
        border-radius: 0.5rem;
        color: white;
        margin: 1rem 0;
    }
    
    /* File uploader styling */
    .uploadedFile {
        border: 2px dashed #3498db !important;
        border-radius: 1rem !important;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# SIDEBAR - CONFIGURATION
# ==============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Konfigurasi")
    st.markdown("---")
    
    # Price configuration
    st.markdown("### 💰 Harga")
    harga_satuan = st.number_input(
        "Harga Satuan (Rp)",
        min_value=0,
        value=config["harga_satuan"],
        step=10000,
        format="%d"
    )
    
    st.markdown("---")
    
    # Company info (collapsible)
    with st.expander("🏢 Info Perusahaan", expanded=True):
        company_name = st.text_input("Nama Perusahaan", value=config["company_name"])
        company_tagline = st.text_input("Tagline", value=config["company_tagline"])
        company_address = st.text_area("Alamat", value=config["company_address"], height=80)
        company_phone = st.text_input("Telepon", value=config["company_phone"])
        company_email = st.text_input("Email", value=config["company_email"])
    
    # Bank info (collapsible)
    with st.expander("🏦 Info Bank", expanded=True):
        bank_name = st.text_input("Nama Bank", value=config["bank_name"])
        bank_account = st.text_input("No. Rekening", value=config["bank_account"])
        bank_holder = st.text_input("Atas Nama", value=config["bank_holder"])
    
    st.markdown("---")
    
    # SAVE BUTTON
    if st.button("💾 Simpan Konfigurasi", type="primary", use_container_width=True):
        new_config = {
            "harga_satuan": harga_satuan,
            "company_name": company_name,
            "company_tagline": company_tagline,
            "company_address": company_address,
            "company_phone": company_phone,
            "company_email": company_email,
            "bank_name": bank_name,
            "bank_account": bank_account,
            "bank_holder": bank_holder
        }
        if save_config(new_config):
            st.success("✅ Konfigurasi tersimpan!")
            st.balloons()
    
    # Show config file location
    st.caption(f"📁 Config: `config.json`")

# ==============================================================================
# MAIN CONTENT
# ==============================================================================

# Header
st.markdown('<h1 class="main-header">🧾 Invoice Generator</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload CSV, pilih data, dan generate invoice PDF secara otomatis</p>', unsafe_allow_html=True)

# File uploader
st.markdown("### 📤 Upload File CSV")
uploaded_file = st.file_uploader(
    "Drag and drop atau klik untuk upload",
    type=['csv'],
    help="Upload file CSV berisi data order"
)

if uploaded_file is not None:
    # Read CSV
    try:
        # Try different encodings
        try:
            df = pd.read_csv(uploaded_file, encoding='utf-8')
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding='latin-1')
        
        # Clean column names
        df = clean_column_names(df)
        
        # Add selection column
        if 'Pilih' not in df.columns:
            df.insert(0, 'Pilih', True)
        
        st.markdown("---")
        
        # Stats
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📊 Total Data", f"{len(df)} baris")
        with col2:
            # Count estimated total
            qty_col = None
            for col in ['Jumlah (QTY)', 'Jumlah', 'QTY', 'Qty']:
                if col in df.columns:
                    qty_col = col
                    break
            if qty_col:
                total_qty = df[qty_col].fillna(1).astype(int).sum()
            else:
                total_qty = len(df)
            st.metric("📦 Total Qty", f"{total_qty} pcs")
        with col3:
            total_value = total_qty * harga_satuan
            st.metric("💰 Est. Total", f"Rp {total_value:,.0f}".replace(',', '.'))
        
        st.markdown("---")
        
        # Data editor with selection
        st.markdown("### 📋 Data Order")
        st.markdown("✅ **Centang kolom 'Pilih' untuk memilih data yang ingin di-generate**")
        
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Pilih": st.column_config.CheckboxColumn(
                    "Pilih",
                    help="Centang untuk generate invoice",
                    default=True,
                ),
            },
            disabled=[col for col in df.columns if col != 'Pilih'],
            num_rows="fixed"
        )
        
        # Filter selected rows
        selected_df = edited_df[edited_df['Pilih'] == True].drop(columns=['Pilih'])
        
        st.markdown(f"**{len(selected_df)}** dari **{len(df)}** data dipilih")
        
        st.markdown("---")
        
        # Generate button
        st.markdown("### 🚀 Generate Invoice")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            if st.button("🧾 Generate Invoice PDF", type="primary", use_container_width=True, disabled=len(selected_df) == 0):
                if len(selected_df) == 0:
                    st.error("Pilih minimal 1 data untuk generate invoice!")
                else:
                    # Create temp directory
                    import tempfile
                    temp_dir = tempfile.mkdtemp()
                    
                    # Progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    generated_files = []
                    errors = []
                    
                    for idx, (index, row) in enumerate(selected_df.iterrows()):
                        try:
                            # Get order ID and name for filename
                            order_id = get_column_value(row, ['ORDER-ID', 'Order ID', 'ORDER ID', 'order-id', 'OrderID'], f'ORDER{idx+1}')
                            nama = get_column_value(row, ['Nama Lengkap', 'Nama', 'Name', 'NAMA LENGKAP'], 'Unknown')
                            
                            # Generate filename
                            safe_order_id = sanitize_filename(order_id)
                            safe_nama = sanitize_filename(nama)
                            filename = f"Invoice_{safe_order_id}_{safe_nama}.pdf"
                            output_path = os.path.join(temp_dir, filename)
                            
                            # Build config dict with current sidebar values
                            pdf_config = {
                                'harga_satuan': harga_satuan,
                                'company_name': company_name,
                                'company_tagline': company_tagline,
                                'company_address': company_address,
                                'company_phone': company_phone,
                                'company_email': company_email,
                                'bank_name': bank_name,
                                'bank_account': bank_account,
                                'bank_holder': bank_holder
                            }
                            
                            # Generate PDF with config
                            create_invoice_pdf(row, idx, output_path, None, config=pdf_config)
                            generated_files.append((filename, output_path))
                            
                            status_text.text(f"Generating: {filename}")
                            
                        except Exception as e:
                            errors.append(f"{order_id}: {str(e)}")
                        
                        # Update progress
                        progress_bar.progress((idx + 1) / len(selected_df))
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Store in session state
                    st.session_state['generated_files'] = generated_files
                    st.session_state['generation_errors'] = errors
                    st.session_state['temp_dir'] = temp_dir
                    
                    st.rerun()
        
        # Show results if available
        if 'generated_files' in st.session_state and st.session_state['generated_files']:
            generated_files = st.session_state['generated_files']
            errors = st.session_state.get('generation_errors', [])
            
            st.markdown(f"""
            <div class="success-box">
                ✅ <strong>Berhasil generate {len(generated_files)} invoice!</strong>
            </div>
            """, unsafe_allow_html=True)
            
            if errors:
                st.warning(f"⚠️ {len(errors)} invoice gagal di-generate")
                with st.expander("Lihat error"):
                    for err in errors:
                        st.text(err)
            
            # Create ZIP file
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for filename, filepath in generated_files:
                    zip_file.write(filepath, filename)
            zip_buffer.seek(0)
            
            # Download buttons
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📦 Download Semua (ZIP)",
                    data=zip_buffer,
                    file_name=f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            
            with col2:
                # Expander for individual downloads
                with st.expander("📄 Download Individual"):
                    for filename, filepath in generated_files:
                        with open(filepath, 'rb') as f:
                            st.download_button(
                                label=f"📄 {filename}",
                                data=f.read(),
                                file_name=filename,
                                mime="application/pdf",
                                key=f"dl_{filename}"
                            )
            
            # Clear button
            if st.button("🗑️ Clear Results"):
                if 'generated_files' in st.session_state:
                    del st.session_state['generated_files']
                if 'generation_errors' in st.session_state:
                    del st.session_state['generation_errors']
                st.rerun()
    
    except Exception as e:
        st.error(f"❌ Error membaca file CSV: {str(e)}")
        st.info("Pastikan file CSV memiliki format yang benar.")

else:
    # Empty state
    st.markdown("""
    <div style="text-align: center; padding: 3rem; background: #f8f9fa; border-radius: 1rem; margin: 2rem 0;">
        <p style="font-size: 4rem; margin: 0;">📁</p>
        <p style="font-size: 1.2rem; color: #6c757d; margin: 1rem 0;">Upload file CSV untuk memulai</p>
        <p style="font-size: 0.9rem; color: #adb5bd;">Supported format: .csv</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sample format
    with st.expander("📋 Format CSV yang diharapkan"):
        st.markdown("""
        File CSV harus memiliki kolom-kolom berikut (nama kolom fleksibel):
        
        | Kolom | Contoh Nama |
        |-------|-------------|
        | Order ID | `ORDER-ID`, `Order ID` |
        | Nama | `Nama Lengkap`, `Nama` |
        | Alamat | `Alamat Pengiriman`, `Alamat` |
        | Ukuran | `Ukuran Kaos (size)`, `Ukuran` |
        | Jumlah | `Jumlah (QTY)`, `Qty` |
        | Metode Bayar | `Metode Pembayaran` |
        | Status | `STATUS PEMBAYARAN` |
        """)

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #adb5bd; font-size: 0.8rem;'>"
    "Invoice Generator v1.1 | Built with Streamlit 🎈"
    "</p>",
    unsafe_allow_html=True
)
