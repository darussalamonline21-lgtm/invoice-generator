"""
Invoice Generator - Automated PDF Invoice Creation from CSV Data
================================================================
Author: AI Assistant
Purpose: Generate professional PDF invoices from order data in CSV format

Dependencies:
    pip install pandas reportlab

Usage:
    1. Place your CSV file named "FORM ORDER PO KAOS - Form Responses 1.csv" in the same directory
    2. (Optional) Place logo.png in the same directory
    3. Run: python invoice_generator.py
    4. Invoices will be generated in the 'output_invoices' folder
"""

import os
import re
import pandas as pd
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ==============================================================================
# KONFIGURASI - Ubah nilai-nilai ini sesuai kebutuhan
# ==============================================================================

# Harga satuan per item (dalam Rupiah)
HARGA_SATUAN = 100000

# Informasi Perusahaan
COMPANY_NAME = "TOKO KAOS KEREN"
COMPANY_TAGLINE = "Quality T-Shirts for Everyone"
COMPANY_ADDRESS = "Jl. Contoh No. 123, Jakarta, Indonesia"
COMPANY_PHONE = "+62 812-3456-7890"
COMPANY_EMAIL = "order@tokokaoskeren.com"
COMPANY_WEBSITE = "www.tokokaoskeren.com"

# Bank Information for Payment
BANK_NAME = "Bank BCA"
BANK_ACCOUNT = "1234567890"
BANK_HOLDER = "PT TOKO KAOS KEREN"

# File paths
CSV_FILENAME = "FORM ORDER PO KAOS - Form Responses 1.csv"
LOGO_FILENAME = "logo.png"
OUTPUT_FOLDER = "output_invoices"

# Colors (RGB values from 0-1)
PRIMARY_COLOR = colors.HexColor("#2C3E50")  # Dark blue
SECONDARY_COLOR = colors.HexColor("#3498DB")  # Light blue
ACCENT_COLOR = colors.HexColor("#E74C3C")  # Red for important info
SUCCESS_COLOR = colors.HexColor("#27AE60")  # Green for paid status

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def clean_column_names(df):
    """Bersihkan nama kolom: hapus spasi berlebih, lowercase untuk konsistensi"""
    df.columns = df.columns.str.strip()
    return df

def sanitize_filename(name):
    """Hapus karakter tidak valid dari nama file"""
    # Ganti karakter tidak valid dengan underscore
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', str(name))
    # Ganti spasi dengan underscore
    sanitized = sanitized.replace(' ', '_')
    # Hapus underscore berlebih
    sanitized = re.sub(r'_+', '_', sanitized)
    return sanitized.strip('_')

def format_currency(amount):
    """Format angka ke format Rupiah"""
    return f"Rp {amount:,.0f}".replace(',', '.')

def wrap_text(text, max_chars=40):
    """Wrap text yang terlalu panjang"""
    if pd.isna(text) or text == '':
        return '-'
    text = str(text)
    if len(text) <= max_chars:
        return text
    
    words = text.split()
    lines = []
    current_line = ""
    
    for word in words:
        if len(current_line) + len(word) + 1 <= max_chars:
            current_line += (" " + word if current_line else word)
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    
    if current_line:
        lines.append(current_line)
    
    return '\n'.join(lines)

def get_column_value(row, possible_names, default='-'):
    """Ambil nilai kolom dengan berbagai kemungkinan nama"""
    for name in possible_names:
        if name in row.index and pd.notna(row[name]) and str(row[name]).strip():
            return str(row[name]).strip()
    return default

# ==============================================================================
# PDF GENERATION
# ==============================================================================

def create_invoice_pdf(row, row_index, output_path, logo_path=None):
    """Generate PDF invoice untuk satu order"""
    
    # Extract data dengan handling berbagai nama kolom
    order_id = get_column_value(row, ['ORDER-ID', 'Order ID', 'ORDER ID', 'order-id', 'OrderID'])
    nama = get_column_value(row, ['Nama Lengkap', 'Nama', 'Name', 'NAMA LENGKAP'])
    alamat = get_column_value(row, ['Alamat Pengiriman', 'Alamat', 'Address', 'ALAMAT PENGIRIMAN'])
    ukuran = get_column_value(row, ['Ukuran Kaos (size)', 'Ukuran', 'Size', 'UKURAN KAOS', 'Ukuran Kaos'])
    
    # Qty handling - pastikan integer
    qty_raw = get_column_value(row, ['Jumlah (QTY)', 'Jumlah', 'QTY', 'Qty', 'JUMLAH (QTY)'], '1')
    try:
        qty = int(float(qty_raw))
    except (ValueError, TypeError):
        qty = 1
    
    metode_bayar = get_column_value(row, ['Metode Pembayaran', 'Metode', 'Payment Method', 'METODE PEMBAYARAN'])
    status_bayar = get_column_value(row, ['STATUS PEMBAYARAN', 'Status Pembayaran', 'Status', 'Payment Status'])
    
    # Timestamp (jika ada)
    timestamp = get_column_value(row, ['Timestamp', 'Tanggal', 'Date', 'TIMESTAMP'], 
                                  datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    
    # Phone/HP (jika ada)
    phone = get_column_value(row, ['No HP', 'No. HP', 'Phone', 'Telepon', 'NO HP', 'Nomor HP'])
    
    # ==== KALKULASI HARGA ====
    total_harga = HARGA_SATUAN * qty
    
    # Cek apakah DP 50%
    is_dp = 'dp' in metode_bayar.lower() or '50%' in metode_bayar.lower()
    
    if is_dp:
        jumlah_dibayar = total_harga * 0.5
        sisa_tagihan = total_harga - jumlah_dibayar
    else:
        jumlah_dibayar = total_harga if 'lunas' in status_bayar.lower() else 0
        sisa_tagihan = total_harga - jumlah_dibayar
    
    # ==== CREATE PDF ====
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=1.5*cm,
        bottomMargin=2*cm
    )
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=PRIMARY_COLOR,
        spaceAfter=5,
        alignment=TA_CENTER
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.grey,
        alignment=TA_CENTER,
        spaceAfter=20
    )
    
    header_style = ParagraphStyle(
        'Header',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=PRIMARY_COLOR,
        spaceBefore=15,
        spaceAfter=10
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=10,
        leading=14
    )
    
    bold_style = ParagraphStyle(
        'Bold',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold'
    )
    
    # Build content
    content = []
    
    # ==== HEADER SECTION ====
    # Logo and Company Info (side by side)
    header_data = []
    
    # Left side: Logo (if exists)
    if logo_path and os.path.exists(logo_path):
        logo = Image(logo_path, width=3*cm, height=3*cm)
        logo_cell = logo
    else:
        # Placeholder jika tidak ada logo
        logo_cell = Paragraph(f"<font size='18' color='#{PRIMARY_COLOR.hexval()[2:]}'><b>{COMPANY_NAME[:2]}</b></font>", styles['Normal'])
    
    # Right side: Company info
    company_info = f"""
    <font size='16' color='#{PRIMARY_COLOR.hexval()[2:]}'><b>{COMPANY_NAME}</b></font><br/>
    <font size='8' color='gray'>{COMPANY_TAGLINE}</font><br/><br/>
    <font size='9'>{COMPANY_ADDRESS}</font><br/>
    <font size='9'>📞 {COMPANY_PHONE}</font><br/>
    <font size='9'>✉ {COMPANY_EMAIL}</font>
    """
    
    header_table = Table(
        [[logo_cell, Paragraph(company_info, styles['Normal'])]],
        colWidths=[4*cm, 13*cm]
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    content.append(header_table)
    content.append(Spacer(1, 0.5*cm))
    
    # ==== INVOICE TITLE ====
    # Divider line
    divider = Table([['']], colWidths=[17*cm])
    divider.setStyle(TableStyle([
        ('LINEABOVE', (0, 0), (-1, 0), 2, PRIMARY_COLOR),
    ]))
    content.append(divider)
    content.append(Spacer(1, 0.3*cm))
    
    # Invoice title and number
    invoice_header = Table([
        [Paragraph("<font size='20'><b>INVOICE / NOTA</b></font>", 
                   ParagraphStyle('', alignment=TA_LEFT, textColor=PRIMARY_COLOR)),
         Paragraph(f"<font size='12'><b>#{order_id}</b></font>", 
                   ParagraphStyle('', alignment=TA_RIGHT, textColor=SECONDARY_COLOR))]
    ], colWidths=[10*cm, 7*cm])
    invoice_header.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    content.append(invoice_header)
    content.append(Spacer(1, 0.3*cm))
    
    # Date
    content.append(Paragraph(f"<font size='10' color='gray'>Tanggal: {timestamp}</font>", styles['Normal']))
    content.append(Spacer(1, 0.5*cm))
    
    # ==== CUSTOMER INFO ====
    content.append(Paragraph("<b>INFORMASI PELANGGAN</b>", header_style))
    
    customer_data = [
        ['Nama', ':', nama],
        ['Alamat', ':', wrap_text(alamat, 50)],
        ['No. HP', ':', phone],
    ]
    
    customer_table = Table(customer_data, colWidths=[3*cm, 0.5*cm, 13.5*cm])
    customer_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    content.append(customer_table)
    content.append(Spacer(1, 0.5*cm))
    
    # ==== ORDER DETAILS TABLE ====
    content.append(Paragraph("<b>DETAIL PESANAN</b>", header_style))
    
    # Table header
    order_data = [
        ['No', 'Deskripsi', 'Ukuran', 'Qty', 'Harga Satuan', 'Subtotal'],
        ['1', 'Kaos Custom', ukuran, str(qty), format_currency(HARGA_SATUAN), format_currency(total_harga)],
    ]
    
    order_table = Table(order_data, colWidths=[1*cm, 5*cm, 2.5*cm, 1.5*cm, 3.5*cm, 3.5*cm])
    order_table.setStyle(TableStyle([
        # Header style
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY_COLOR),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Body style
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # No column
        ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # Ukuran
        ('ALIGN', (3, 1), (3, -1), 'CENTER'),  # Qty
        ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),  # Prices
        
        # Grid
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        
        # Alternating row colors
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#F8F9FA")),
    ]))
    content.append(order_table)
    content.append(Spacer(1, 0.3*cm))
    
    # ==== PAYMENT SUMMARY ====
    summary_data = [
        ['', '', 'Total Harga:', format_currency(total_harga)],
    ]
    
    if is_dp:
        summary_data.append(['', '', 'DP 50% Dibayar:', format_currency(jumlah_dibayar)])
        summary_data.append(['', '', 'SISA TAGIHAN:', format_currency(sisa_tagihan)])
    else:
        summary_data.append(['', '', 'Jumlah Dibayar:', format_currency(jumlah_dibayar)])
        if sisa_tagihan > 0:
            summary_data.append(['', '', 'SISA TAGIHAN:', format_currency(sisa_tagihan)])
    
    summary_table = Table(summary_data, colWidths=[5*cm, 5*cm, 3.5*cm, 3.5*cm])
    summary_style = [
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
        ('FONTNAME', (2, 0), (3, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]
    
    # Highlight sisa tagihan
    if len(summary_data) > 1:
        summary_style.append(('TEXTCOLOR', (2, -1), (3, -1), ACCENT_COLOR))
        summary_style.append(('FONTSIZE', (2, -1), (3, -1), 12))
    
    summary_table.setStyle(TableStyle(summary_style))
    content.append(summary_table)
    content.append(Spacer(1, 0.5*cm))
    
    # ==== PAYMENT STATUS ====
    status_color = SUCCESS_COLOR if 'lunas' in status_bayar.lower() else ACCENT_COLOR
    status_text = f"<font size='12' color='#{status_color.hexval()[2:]}'><b>STATUS: {status_bayar.upper()}</b></font>"
    
    status_para = Paragraph(status_text, ParagraphStyle('', alignment=TA_CENTER))
    status_table = Table([[status_para]], colWidths=[17*cm])
    status_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8F9FA")),
        ('BOX', (0, 0), (-1, -1), 1, status_color),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    content.append(status_table)
    content.append(Spacer(1, 0.5*cm))
    
    # ==== PAYMENT METHOD ====
    content.append(Paragraph("<b>METODE PEMBAYARAN</b>", header_style))
    content.append(Paragraph(f"<font size='10'>{metode_bayar}</font>", styles['Normal']))
    content.append(Spacer(1, 0.3*cm))
    
    # Bank info (if there's remaining payment)
    if sisa_tagihan > 0:
        bank_info = f"""
        <font size='10'><b>Transfer ke:</b></font><br/>
        <font size='10'>Bank: {BANK_NAME}</font><br/>
        <font size='10'>No. Rekening: {BANK_ACCOUNT}</font><br/>
        <font size='10'>Atas Nama: {BANK_HOLDER}</font>
        """
        
        bank_table = Table([[Paragraph(bank_info, styles['Normal'])]], colWidths=[17*cm])
        bank_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FFF3CD")),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#FFC107")),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        content.append(bank_table)
    
    content.append(Spacer(1, 1*cm))
    
    # ==== FOOTER ====
    footer_divider = Table([['']], colWidths=[17*cm])
    footer_divider.setStyle(TableStyle([
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.grey),
    ]))
    content.append(footer_divider)
    content.append(Spacer(1, 0.3*cm))
    
    footer_text = f"""
    <font size='8' color='gray'>
    Terima kasih atas pesanan Anda! | {COMPANY_WEBSITE}<br/>
    Invoice ini sah sebagai bukti pemesanan. Harap simpan untuk referensi.
    </font>
    """
    content.append(Paragraph(footer_text, ParagraphStyle('', alignment=TA_CENTER)))
    
    # Build PDF
    doc.build(content)
    return True

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    """Main function to process CSV and generate invoices"""
    
    print("=" * 60)
    print("       INVOICE GENERATOR - PDF dari Data CSV")
    print("=" * 60)
    print()
    
    # Get script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, CSV_FILENAME)
    logo_path = os.path.join(script_dir, LOGO_FILENAME)
    output_dir = os.path.join(script_dir, OUTPUT_FOLDER)
    
    # Check CSV file
    if not os.path.exists(csv_path):
        print(f"❌ ERROR: File CSV tidak ditemukan!")
        print(f"   Expected: {csv_path}")
        print()
        print("   Pastikan file CSV bernama:")
        print(f"   '{CSV_FILENAME}'")
        print("   Berada di folder yang sama dengan script ini.")
        return
    
    # Create output folder
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"📁 Folder output dibuat: {output_dir}")
    
    # Check logo
    if os.path.exists(logo_path):
        print(f"✅ Logo ditemukan: {logo_path}")
    else:
        print(f"⚠️  Logo tidak ditemukan, akan menggunakan placeholder")
        logo_path = None
    
    print()
    
    # Read CSV
    try:
        # Try different encodings
        try:
            df = pd.read_csv(csv_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_path, encoding='latin-1')
        
        # Clean column names
        df = clean_column_names(df)
        
        print(f"📊 Data berhasil dibaca: {len(df)} baris")
        print(f"   Kolom: {list(df.columns)}")
        print()
        
    except Exception as e:
        print(f"❌ ERROR membaca CSV: {e}")
        return
    
    # Process each row
    success_count = 0
    error_count = 0
    
    print("🔄 Memproses invoice...")
    print("-" * 60)
    
    for index, row in df.iterrows():
        try:
            # Get order ID and name for filename
            order_id = get_column_value(row, ['ORDER-ID', 'Order ID', 'ORDER ID', 'order-id', 'OrderID'], f'ORDER{index+1}')
            nama = get_column_value(row, ['Nama Lengkap', 'Nama', 'Name', 'NAMA LENGKAP'], 'Unknown')
            
            # Generate filename
            safe_order_id = sanitize_filename(order_id)
            safe_nama = sanitize_filename(nama)
            filename = f"Invoice_{safe_order_id}_{safe_nama}.pdf"
            output_path = os.path.join(output_dir, filename)
            
            # Generate PDF
            create_invoice_pdf(row, index, output_path, logo_path)
            
            success_count += 1
            print(f"   ✅ [{index+1}/{len(df)}] {filename}")
            
        except Exception as e:
            error_count += 1
            print(f"   ❌ [{index+1}/{len(df)}] Error: {e}")
    
    # Summary
    print()
    print("-" * 60)
    print("📋 RINGKASAN:")
    print(f"   ✅ Berhasil: {success_count} invoice")
    print(f"   ❌ Gagal: {error_count} invoice")
    print(f"   📂 Output: {output_dir}")
    print()
    print("=" * 60)
    print("                    SELESAI!")
    print("=" * 60)

if __name__ == "__main__":
    main()
