import streamlit as st
from supabase import create_client
from datetime import datetime, timezone, timedelta
import pandas as pd
import json
import time
import io
import re

# ==========================================
# 1. ULTRA-PREMIUM UI & CSS INJECTION
# ==========================================
st.set_page_config(page_title="Galaxy ERP", layout="wide", page_icon="🌌", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp { background-color: #f4f6f9; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 5px 5px 0 0; padding: 10px 20px; box-shadow: 0px -2px 5px rgba(0,0,0,0.05); }
    .stTabs [aria-selected="true"] { background-color: #1a237e; color: white !important; border-bottom: none; font-weight: bold; }
    .stButton>button { width: 100%; border-radius: 6px; background-color: #1a237e; color: white; font-weight: 600; border: none; padding: 10px; transition: all 0.3s ease; }
    .stButton>button:hover { background-color: #0d145c; box-shadow: 0 4px 8px rgba(0,0,0,0.2); color: white;}
    [data-testid="stMetric"] { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #1a237e; }
    [data-testid="stForm"] { background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); border: 1px solid #e0e0e0; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <div style="background: linear-gradient(135deg, #1a237e 0%, #3949ab 100%); padding: 25px; border-radius: 10px; color: white; text-align: center; margin-bottom: 25px; margin-top: -40px; box-shadow: 0 4px 15px rgba(26,35,126,0.3);">
        <h1 style="margin:0; font-size: 36px; font-weight: 800; letter-spacing: 2px; color: white;">🌌 GALAXY AUTOMOBILES</h1>
        <p style="margin:5px 0 0 0; font-size: 16px; opacity: 0.9;">Premium Multi-Brand Auto Workshop ERP</p>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 2. DATABASE SETUP
# ==========================================
URL = "https://xthuqvzuvsdbtqaxgrlq.supabase.co"
KEY = "sb_publishable_vniJjmRGyI50rLx_Oyctnw_v6gqkw_a"
supabase = create_client(URL, KEY)
IST = timezone(timedelta(hours=5, minutes=30))

def get_ist(utc_time_str):
    if not utc_time_str: return "N/A", "N/A"
    try:
        dt_utc = datetime.fromisoformat(utc_time_str.replace('Z', '+00:00'))
        dt_ist = dt_utc.astimezone(IST)
        return dt_ist.strftime('%d-%m-%Y'), dt_ist.strftime('%I:%M %p')
    except: return str(utc_time_str)[:10], ""

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Galaxy_Data')
    return output.getvalue()

@st.cache_data(ttl=2)
def fetch_galaxy_data():
    try:
        b = supabase.table("galaxy_billing").select("*").order("created_at", desc=True).execute().data
        return b
    except: return []

bills_data = fetch_galaxy_data()

# ==========================================
# CLASSIC STANDARD INVOICE GENERATOR
# ==========================================
def generate_invoice_html(inv_no, inv_date, veh, parts, labor, gst, total, paid, doc_type="Tax Invoice", items_list=None, shop_gst="05AHYPG3733B2ZG", db_status="", customer_name="", customer_gst="", customer_address="", km_reading=""):
    balance = total - paid
    
    title_text = "TAX INVOICE" if doc_type == "Tax Invoice" else doc_type.upper()
    if db_status == "Cancelled": title_text = "CANCELLED DOCUMENT"
    
    parts_rows = ""
    labor_rows = ""
    p_idx = 1
    l_idx = 1
    
    hsn_html = ""
    
    if items_list:
        # Generate Item Rows
        for item in items_list:
            amt = float(item.get('Qty', 1)) * float(item.get('Rate', 0))
            
            # Extract just the 4-digit code for printing, safely handle empty values
            raw_hsn = str(item.get('HSN', '-')).strip()
            clean_hsn = raw_hsn.split()[0] if raw_hsn and raw_hsn != '-' else '-'
            item_hsn = "9987" if item.get('Type') == 'Labor' else clean_hsn
            
            row_html = f"""
            <tr style='border-bottom: 1px solid #e0e0e0;'>
                <td style='padding: 6px 8px; text-align: center; color: #555;'>{{idx}}</td>
                <td style='padding: 6px 8px;'>{item.get('Description','')}</td>
                <td style='padding: 6px 8px; text-align: center;'>{item_hsn}</td>
                <td style='padding: 6px 8px; text-align: center;'>{item.get('Qty',1)}</td>
                <td style='padding: 6px 8px; text-align: right;'>{float(item.get('Rate',0)):.2f}</td>
                <td style='padding: 6px 8px; text-align: right; font-weight: 600;'>{amt:.2f}</td>
            </tr>"""
            
            if item.get('Type') == "Part":
                parts_rows += row_html.format(idx=p_idx)
                p_idx += 1
            else:
                labor_rows += row_html.format(idx=l_idx)
                l_idx += 1

        # HSN Summary Calculation
        hsn_summary = {}
        raw_labor_total = sum(float(item.get('Qty', 1)) * float(item.get('Rate', 0)) for item in items_list if isinstance(item, dict) and item.get('Type') == 'Labor')
        
        # Determine labor multiplier in case a global labor discount was applied
        labor_multiplier = (labor / raw_labor_total) if raw_labor_total > 0 else 1.0
        
        for item in items_list:
            if not isinstance(item, dict): continue
            
            # Extract just the 4-digit code for printing summary
            raw_hsn = str(item.get('HSN', '-')).strip()
            clean_hsn = raw_hsn.split()[0] if raw_hsn and raw_hsn != '-' else '-'
            hsn = "9987" if item.get('Type') == 'Labor' else clean_hsn
            
            raw_amt = float(item.get('Qty', 1)) * float(item.get('Rate', 0))
            taxable_amt = raw_amt * labor_multiplier if item.get('Type') == 'Labor' else raw_amt
            
            if hsn not in hsn_summary:
                hsn_summary[hsn] = 0.0
            hsn_summary[hsn] += taxable_amt
            
        has_gst = gst > 0
        hsn_rows_html = ""
        
        for hsn, taxable in hsn_summary.items():
            cgst_amt = (taxable * 0.09) if has_gst else 0.0
            sgst_amt = (taxable * 0.09) if has_gst else 0.0
            hsn_total = taxable + cgst_amt + sgst_amt
            
            hsn_rows_html += f"""
            <tr>
                <td class="left">{hsn}</td>
                <td>{taxable:.2f}</td>
                <td>{cgst_amt:.2f}</td>
                <td>{sgst_amt:.2f}</td>
                <td style="font-weight: 600;">{hsn_total:.2f}</td>
            </tr>"""
            
        if hsn_rows_html:
            hsn_html = f"""
            <div style="width: 55%; float: left; margin-top: 5px;">
                <div style="font-size: 10.5px; font-weight: 700; color: #64748b; text-transform: uppercase; margin-bottom: 4px;">HSN / SAC Summary</div>
                <table class="hsn-table">
                    <thead>
                        <tr>
                            <th class="left">HSN/SAC</th>
                            <th>Taxable Value</th>
                            <th>CGST (9%)</th>
                            <th>SGST (9%)</th>
                            <th>Total Amount</th>
                        </tr>
                    </thead>
                    <tbody>
                        {hsn_rows_html}
                    </tbody>
                </table>
            </div>
            """

    table_content = ""
    if parts_rows:
        table_content += """
        <tr style="background-color: #fafdff; border-bottom: 1px solid #d0d7de;">
            <td colspan="6" style="padding: 5px 8px; font-weight: bold; font-size: 11px; color: #0f172a; letter-spacing: 0.5px;">⚙️ SPARE PARTS</td>
        </tr>"""
        table_content += parts_rows
        
    if labor_rows:
        if parts_rows:
            table_content += """<tr style="height: 12px;"><td colspan="6" style="border: none;"></td></tr>"""
        table_content += """
        <tr style="background-color: #fffdfa; border-bottom: 1px solid #d0d7de;">
            <td colspan="6" style="padding: 5px 8px; font-weight: bold; font-size: 11px; color: #0f172a; letter-spacing: 0.5px;">🛠️ LABOR & SERVICES</td>
        </tr>"""
        table_content += labor_rows

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; margin: 0; padding: 10px; background: #fff; font-size: 12px; }}
            .invoice-box {{ max-width: 850px; margin: auto; padding: 20px; border: 1px solid #cbd5e1; border-radius: 4px; }}
            .top-header {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
            .company-title {{ font-size: 24px; font-weight: 700; color: #0f172a; margin: 0; letter-spacing: 0.5px; }}
            .company-details {{ color: #475569; font-size: 11px; line-height: 1.4; margin-top: 3px; }}
            .doc-title {{ font-size: 18px; font-weight: 700; color: #1e3a8a; margin: 0; text-transform: uppercase; letter-spacing: 1px; }}
            .client-card {{ background: #f8fafc; padding: 10px; border: 1px solid #e2e8f0; border-radius: 4px; width: 100%; box-sizing: border-box; }}
            .items-table {{ width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11.5px; }}
            .items-table th {{ background-color: #0f172a; color: white; padding: 6px 8px; font-weight: 600; text-transform: uppercase; font-size: 11px; }}
            .summary-wrapper {{ width: 100%; margin-top: 15px; }}
            .totals-table {{ width: 280px; float: right; border-collapse: collapse; font-size: 12px; }}
            .totals-table td {{ padding: 4px 6px; border-bottom: 1px solid #e2e8f0; }}
            .grand-total-row {{ font-weight: 700; background-color: #f1f5f9; color: #0f172a; font-size: 13px; border-top: 1px solid #94a3b8; border-bottom: 2px double #94a3b8 !important; }}
            .hsn-table {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
            .hsn-table th {{ background-color: #f1f5f9; color: #0f172a; padding: 5px; border: 1px solid #cbd5e1; text-align: right; }}
            .hsn-table th.left {{ text-align: center; }}
            .hsn-table td {{ padding: 5px; border: 1px solid #e2e8f0; text-align: right; color: #334155; }}
            .hsn-table td.left {{ text-align: center; }}
            .signatures-box {{ clear: both; margin-top: 45px; width: 100%; }}
            .sig-text {{ font-size: 11px; color: #475569; border-top: 1px solid #64748b; width: 160px; text-align: center; padding-top: 4px; }}
        </style>
    </head>
    <body>
        <div class="invoice-box">
            <table class="top-header">
                <tr>
                    <td style="width: 55%; vertical-align: top;">
                        <div class="company-title">GALAXY AUTOMOBILES</div>
                        <div class="company-details">
                            Near Ambedkar Chowk, Bareilly Road, Kichha, Uttarakhand<br>
                            <b>GSTIN:</b> {shop_gst}<br>
                            <b>Mobile:</b> 9837103330, 8279413595
                        </div>
                    </td>
                    <td style="width: 45%; text-align: right; vertical-align: top;">
                        <div class="doc-title" style="margin-bottom: 8px;">{title_text}</div>
                        <table style="border-collapse: collapse; font-size: 12px; line-height: 1.6; float: right;">
                            <tr>
                                <td style="color: #64748b; font-weight: 600; text-align: left; padding-right: 15px;">Invoice No:</td>
                                <td style="font-weight: 700; text-align: right; color: #0f172a;">{inv_no}</td>
                            </tr>
                            <tr>
                                <td style="color: #64748b; font-weight: 600; text-align: left; padding-right: 15px;">Date:</td>
                                <td style="font-weight: 600; text-align: right; color: #0f172a;">{inv_date}</td>
                            </tr>
                            <tr>
                                <td style="color: #64748b; font-weight: 600; text-align: left; padding-right: 15px;">Vehicle No:</td>
                                <td style="font-weight: 700; text-align: right; color: #1e3a8a;">{veh if veh else "-"}</td>
                            </tr>
                            <tr>
                                <td style="color: #64748b; font-weight: 600; text-align: left; padding-right: 15px;">KM Reading:</td>
                                <td style="font-weight: 700; text-align: right; color: #0f172a;">{km_reading if km_reading else "-"}</td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>

            <div class="client-card">
                <span style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase; display: block; margin-bottom: 2px;">Customer Details:</span>
                <span style="font-size: 13px; font-weight: 700; color: #0f172a;">{customer_name if customer_name else "Walk-In Customer"}</span>
                {f'<br><span style="color:#475569;">{customer_address}</span>' if customer_address else ''}
                {f'<br><span style="font-weight:600; color:#334155;">GSTIN: {customer_gst}</span>' if customer_gst else ''}
            </div>

            <table class="items-table">
                <thead>
                    <tr>
                        <th style="text-align: center; width: 5%;">#</th>
                        <th style="text-align: left; width: 50%;">Description</th>
                        <th style="text-align: center; width: 10%;">HSN</th>
                        <th style="text-align: center; width: 8%;">Qty</th>
                        <th style="text-align: right; width: 12%;">Rate</th>
                        <th style="text-align: right; width: 15%;">Amount (₹)</th>
                    </tr>
                </thead>
                <tbody>
                    {table_content}
                </tbody>
            </table>

            <div class="summary-wrapper">
                <table class="totals-table">
                    <tr><td>Parts Total:</td><td style="text-align: right;">₹{parts:.2f}</td></tr>
                    <tr><td>Labor Total:</td><td style="text-align: right;">₹{labor:.2f}</td></tr>
                    <tr><td>CGST (9%):</td><td style="text-align: right;">₹{gst/2:.2f}</td></tr>
                    <tr><td>SGST (9%):</td><td style="text-align: right;">₹{gst/2:.2f}</td></tr>
                    <tr class="grand-total-row"><td>Grand Total:</td><td style="text-align: right;">₹{total:.2f}</td></tr>
                    <tr><td style="color: #059669;">Amount Paid:</td><td style="text-align: right; color: #059669; font-weight: 600;">₹{paid:.2f}</td></tr>
                    <tr><td style="color: #b91c1c;">Balance Due:</td><td style="text-align: right; color: #b91c1c; font-weight: 600;">₹{balance:.2f}</td></tr>
                </table>
                {hsn_html}
                <div style="clear: both;"></div>
            </div>

            <table class="signatures-box">
                <tr>
                    <td style="width: 50%; vertical-align: bottom; height: 50px;">
                        <div class="sig-text">Customer Signature</div>
                    </td>
                    <td style="width: 50%; vertical-align: bottom; text-align: right; height: 50px;">
                        <div class="sig-text" style="float: right;">Authorized Signatory<br><b style="color:#0f172a;">Galaxy Automobiles</b></div>
                    </td>
                </tr>
            </table>
        </div>
    </body>
    </html>
    """

# ==========================================
# PAYMENT RECEIPT GENERATOR
# ==========================================
def generate_receipt_html(receipt_date, inv_no, veh, customer_name, total_amount, paid_amount, shop_gst="05AHYPG3733B2ZG", km_reading=""):
    balance = total_amount - paid_amount
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; margin: 0; padding: 20px; background: #fff; font-size: 14px; }}
            .receipt-box {{ max-width: 600px; margin: auto; padding: 30px; border: 2px solid #1e3a8a; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); }}
            .header {{ text-align: center; margin-bottom: 20px; border-bottom: 1px solid #cbd5e1; padding-bottom: 15px; }}
            .title {{ font-size: 24px; font-weight: 800; color: #1e3a8a; letter-spacing: 1px; margin-bottom: 5px; }}
            .subtitle {{ font-size: 18px; font-weight: bold; margin-top: 15px; text-transform: uppercase; color: #475569; letter-spacing: 0.5px; background-color: #f8fafc; display: inline-block; padding: 5px 15px; border-radius: 4px; border: 1px solid #e2e8f0; }}
            .details-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            .details-table td {{ padding: 12px 10px; border-bottom: 1px dashed #cbd5e1; }}
            .label {{ font-weight: 600; color: #64748b; width: 45%; }}
            .value {{ font-weight: 700; color: #0f172a; text-align: right; }}
            .amount-box {{ background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 20px; text-align: center; margin-top: 25px; border-radius: 8px; }}
            .amount-text {{ font-size: 28px; font-weight: 800; color: #166534; }}
            .footer {{ margin-top: 40px; text-align: center; font-size: 12px; color: #64748b; font-style: italic; }}
        </style>
    </head>
    <body>
        <div class="receipt-box">
            <div class="header">
                <div class="title">GALAXY AUTOMOBILES</div>
                <div style="font-size: 12px; color: #475569;">Near Ambedkar Chowk, Bareilly Road, Kichha, Uttarakhand<br>GSTIN: {shop_gst}<br>Mobile: 9837103330, 8279413595</div>
                <div class="subtitle">Payment Receipt</div>
            </div>
            <table class="details-table">
                <tr><td class="label">Receipt Date:</td><td class="value">{receipt_date}</td></tr>
                <tr><td class="label">Linked Document No:</td><td class="value">{inv_no}</td></tr>
                <tr><td class="label">Customer Name:</td><td class="value">{customer_name if customer_name else 'Walk-In Customer'}</td></tr>
                <tr><td class="label">Vehicle No:</td><td class="value" style="color: #1e3a8a;">{veh if veh else '-'}</td></tr>
                <tr><td class="label">KM Reading:</td><td class="value">{km_reading if km_reading else '-'}</td></tr>
                <tr><td class="label">Total Invoice Amount:</td><td class="value">₹{total_amount:,.2f}</td></tr>
                <tr><td class="label">Balance Due (Pending):</td><td class="value" style="color: #b91c1c;">₹{balance:,.2f}</td></tr>
            </table>
            <div class="amount-box">
                <div style="font-size: 13px; color: #166534; margin-bottom: 5px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase;">Amount Received</div>
                <div class="amount-text">₹{paid_amount:,.2f}</div>
            </div>
            <div style="margin-top: 60px; display: flex; justify-content: space-between;">
                <div style="border-top: 1px solid #64748b; padding-top: 8px; width: 40%; text-align: center; font-size: 12px; font-weight: 600; color: #475569;">Customer Signature</div>
                <div style="border-top: 1px solid #64748b; padding-top: 8px; width: 40%; text-align: center; font-size: 12px; font-weight: 600; color: #475569;">Authorized Signatory</div>
            </div>
            <div class="footer">Thank you for choosing Galaxy Automobiles!</div>
        </div>
    </body>
    </html>
    """

# ==========================================
# SIDEBAR DASHBOARD
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #1a237e;'>Control Panel</h2>", unsafe_allow_html=True)
    st.info(f"📅 **Date:** {datetime.now(IST).strftime('%d-%m-%Y')}\n\n🕒 **Time:** {datetime.now(IST).strftime('%I:%M %p')}")
    if st.button("🔄 FORCE REFRESH DATA"): st.rerun()
    st.divider()
    if bills_data:
        total_rev = sum(float(b['total_amount']) for b in bills_data if not b.get('is_estimate') and b.get('payment_status') != 'Cancelled')
        st.metric("💰 Total Revenue (All Time)", f"₹{total_rev:,.0f}")
        
        today = datetime.now(IST).strftime('%d-%m-%Y')
        today_rev = sum(float(b.get('amount_paid', 0)) for b in bills_data if not b.get('is_estimate') and get_ist(b['created_at'])[0] == today)
        st.metric("📊 Today's Collection", f"₹{today_rev:,.0f}")

# ==========================================
# 3. CORE TABS
# ==========================================
tab1, tab2 = st.tabs(["🧾 PROFESSIONAL BILLING", "📂 DATABASE & EXCEL"])

if 'edit_bill' not in st.session_state:
    st.session_state.edit_bill = None

# --- TAB 1: BILLING ---
with tab1:
    if st.session_state.edit_bill:
        st.warning(f"✏️ **EDIT MODE ACTIVE:** You are currently editing Document **{st.session_state.edit_bill['invoice_number']}**.")
        if st.button("❌ Cancel Edit Mode"):
            st.session_state.edit_bill = None
            st.rerun()
    
    st.write("### 📝 Invoice Generation Engine")
    
    b_veh_val = st.session_state.edit_bill['vehicle_number'] if st.session_state.edit_bill else ""
    b_km_val = st.session_state.edit_bill.get('km_reading', '') if st.session_state.edit_bill else ""
    c_name_val = st.session_state.edit_bill['customer_name'] if st.session_state.edit_bill else ""
    c_gst_val = st.session_state.edit_bill['customer_gst'] if st.session_state.edit_bill else ""
    shop_gst_val = st.session_state.edit_bill['shop_gst'] if st.session_state.edit_bill else "05AHYPG3733B2ZG"
    c_addr_val = st.session_state.edit_bill['customer_address'] if st.session_state.edit_bill else ""
    
    doc_options = ["Tax Invoice", "Estimate", "Pre-Invoice"]
    doc_idx = 0
    if st.session_state.edit_bill:
        orig_no = st.session_state.edit_bill['invoice_number']
        if orig_no.startswith('EST-'): doc_idx = 1
        elif orig_no.startswith('PRE-'): doc_idx = 2

    with st.container():
        c1, c1a, c2, c3 = st.columns([1.5, 1, 2, 1.5])
        b_veh = c1.text_input("Vehicle Number", value=b_veh_val).upper()
        b_km = c1a.text_input("KM Reading", value=b_km_val)
        c_name = c2.text_input("Customer / Billing Name", value=c_name_val)
        c_gst = c3.text_input("Customer GSTIN (Optional)", value=c_gst_val)
        
        c4, c5, c6 = st.columns(3)
        shop_gst_input = c4.text_input("Workshop GSTIN", value=shop_gst_val)
        c_addr = c5.text_input("Customer Address (Optional)", value=c_addr_val)
        doc_type = c6.selectbox("Document Type", doc_options, index=doc_idx)

    st.write("#### 🛒 Itemized Parts & Labor")
    
    # Pre-defined descriptive HSN Mapping Options
    hsn_mapping = {
        "8708": "8708 - Auto Parts",
        "8709": "8709 - CV Parts",
        "8507": "8507 - Batteries",
        "4011": "4011 - Tyres",
        "2710": "2710 - Engine Oil / Lubes",
        "9987": "9987 - Labor Services"
    }
    
    if st.session_state.edit_bill and st.session_state.edit_bill.get('invoice_details'):
        raw_items = json.loads(st.session_state.edit_bill['invoice_details'])
        # Map old raw HSN codes to new descriptive labels for the UI
        for it in raw_items:
            old_hsn = str(it.get('HSN', '-')).strip()
            if old_hsn in hsn_mapping:
                it['HSN'] = hsn_mapping[old_hsn]
        init_items = pd.DataFrame(raw_items)
    else:
        init_items = pd.DataFrame([
            {"Type": "Part", "Description": "Premium Engine Oil", "HSN": "2710 - Engine Oil / Lubes", "Qty": 1.0, "Rate": 2500.0},
            {"Type": "Labor", "Description": "General Service", "HSN": "9987 - Labor Services", "Qty": 1.0, "Rate": 1200.0}
        ])
    
    edited_items = st.data_editor(
        init_items, num_rows="dynamic", use_container_width=True,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Part", "Labor"], required=True),
            "Description": st.column_config.TextColumn("Description / Part Name", required=True),
            "HSN": st.column_config.SelectboxColumn("HSN Code", options=["8708 - Auto Parts", "8709 - CV Parts", "8507 - Batteries", "4011 - Tyres", "2710 - Engine Oil / Lubes", "9987 - Labor Services", "-"]),
            "Qty": st.column_config.NumberColumn("Qty", min_value=0.1, format="%.1f"),
            "Rate": st.column_config.NumberColumn("Rate (₹)", min_value=0.0, format="%.2f")
        }
    )
    
    p_total = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items.iterrows() if r['Type'] == "Part")
    l_total = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items.iterrows() if r['Type'] == "Labor")
    
    st.divider()
    col_d1, col_d2 = st.columns(2)
    l_disc = col_d1.number_input("Labor Discount (%)", min_value=0, max_value=100, value=0)
    apply_gst = col_d2.checkbox("Apply 18% GST (9% CGST / 9% SGST)", value=True)
    
    final_labor = l_total - (l_total * (l_disc / 100))
    gst_val = (p_total + final_labor) * 0.18 if apply_gst else 0.0
    cgst_val = gst_val / 2
    sgst_val = gst_val / 2
    grand_total = round(p_total + final_labor + gst_val)
    
    st.markdown(f"""
        <div style="background-color: #e8eaf6; padding: 15px; border-radius: 8px; border: 1px solid #c5cae9; text-align: center; margin-bottom: 20px;">
            <h3 style="margin:0; color: #1a237e;">GRAND TOTAL: ₹{grand_total:,.2f}</h3>
            <p style="margin:5px 0 0 0; font-size: 14px; color: #3949ab;">Parts: ₹{p_total:,.2f} | Labor: ₹{final_labor:,.2f} | CGST: ₹{cgst_val:,.2f} | SGST: ₹{sgst_val:,.2f}</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.form("save_invoice_form"):
        c_pay1, c_pay2 = st.columns(2)
        
        paid_val = float(st.session_state.edit_bill['amount_paid']) if st.session_state.edit_bill else 0.0
        paid_amt = c_pay1.number_input("Amount Received Now (₹)", min_value=0.0, max_value=float(grand_total), value=paid_val)
        due_date_input = c_pay2.date_input("Payment Due Date (If Pending Balance)")
        
        submit_label = "💾 SAVE DOCUMENT CHANGES" if st.session_state.edit_bill else "🚀 GENERATE NEW DOCUMENT"
        
        if st.form_submit_button(submit_label):
            if b_veh:
                items_list = edited_items.to_dict('records')
                
                # FORCE HSN 9987 INTO THE DATABASE FOR LABOR ITEMS
                for it in items_list:
                    if it.get('Type') == 'Labor':
                        it['HSN'] = '9987 - Labor Services'
                        
                is_final_bill = (doc_type == "Tax Invoice")
                payment_status = "Paid" if paid_amt >= grand_total else ("Pending" if paid_amt == 0 else "Partial")
                
                is_editing_existing_tax_invoice = False
                if st.session_state.edit_bill:
                    orig_no = st.session_state.edit_bill.get('invoice_number', '')
                    if (orig_no.startswith('GA/26-27/') or orig_no.startswith('25/26-A')) and doc_type == "Tax Invoice":
                        is_editing_existing_tax_invoice = True

                if is_editing_existing_tax_invoice:
                    new_assigned_no = st.session_state.edit_bill['invoice_number']
                else:
                    if doc_type == "Tax Invoice":
                        tax_nums = [0]
                        for b in bills_data:
                            inv_str = b.get('invoice_number', '')
                            if inv_str.startswith('GA/26-27/'):
                                num_part = inv_str.split('/')[-1]
                                if num_part.isdigit():
                                    tax_nums.append(int(num_part))
                        next_num = max(tax_nums) + 1
                        new_assigned_no = f"GA/26-27/{next_num}"
                    
                    elif doc_type == "Estimate":
                        est_nums = [0]
                        for b in bills_data:
                            inv_str = b.get('invoice_number', '')
                            if inv_str.startswith('EST-'):
                                num_part = inv_str.split('-')[-1]
                                if num_part.isdigit():
                                    est_nums.append(int(num_part))
                        next_num = max(est_nums) + 1
                        new_assigned_no = f"EST-{next_num:03d}"
                    
                    else:
                        pre_nums = [0]
                        for b in bills_data:
                            inv_str = b.get('invoice_number', '')
                            if inv_str.startswith('PRE-'):
                                num_part = inv_str.split('-')[-1]
                                if num_part.isdigit():
                                    pre_nums.append(int(num_part))
                        next_num = max(pre_nums) + 1
                        new_assigned_no = f"PRE-{next_num:03d}"
                
                db_payload = {
                    "invoice_number": new_assigned_no, "vehicle_number": b_veh, "customer_name": c_name, 
                    "customer_gst": c_gst, "customer_address": c_addr, "total_amount": grand_total, 
                    "amount_paid": paid_amt, "parts_cost": p_total, "final_labor": final_labor, 
                    "gst_amount": gst_val, "payment_status": payment_status, "invoice_details": json.dumps(items_list),
                    "is_estimate": (not is_final_bill), "shop_gst": shop_gst_input, "due_date": str(due_date_input),
                    "km_reading": str(b_km)
                }

                if st.session_state.edit_bill:
                    supabase.table("galaxy_billing").delete().eq("id", st.session_state.edit_bill['id']).execute()
                    supabase.table("galaxy_billing").insert(db_payload).execute()
                    st.success(f"✅ Document successfully saved as: {new_assigned_no}")
                    st.session_state.edit_bill = None
                else:
                    supabase.table("galaxy_billing").insert(db_payload).execute()
                    st.success(f"✅ {doc_type} {new_assigned_no} created successfully!")
                
                time.sleep(1.5); st.rerun()
            else:
                st.error("Please enter a Vehicle Number.")

    # --- LIVE SYSTEM FORM PREVIEW (DRAFT SCREEN) ---
    st.write("---")
    st.write("### 👁️ Live Invoice Draft Preview (System Form)")
    
    is_editing_existing_tax_invoice = False
    if st.session_state.edit_bill:
        orig_no = st.session_state.edit_bill.get('invoice_number', '')
        if (orig_no.startswith('GA/26-27/') or orig_no.startswith('25/26-A')) and doc_type == "Tax Invoice":
            is_editing_existing_tax_invoice = True

    if is_editing_existing_tax_invoice:
        draft_inv_no = st.session_state.edit_bill['invoice_number']
    else:
        if doc_type == "Tax Invoice":
            tax_nums = [0]
            for b in bills_data:
                inv_str = b.get('invoice_number', '')
                if inv_str.startswith('GA/26-27/'):
                    num_part = inv_str.split('/')[-1]
                    if num_part.isdigit():
                        tax_nums.append(int(num_part))
            next_num = max(tax_nums) + 1
            draft_inv_no = f"GA/26-27/{next_num}"
        elif doc_type == "Estimate":
            est_nums = [0]
            for b in bills_data:
                inv_str = b.get('invoice_number', '')
                if inv_str.startswith('EST-'):
                    num_part = inv_str.split('-')[-1]
                    if num_part.isdigit():
                        est_nums.append(int(num_part))
            next_num = max(est_nums) + 1
            draft_inv_no = f"EST-{next_num:03d}"
        else:
            pre_nums = [0]
            for b in bills_data:
                inv_str = b.get('invoice_number', '')
                if inv_str.startswith('PRE-'):
                    num_part = inv_str.split('-')[-1]
                    if num_part.isdigit():
                        pre_nums.append(int(num_part))
            next_num = max(pre_nums) + 1
            draft_inv_no = f"PRE-{next_num:03d}"
            
    draft_date = get_ist(st.session_state.edit_bill['created_at'])[0] if st.session_state.edit_bill else datetime.now(IST).strftime('%d-%m-%Y')
    
    # Process draft items for preview to force 9987 for labor
    preview_items = edited_items.to_dict('records')
    for it in preview_items:
        if it.get('Type') == 'Labor':
            it['HSN'] = '9987 - Labor Services'
            
    draft_html = generate_invoice_html(
        inv_no=draft_inv_no, inv_date=draft_date, veh=b_veh,
        parts=p_total, labor=final_labor, gst=gst_val, total=grand_total, paid=paid_amt,
        doc_type=doc_type, items_list=preview_items, shop_gst=shop_gst_input,
        db_status="", customer_name=c_name, customer_gst=c_gst, customer_address=c_addr,
        km_reading=b_km
    )
    st.components.v1.html(draft_html, height=520, scrolling=True)

# --- TAB 2: MANAGE DATABASE ---
with tab2:
    st.write("### 📂 Document Repository & Database")
    
    if bills_data:
        st.write("#### 📥 Secure Excel Backup")
        df_bills = pd.DataFrame(bills_data)
        st.download_button("📊 Export All Galaxy Bills to Excel (.xlsx)", data=to_excel(df_bills), file_name=f"Galaxy_Full_Backup_{datetime.now(IST).strftime('%d-%m-%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.divider()
        
        # --- CA HSN SUMMARY REPORT ---
        st.write("#### 📑 CA Reports (HSN/SAC Summary)")
        st.info("Select a date range to generate a comprehensive HSN summary for your Chartered Accountant. This strictly includes valid Tax Invoices only (estimates, pre-invoices, and cancelled bills are ignored).")
        
        c_start, c_end = st.columns(2)
        start_date = c_start.date_input("Report Start Date", datetime.now(IST).date().replace(day=1))
        end_date = c_end.date_input("Report End Date", datetime.now(IST).date())
        
        if st.button("📊 GENERATE HSN REPORT FOR CA", type="primary"):
            hsn_agg = {}
            valid_invoices_count = 0
            
            for b in bills_data:
                try:
                    b_date = datetime.fromisoformat(b['created_at'].replace('Z', '+00:00')).astimezone(IST).date()
                except: continue
                
                inv_no = b.get('invoice_number', '')
                is_cancelled = b.get('payment_status') == 'Cancelled'
                # Strictly define a valid Tax Invoice as one that is NOT cancelled, NOT an estimate, and NOT a pre-invoice
                is_tax_invoice = not inv_no.startswith('EST-') and not inv_no.startswith('PRE-')
                
                if start_date <= b_date <= end_date and is_tax_invoice and not is_cancelled:
                    
                    # --- SAFELY PARSE JSON DETAILS ---
                    raw_details = b.get('invoice_details', '[]')
                    if not raw_details: raw_details = '[]'
                    
                    try:
                        details = json.loads(raw_details) if isinstance(raw_details, str) else raw_details
                    except:
                        details = []
                        
                    # Force into a list if it's a single dictionary or string
                    if isinstance(details, dict):
                        details = [details]
                    elif not isinstance(details, list):
                        details = []
                    # ---------------------------------
                    
                    valid_invoices_count += 1
                    final_labor = float(b.get('final_labor', 0))
                    has_gst = float(b.get('gst_amount', 0)) > 0
                    
                    # Safely calculate total labor for proportion discounting
                    raw_labor_total = sum(float(i.get('Qty', 1)) * float(i.get('Rate', 0)) for i in details if isinstance(i, dict) and i.get('Type') == 'Labor')
                    labor_multiplier = (final_labor / raw_labor_total) if raw_labor_total > 0 else 1.0
                    
                    for item in details:
                        if not isinstance(item, dict): 
                            continue # Skip anything that isn't a proper item dictionary
                            
                        # FORCE LABOR HSN TO 9987 AND STRIP DESCRIPTIONS FOR CA REPORT
                        raw_hsn = str(item.get('HSN', '-')).strip()
                        clean_hsn = raw_hsn.split()[0] if raw_hsn and raw_hsn != '-' else '-'
                        hsn = "9987" if item.get('Type') == 'Labor' else clean_hsn
                        
                        qty = float(item.get('Qty', 1))
                        raw_amt = qty * float(item.get('Rate', 0))
                        taxable = raw_amt * labor_multiplier if item.get('Type') == 'Labor' else raw_amt
                        
                        if hsn not in hsn_agg:
                            hsn_agg[hsn] = {"Total Qty": 0.0, "Taxable Value (₹)": 0.0, "CGST (₹)": 0.0, "SGST (₹)": 0.0, "Total Amount (₹)": 0.0}
                        
                        cgst = (taxable * 0.09) if has_gst else 0.0
                        sgst = (taxable * 0.09) if has_gst else 0.0
                        
                        hsn_agg[hsn]["Total Qty"] += qty
                        hsn_agg[hsn]["Taxable Value (₹)"] += taxable
                        hsn_agg[hsn]["CGST (₹)"] += cgst
                        hsn_agg[hsn]["SGST (₹)"] += sgst
                        hsn_agg[hsn]["Total Amount (₹)"] += (taxable + cgst + sgst)
                        
            if hsn_agg:
                st.success(f"Successfully processed {valid_invoices_count} valid Tax Invoices for this report.")
                report_df = pd.DataFrame.from_dict(hsn_agg, orient='index').reset_index()
                report_df.rename(columns={'index': 'HSN/SAC Code'}, inplace=True)
                report_df = report_df.round(2)
                
                st.dataframe(report_df, use_container_width=True, hide_index=True)
                
                ca_excel = io.BytesIO()
                with pd.ExcelWriter(ca_excel, engine='xlsxwriter') as writer:
                    report_df.to_excel(writer, index=False, sheet_name='CA_HSN_Summary')
                st.download_button("📥 DOWNLOAD EXCEL FOR CA", data=ca_excel.getvalue(), file_name=f"CA_HSN_Summary_{start_date}_to_{end_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.warning("No valid Tax Invoices found in this date range.")
                
        st.divider()
        # --- END CA REPORT ---
        
        search_inv = st.text_input("🔍 Enter exact Document Number to Print/Edit/Cancel (e.g., GA/26-27/1)").strip().upper()
        if search_inv:
            target = next((b for b in bills_data if b['invoice_number'] == search_inv), None)
            if target:
                st.success(f"Found: {target['vehicle_number']} | Total: ₹{target['total_amount']} | Status: {target['payment_status']}")
                
                saved_name = target.get('customer_name', '')
                receipt_name = st.text_input("✏️ Edit Customer Name specifically for the Payment Receipt:", value=saved_name if saved_name else "Walk-In Customer")
                
                st.write("#### Available Actions:")
                c1, c2, c3, c4 = st.columns(4)
                
                with c1:
                    date_str, _ = get_ist(target['created_at'])
                    
                    if target['invoice_number'].startswith('PRE-'):
                        d_type = "Pre-Invoice"
                    elif target['invoice_number'].startswith('EST-'):
                        d_type = "Estimate"
                    else:
                        d_type = "Tax Invoice"
                    
                    html = generate_invoice_html(
                        inv_no=target['invoice_number'], inv_date=date_str, veh=target['vehicle_number'],
                        parts=float(target.get('parts_cost', 0)), labor=float(target.get('final_labor', 0)),
                        gst=float(target.get('gst_amount', 0)), total=float(target['total_amount']),
                        paid=float(target.get('amount_paid', 0)), doc_type=d_type,
                        items_list=json.loads(target.get('invoice_details', '[]')), shop_gst=target.get('shop_gst', '05AHYPG3733B2ZG'),
                        db_status=target.get('payment_status', ''), customer_name=target.get('customer_name', ''),
                        customer_gst=target.get('customer_gst', ''), customer_address=target.get('customer_address', ''),
                        km_reading=target.get('km_reading', '')
                    )
                    st.download_button("🖨️ Print Original Bill", data=html, file_name=f"{target['invoice_number'].replace('/', '_')}.html", mime="text/html", use_container_width=True)
                
                with c2:
                    receipt_html = generate_receipt_html(
                        receipt_date=datetime.now(IST).strftime('%d-%m-%Y'),
                        inv_no=target['invoice_number'],
                        veh=target['vehicle_number'],
                        customer_name=receipt_name,
                        total_amount=float(target['total_amount']),
                        paid_amount=float(target.get('amount_paid', 0)),
                        shop_gst=target.get('shop_gst', '05AHYPG3733B2ZG'),
                        km_reading=target.get('km_reading', '')
                    )
                    st.download_button("🧾 Print Custom Receipt", data=receipt_html, file_name=f"Receipt_{target['invoice_number'].replace('/', '_')}.html", mime="text/html", use_container_width=True)

                with c3:
                    if st.button("✏️ EDIT INVOICE DATA", type="secondary", use_container_width=True):
                        st.session_state.edit_bill = target
                        st.rerun()

                with c4:
                    if st.button("❌ CANCEL THIS DOCUMENT", type="primary", use_container_width=True):
                        supabase.table("galaxy_billing").update({"payment_status": "Cancelled"}).eq("id", target['id']).execute()
                        st.rerun()
                
                st.write("---")
                st.write("#### 📄 Document View (System Form)")
                st.components.v1.html(html, height=520, scrolling=True)
            else:
                st.warning("Document not found.")

        st.write("#### 🧾 Recent Documents")
        cols = [c for c in ['invoice_number', 'vehicle_number', 'customer_name', 'total_amount', 'amount_paid', 'payment_status'] if c in df_bills.columns]
        st.dataframe(df_bills[cols], use_container_width=True, hide_index=True)
    else:
        st.info("No documents found in the database yet.")
