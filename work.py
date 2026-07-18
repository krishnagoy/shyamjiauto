import streamlit as st
from supabase import create_client
import urllib.parse
from datetime import datetime, timezone, timedelta
import pandas as pd
import json
import time
import io
import streamlit.components.v1 as components

# ==========================================
# 1. ULTRA-PREMIUM UI & SECURITY
# ==========================================
st.set_page_config(page_title="Shyamji ERP", layout="wide", page_icon="🛠️", initial_sidebar_state="expanded")

# --- LOGIN SYSTEM ---
VALID_USERS = {"krishna": "admin123", "manager": "shyamji2026"}
if 'logged_in' not in st.session_state: 
    st.session_state['logged_in'] = False
    st.session_state['user'] = ""

if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        with st.form("login"):
            st.markdown("<h2 style='text-align:center; color:#8b0000;'>🔒 Authorized Access</h2>", unsafe_allow_html=True)
            u = st.text_input("Username").lower()
            p = st.text_input("Password", type="password")
            if st.form_submit_button("LOGIN", type="primary"):
                if u in VALID_USERS and VALID_USERS[u] == p:
                    st.session_state['logged_in'] = True
                    st.session_state['user'] = u
                    st.rerun()
                else: st.error("Invalid Credentials")
    st.stop()

# --- CSS STYLING ---
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    .stApp { background-color: #f0f2f6; font-family: sans-serif; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 5px 5px 0 0; padding: 10px 20px; box-shadow: 0px -2px 5px rgba(0,0,0,0.05); }
    .stTabs [aria-selected="true"] { background-color: #8b0000; color: white !important; border-bottom: none; font-weight: bold; }
    .stButton>button { width: 100%; border-radius: 6px; background-color: #8b0000; color: white; font-weight: 600; border: none; padding: 10px; transition: all 0.3s ease; }
    .stButton>button:hover { background-color: #660000; box-shadow: 0 4px 8px rgba(0,0,0,0.2); color: white;}
    [data-testid="stMetric"] { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #8b0000; }
    [data-testid="stForm"] { background-color: white; padding: 20px; border-radius: 10px; border: 1px solid #e0e0e0; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <div style="background: linear-gradient(135deg, #8b0000 0%, #d32f2f 100%); padding: 25px; border-radius: 10px; color: white; text-align: center; margin-bottom: 25px; margin-top: -40px; box-shadow: 0 4px 15px rgba(139,0,0,0.3);">
        <h1 style="margin:0; font-size: 36px; font-weight: 800; letter-spacing: 1px; color: white;">🛠️ SHRI SHYAMJI AUTO SERVICE CENTER</h1>
        <p style="margin:5px 0 0 0; font-size: 16px; opacity: 0.9;">Authorized Maruti Suzuki Service Station | Master ERP</p>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 2. DATABASE SETUP & HTML GENERATORS
# ==========================================
URL = "https://xthuqvzuvsdbtqaxgrlq.supabase.co"
KEY = "sb_publishable_vniJjmRGyI50rLx_Oyctnw_v6gqkw_a"
supabase = create_client(URL, KEY)
IST = timezone(timedelta(hours=5, minutes=30))

def safe_float(value):
    try: return float(value) if value is not None else 0.0
    except: return 0.0

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
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ==========================================
# CUSTOM NUMBER TO WORDS (INDIAN SYSTEM)
# ==========================================
def number_to_words_indian(n):
    if n == 0: return "Zero"
    words = {1: 'One', 2: 'Two', 3: 'Three', 4: 'Four', 5: 'Five', 6: 'Six', 7: 'Seven', 8: 'Eight', 9: 'Nine', 10: 'Ten',
             11: 'Eleven', 12: 'Twelve', 13: 'Thirteen', 14: 'Fourteen', 15: 'Fifteen', 16: 'Sixteen', 17: 'Seventeen', 18: 'Eighteen', 19: 'Nineteen',
             20: 'Twenty', 30: 'Thirty', 40: 'Forty', 50: 'Fifty', 60: 'Sixty', 70: 'Seventy', 80: 'Eighty', 90: 'Ninety'}
    
    def num_to_words(num):
        if num == 0: return ""
        elif num < 20: return words[num] + " "
        elif num < 100: return words[(num // 10) * 10] + " " + num_to_words(num % 10)
        elif num < 1000: return words[num // 100] + " Hundred " + num_to_words(num % 100)
        elif num < 100000: return num_to_words(num // 1000) + " Thousand " + num_to_words(num % 1000)
        elif num < 10000000: return num_to_words(num // 100000) + " Lakh " + num_to_words(num % 100000)
        else: return num_to_words(num // 10000000) + " Crore " + num_to_words(num % 10000000)
    
    return "Rupees " + num_to_words(int(n)).strip() + " Only"

def generate_jobcard_html(jc_no, jc_date, veh, name, phone, km, advisor, mech, jc_type, demands, parts):
    demand_rows = "".join([f"<tr><td style='border: 1px solid #ddd; padding: 8px;'>{idx+1}</td><td style='border: 1px solid #ddd; padding: 8px;'>{d.strip()}</td><td style='border: 1px solid #ddd; padding: 8px; color: green; font-weight:bold;'>Assigned</td></tr>" for idx, d in enumerate(demands.split(',')) if d.strip()]) if demands and demands.strip() else "<tr><td colspan='3' style='border: 1px solid #ddd; padding: 8px; text-align:center; color:#777;'>No demanding works cataloged yet.</td></tr>"
    parts_rows = "".join([f"<tr><td style='border: 1px solid #ddd; padding: 8px;'>{idx+1}</td><td style='border: 1px solid #ddd; padding: 8px;'>{p.strip()}</td></tr>" for idx, p in enumerate(parts.split(',')) if p.strip()]) if parts and parts.strip() else "<tr><td colspan='2' style='border: 1px solid #ddd; padding: 8px; text-align:center; color:#777;'>No parts checklist cataloged yet.</td></tr>"
    
    return f"""
    <html>
    <body style="font-family: sans-serif; color: #333; padding: 20px; border: 1px solid #ccc; max-width: 800px; margin: auto;">
        <div style="text-align:center; border-bottom: 3px solid #8b0000; padding-bottom:10px;">
            <h2 style="margin:0; color:#8b0000; letter-spacing: 1px;">SHRI SHYAMJI AUTO SERVICE CENTER</h2>
            <p style="margin:4px 0; font-size: 13px;"><b>Authorized Maruti Suzuki Service Station</b><br>Near Ambedkar Chowk, Bareilly Road, Kichha</p>
            <h3 style="margin:10px 0 0 0; background:#8b0000; color:white; padding:5px; border-radius:3px;">WORKSHOP JOB CARD</h3>
        </div>
        <table style="width:100%; margin-top:15px; font-size:13px; border-collapse:collapse;" border="0">
            <tr><td><b>Job Card No:</b> {jc_no}</td><td style="text-align:right;"><b>Date/Time:</b> {jc_date}</td></tr>
            <tr><td><b>Vehicle No:</b> {veh}</td><td style="text-align:right;"><b>Odometer Reading:</b> {km} KM</td></tr>
            <tr><td><b>Customer Name:</b> {name}</td><td style="text-align:right;"><b>Phone:</b> {phone}</td></tr>
            <tr><td><b>Service Advisor:</b> {advisor}</td><td style="text-align:right;"><b>Assigned Mechanic:</b> {mech}</td></tr>
            <tr><td><b>Service Type:</b> {jc_type}</td><td></td></tr>
        </table>
        <h4 style="color:#8b0000; margin-top:20px; border-bottom:1px solid #8b0000; padding-bottom:3px;">CUSTOMER VOICE / WORK DEMANDED</h4>
        <table style="width:100%; border-collapse:collapse; font-size:12px;" border="1">
            <tr style="background:#f2f2f2;"><th style="width:10%; padding:8px;">Sl No.</th><th style="padding:8px;">Description of Job Demand</th><th style="width:20%; padding:8px;">Status</th></tr>
            {demand_rows}
        </table>
        <h4 style="color:#8b0000; margin-top:20px; border-bottom:1px solid #8b0000; padding-bottom:3px;">ESTIMATED / REQUIRED PARTS CHECKLIST</h4>
        <table style="width:100%; border-collapse:collapse; font-size:12px;" border="1">
            <tr style="background:#f2f2f2;"><th style="width:10%; padding:8px;">Sl No.</th><th style="padding:8px;">Part Description Checklist</th></tr>
            {parts_rows}
        </table>
        <div style="margin-top:40px; display:flex; justify-content:space-between; font-size:11px;">
            <div style="border-top: 1px solid #333; width: 180px; text-align: center; padding-top:5px;">Customer Signature</div>
            <div style="border-top: 1px solid #333; width: 180px; text-align: center; padding-top:5px;">Service Advisor Signature</div>
        </div>
    </body>
    </html>
    """

def generate_expense_slip_html(slip_id, date, cat, amt, desc, paid_by):
    return f"""
    <html>
    <body style="font-family: sans-serif; color: #333; padding: 25px; border: 2px dashed #333; max-width: 500px; margin: auto; background:#fffcf5;">
        <div style="text-align:center; border-bottom: 1px solid #333; padding-bottom:8px;">
            <h3 style="margin:0; color:#8b0000; letter-spacing: 1px;">SHRI SHYAMJI AUTO SERVICE CENTER</h3>
            <p style="margin:3px 0; font-size: 11px;">Kichha, Uttarakhand</p>
            <h4 style="margin:5px 0 0 0; background:#333; color:white; padding:3px; font-size:12px;">EXPENSE PAYMENT VOUCHER</h4>
        </div>
        <table style="width:100%; margin-top:15px; font-size:13px; line-height:22px;">
            <tr><td><b>Voucher No:</b> {slip_id}</td><td style="text-align:right;"><b>Date:</b> {date}</td></tr>
            <tr><td><b>Category:</b> {cat}</td><td></td></tr>
            <tr><td colspan="2"><b>Description/Notes:</b> {desc}</td></tr>
            <tr><td colspan="2" style="padding-top:10px; font-size:16px; color:#8b0000;"><b>PAID AMOUNT: ₹{float(amt):,.2f}</b></td></tr>
            <tr><td colspan="2" style="font-size:12px; font-style:italic; border-top:1px solid #eee; padding-top:5px;">Authorized By Whom: {paid_by}</td></tr>
        </table>
        <div style="margin-top:35px; display:flex; justify-content:space-between; font-size:11px;">
            <div style="border-top: 1px solid #333; width: 120px; text-align: center; padding-top:3px;">Receiver Signature</div>
            <div style="border-top: 1px solid #333; width: 120px; text-align: center; padding-top:3px;">Manager Approved</div>
        </div>
    </body>
    </html>
    """

def generate_receipt_html(receipt_date, inv_no, veh, customer_name, total_amount, paid_amount, shop_gst="05AHYPG3732A1ZK", p_mode="Cash"):
    balance = total_amount - paid_amount
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: sans-serif; color: #1e293b; margin: 0; padding: 20px; background: #fff; font-size: 14px; }}
            .receipt-box {{ max-width: 600px; margin: auto; padding: 30px; border: 2px solid #1a237e; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); }}
            .header {{ text-align: center; margin-bottom: 20px; border-bottom: 1px solid #cbd5e1; padding-bottom: 15px; }}
            .title {{ font-size: 24px; font-weight: 800; color: #1a237e; letter-spacing: 1px; margin-bottom: 5px; }}
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
                <div class="title">SHRI SHYAMJI AUTO SERVICE CENTER</div>
                <div style="font-size: 12px; color: #475569;">Authorized Maruti Suzuki Service Station<br>Near Ambedkar Chowk, Bareilly Road, Kichha<br>GSTIN: {shop_gst}</div>
                <div class="subtitle">Payment Receipt</div>
            </div>
            <table class="details-table">
                <tr><td class="label">Receipt Date:</td><td class="value">{receipt_date}</td></tr>
                <tr><td class="label">Linked Document No:</td><td class="value">{inv_no}</td></tr>
                <tr><td class="label">Customer Name:</td><td class="value">{customer_name if customer_name else 'Walk-In Customer'}</td></tr>
                <tr><td class="label">Vehicle No:</td><td class="value" style="color: #1e3a8a;">{veh if veh and veh != 'COUNTER SALE' else '-'}</td></tr>
                <tr><td class="label">Total Invoice Amount:</td><td class="value">₹{total_amount:,.2f}</td></tr>
                <tr><td class="label">Balance Due (Pending):</td><td class="value" style="color: #b91c1c;">₹{balance:,.2f}</td></tr>
                <tr><td class="label">Payment Mode:</td><td class="value">{p_mode}</td></tr>
            </table>
            <div class="amount-box">
                <div style="font-size: 13px; color: #166534; margin-bottom: 5px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase;">Amount Received</div>
                <div class="amount-text">₹{paid_amount:,.2f}</div>
            </div>
            <div style="margin-top: 60px; display: flex; justify-content: space-between;">
                <div style="border-top: 1px solid #64748b; padding-top: 8px; width: 40%; text-align: center; font-size: 12px; font-weight: 600; color: #475569;">Customer Signature</div>
                <div style="border-top: 1px solid #64748b; padding-top: 8px; width: 40%; text-align: center; font-size: 12px; font-weight: 600; color: #475569;">Authorized Signatory</div>
            </div>
        </div>
    </body>
    </html>
    """

# ==========================================
# 1. COUNTER SALE INVOICE HTML GENERATOR
# ==========================================
def generate_counter_sale_html(inv_no, inv_date, items_list, shop_gst="05AHYPG3732A1ZK", customer_name="", customer_gst="", customer_address="", customer_code="5648", customer_category="MASS"):
    rows_html = ""
    parts_sub = 0.0
    total_qty = 0.0
    
    if items_list:
        for idx, item in enumerate(items_list, 1):
            qty = safe_float(item.get('Qty', 1.0))
            rate = safe_float(item.get('Rate', 0.0))
            amt = qty * rate
            parts_sub += amt
            total_qty += qty
            
            raw_pn = item.get('Part_Number')
            p_num = str(raw_pn).replace('None', '').replace('nan', '').strip() if raw_pn is not None else ''
            raw_desc = item.get('Description')
            desc = str(raw_desc).replace('None', '').replace('nan', '').strip() if raw_desc is not None else ''
            raw_hsn = item.get('HSN')
            hsn = str(raw_hsn).replace('None', '8708').strip() if raw_hsn is not None else '8708'
            
            rows_html += f"""
            <tr style="border-bottom: 1px solid #ddd; height: 25px;">
                <td style="text-align:center;">{idx}</td>
                <td>{p_num}</td>
                <td style="text-align:center;">-</td>
                <td>{desc}</td>
                <td style="text-align:center;">{hsn}</td>
                <td style="text-align:center;">9%</td>
                <td style="text-align:center;">9%</td>
                <td style="text-align:center;">{qty:.2f}</td>
                <td style="text-align:right;">{rate:.3f}</td>
                <td style="text-align:right;">{amt:.2f}</td>
            </tr>
            """
            
    cgst = parts_sub * 0.09
    sgst = parts_sub * 0.09
    net_total = round(parts_sub + cgst + sgst)
    amount_in_words = number_to_words_indian(net_total)
    
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; font-size: 11px; color: #000; padding: 20px; }}
            .header-text {{ text-align: center; font-size: 10px; margin-bottom: 2px; }}
            .company-name {{ text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 2px; }}
            .company-address {{ text-align: center; font-size: 10px; line-height: 1.2; margin-bottom: 10px; }}
            .title-bar {{ background-color: #d3d3d3; text-align: center; font-weight: bold; padding: 4px; margin-bottom: 15px; border-top: 1px solid #000; border-bottom: 1px solid #000; font-size: 13px; }}
            .info-table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 11px; }}
            .info-table td {{ vertical-align: top; padding: 2px; }}
            .item-table {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
            .item-table th {{ background-color: #d3d3d3; font-weight: bold; text-align: center; padding: 5px; border-top: 1px solid #000; border-bottom: 1px solid #000; }}
            .item-table td {{ padding: 4px 5px; }}
            .totals-container {{ width: 100%; margin-top: 30px; display: table; }}
            .totals-left {{ display: table-cell; width: 40%; vertical-align: top; font-size: 10px; }}
            .totals-right {{ display: table-cell; width: 60%; vertical-align: top; }}
            .totals-table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
            .totals-table td {{ padding: 4px; border-bottom: 1px solid #ddd; }}
            .totals-table .val {{ text-align: right; font-weight: bold; }}
            .net-row {{ border-top: 1px solid #000; border-bottom: 1px solid #000; font-weight: bold; }}
            .signature-box {{ width: 100%; margin-top: 40px; display: table; font-size: 11px; }}
            .sig-col {{ display: table-cell; width: 33%; vertical-align: bottom; }}
        </style>
    </head>
    <body>
        <div class="header-text">ORIGINAL FOR RECIPIENT/DUPLICATE FOR TRANSPORTER/TRIPLICATE FOR SUPPLIER</div>
        <div class="company-name">SHRI SHYAMJI AUTO SERVICE CENTER</div>
        <div class="company-address">
            BAREILLY ROAD, KICHHA<br>
            UDHAM SINGH NAGAR, UTTARAKHAND, PIN: 263148<br>
            GST Registration No.: {shop_gst}
        </div>
        
        <div class="title-bar">Counter Sale Tax Invoice</div>
        
        <table class="info-table">
            <tr>
                <td style="width: 15%;">Customer Code</td><td style="width: 35%;">: {customer_code}</td>
                <td style="width: 25%; text-align:right;">Invoice No.:</td><td style="width: 25%; font-weight:bold;">{inv_no}</td>
            </tr>
            <tr>
                <td>Customer Category</td><td>: {customer_category}</td>
                <td style="text-align:right;">Date:</td><td style="font-weight:bold;">{inv_date}</td>
            </tr>
            <tr>
                <td>Name</td><td>: <b>{customer_name}</b></td>
                <td style="text-align:right;">Sale Type:</td><td>Local Sale</td>
            </tr>
            <tr>
                <td>Address</td><td>: {customer_address}</td>
                <td style="text-align:right;">Place Of Supply:</td><td>UTTARAKHAND</td>
            </tr>
            <tr>
                <td>State</td><td>: UTTARAKHAND</td>
                <td style="text-align:right;">Cust.Ref.:</td><td></td>
            </tr>
            <tr>
                <td>State Code</td><td>: 05</td>
                <td style="text-align:right;">Gate Pass No.:</td><td>GP-{inv_no.split('-')[-1]}</td>
            </tr>
            <tr>
                <td>Contact No.</td><td>: </td>
                <td style="text-align:right;">Gate Pass Date:</td><td>{inv_date}</td>
            </tr>
            <tr>
                <td>GST No..</td><td>: <b>{customer_gst}</b></td>
                <td style="text-align:right;">PAN :</td><td>{shop_gst[2:12] if len(shop_gst) > 12 else ''}</td>
            </tr>
        </table>

        <table class="item-table">
            <tr>
                <th style="width: 5%;">Srl.</th>
                <th style="width: 15%;">Part Number</th>
                <th style="width: 8%;">Batch</th>
                <th style="width: 30%;">Description</th>
                <th style="width: 10%;">HSN</th>
                <th style="width: 6%;">CGST</th>
                <th style="width: 6%;">SGST/UTGST</th>
                <th style="width: 5%;">Qty</th>
                <th style="width: 7%; text-align:right;">Rate</th>
                <th style="width: 8%; text-align:right;">Taxable Amount</th>
            </tr>
            {rows_html}
        </table>

        <div class="totals-container">
            <div class="totals-left">
                <b>Remarks :</b> ad-3000- FCS<br><br>
                CR.Days : 15 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Limit : 25000<br>
                No.of Items : {len(items_list)} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total Qty.: {total_qty:.2f}
            </div>
            <div class="totals-right">
                <table class="totals-table">
                    <tr><td>Part Sub Total Amount</td><td class="val">{parts_sub:.2f}</td></tr>
                    <tr><td>Part Total Taxable Amount</td><td class="val">{parts_sub:.2f}</td></tr>
                    <tr><td>CGST @ 9%</td><td class="val">{cgst:.2f}</td></tr>
                    <tr><td>SGST @ 9%</td><td class="val">{sgst:.2f}</td></tr>
                    <tr class="net-row"><td>Sub Total Amount</td><td class="val">{net_total:.2f}</td></tr>
                    <tr><td><b>Net Bill Amount</b><br>{amount_in_words}</td><td class="val"><br>{net_total:.2f}</td></tr>
                </table>
            </div>
        </div>

        <div class="signature-box">
            <div class="sig-col" style="text-align:left;">
                Customer Signature<br><br>
                <b>For SHRI SHYAMJI AUTO SERVICE CENTER</b><br><br><br>
                (Authorised Signatory)<br>
                Rel: 1.2.9
            </div>
            <div class="sig-col" style="text-align:center;">
                Printed By: NA
            </div>
            <div class="sig-col" style="text-align:right;">
                Created By: RAJESH GOYAL<br>
                Printed On: {inv_date}
            </div>
        </div>
    </body>
    </html>
    """

# ==========================================
# 2. JOB CARD RETAIL INVOICE HTML GENERATOR
# ==========================================
def generate_jobcard_invoice_html(inv_no, inv_date, veh, items_list, shop_gst="05AHYPG3732A1ZK", 
                          customer_name="", customer_gst="", customer_address="", 
                          jc_no="NA", model="NA", chassis="NA", sa_name="NA", mileage="NA", service_type="NA"):
    
    parts_html, labor_html = "", ""
    parts_sub, labor_sub = 0.0, 0.0
    p_idx, l_idx = 1, 1
    
    if items_list:
        for item in items_list:
            qty = safe_float(item.get('Qty', 1.0))
            rate = safe_float(item.get('Rate', 0.0))
            amt = qty * rate
            
            raw_pn = item.get('Part_Number')
            p_num = str(raw_pn).replace('None', '').replace('nan', '').strip() if raw_pn is not None else ''
            raw_desc = item.get('Description')
            desc = str(raw_desc).replace('None', '').replace('nan', '').strip() if raw_desc is not None else ''
            raw_hsn = item.get('HSN')
            hsn = str(raw_hsn).replace('None', '8708').strip() if raw_hsn is not None else '8708'
            
            if item.get('Type') == "Part":
                parts_sub += amt
                parts_html += f"""
                <tr>
                    <td>{p_idx}</td><td>{p_num}</td><td>{desc}</td><td>-</td><td>{hsn}</td>
                    <td>18%</td><td>{qty:.3f}</td><td>{rate:.2f}</td><td>{amt:.2f}</td><td>0.00</td><td></td>
                </tr>"""
                p_idx += 1
            else:
                labor_sub += amt
                labor_html += f"""
                <tr>
                    <td>{l_idx}</td><td>{p_num}</td><td>{desc}</td><td>-</td><td>998729</td>
                    <td>18%</td><td>{qty:.3f}</td><td>{rate:.2f}</td><td></td><td>0.00</td><td>{amt:.2f}</td>
                </tr>"""
                l_idx += 1

    cgst = (parts_sub + labor_sub) * 0.09
    sgst = (parts_sub + labor_sub) * 0.09
    net_total = round(parts_sub + labor_sub + cgst + sgst)
    amount_in_words = number_to_words_indian(net_total)

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; font-size: 11px; color: #000; padding: 20px; }}
            .header-text {{ text-align: center; font-weight: bold; margin-bottom: 5px; font-size: 12px; }}
            .sub-header {{ text-align: center; font-weight: bold; margin-bottom: 15px; text-decoration: underline; font-size: 14px; }}
            .info-table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 11px; }}
            .info-table td {{ vertical-align: top; width: 50%; padding: 2px; }}
            .item-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10px; }}
            .item-table th, .item-table td {{ border: 1px solid #000; padding: 4px; text-align: left; }}
            .item-table th {{ background-color: #f2f2f2; font-weight: bold; text-align: center; }}
            .totals-table {{ width: 100%; border-collapse: collapse; margin-top: 5px; font-size: 11px; border: 1px solid #000; }}
            .totals-table td {{ padding: 4px; border: 1px solid #000; }}
            .no-border {{ border: none !important; }}
            .right-align {{ text-align: right; }}
            .bold {{ font-weight: bold; }}
            .footer-notes {{ font-size: 9px; margin-top: 15px; text-align: justify; }}
            .signature-box {{ width: 100%; margin-top: 40px; display: table; }}
            .sig-col {{ display: table-cell; width: 33%; text-align: center; font-weight: bold; }}
            .gate-pass {{ border-top: 2px dashed #000; margin-top: 30px; padding-top: 10px; }}
        </style>
    </head>
    <body>
        <div class="header-text">ORIGINAL FOR RECIPIENT / DUPLICATE FOR TRANSPORTER / TRIPLICATE FOR SUPPLIER</div>
        <div class="sub-header">Job Card Retail - Tax Invoice</div>
        
        <table class="info-table">
            <tr>
                <td>
                    <b>Customer Name & Address:</b><br>
                    {customer_name}<br>
                    {customer_address}<br>
                    <b>State & Code:</b> 05-UTTARAKHAND<br>
                    <b>Mobile:</b> NA<br>
                    <b>Cust GSTIN/UIN:</b> {customer_gst}<br>
                </td>
                <td>
                    <table style="width:100%; font-size:11px;">
                        <tr><td><b>Invoice No.:</b></td><td>{inv_no}</td><td><b>Date:</b></td><td>{inv_date}</td></tr>
                        <tr><td><b>Job Card No.:</b></td><td>{jc_no}</td><td><b>Reg.No.:</b></td><td><b>{veh}</b></td></tr>
                        <tr><td><b>SA Name:</b></td><td>{sa_name}</td><td><b>Model:</b></td><td>{model}</td></tr>
                        <tr><td><b>Service type:</b></td><td>{service_type}</td><td><b>Mileage:</b></td><td>{mileage}</td></tr>
                        <tr><td><b>Place of Supply:</b></td><td>UTTARAKHAND</td><td><b>Dealer GSTIN:</b></td><td><b>{shop_gst}</b></td></tr>
                    </table>
                </td>
            </tr>
        </table>

        <table class="item-table">
            <tr>
                <th>Srl.</th><th>Part Number</th><th>Description</th><th>Batch</th><th>HSN/SAC</th>
                <th>Tax</th><th>Qty.</th><th>Rate</th><th>Taxable Amount</th><th>Tax Paid</th><th>Labour Charges</th>
            </tr>
            <tr><td colspan="11" class="bold">Parts</td></tr>
            {parts_html}
            <tr><td colspan="11" class="bold">Demanded Repairs-Others/ Suggested Jobs</td></tr>
            {labor_html}
        </table>

        <table class="totals-table">
            <tr>
                <td rowspan="6" style="width: 50%; vertical-align: bottom;">
                    <b>For SHRI SHYAMJI AUTO SERVICE CENTRE</b><br><br><br>
                    Authorised Signatory<br>
                </td>
                <td class="bold">Sub Total Amount</td>
                <td class="right-align">{parts_sub:.2f}</td>
                <td class="right-align">0.00</td>
                <td class="right-align">{labor_sub:.2f}</td>
            </tr>
            <tr>
                <td class="bold">CGST @ 9%</td>
                <td class="right-align">{(parts_sub * 0.09):.2f}</td>
                <td class="right-align"></td>
                <td class="right-align">{(labor_sub * 0.09):.2f}</td>
            </tr>
            <tr>
                <td class="bold">SGST @ 9%</td>
                <td class="right-align">{(parts_sub * 0.09):.2f}</td>
                <td class="right-align"></td>
                <td class="right-align">{(labor_sub * 0.09):.2f}</td>
            </tr>
            <tr>
                <td class="bold">Tax Collection at Source</td>
                <td class="right-align">0.00</td>
                <td></td><td></td>
            </tr>
            <tr>
                <td class="bold">Sub Total Amount</td>
                <td class="right-align">{(parts_sub * 1.18):.2f}</td>
                <td class="right-align">0.00</td>
                <td class="right-align">{(labor_sub * 1.18):.2f}</td>
            </tr>
            <tr>
                <td class="bold">Net Bill Amount (Rounded) :</td>
                <td colspan="3" class="right-align bold" style="font-size: 14px;">{net_total:.2f}</td>
            </tr>
        </table>
        
        <div style="margin-top: 5px; font-weight: bold; font-size: 11px;">
            {amount_in_words}
        </div>

        <div class="footer-notes">
            This is a system generated soft copy of invoice for insurance company records.<br>
            I acknowledge that the jobs/repairs/service carried out in my vehicle and the respective cost estimates were explained to me. I have received my vehicle after completion of all repairs being carried out to my satisfaction and I confirm that my vehicle is in good condition. I further authorize this workshop to contact me by call or sms to inform me with any other information in relation to my vehicle.
        </div>

        <div class="signature-box">
            <div class="sig-col"><br><br>Customer Signature</div>
            <div class="sig-col"></div>
            <div class="sig-col"><br><br>Authorized Signatory</div>
        </div>

        <div class="gate-pass">
            <h3 style="text-align:center; margin:0;">Gate Pass</h3>
            <table style="width:100%; font-size:11px; margin-top:10px;">
                <tr>
                    <td><b>Cust. Name:</b> {customer_name}</td>
                    <td><b>Date:</b> {inv_date}</td>
                    <td><b>Reg. No.:</b> {veh}</td>
                </tr>
            </table>
            <table class="item-table" style="margin-top:10px;">
                <tr><th>Bill.No.</th><th>Bill Date</th><th>Amount</th></tr>
                <tr><td style="text-align:center;">{inv_no}</td><td style="text-align:center;">{inv_date}</td><td style="text-align:center;">{net_total:.2f}</td></tr>
            </table>
            <div class="signature-box" style="margin-top:30px;">
                <div class="sig-col" style="text-align:left;">Customer Signature</div>
                <div class="sig-col" style="text-align:right;">Accountant Signature</div>
            </div>
        </div>
    </body>
    </html>
    """

# ==========================================
# MASTER INVOICE ROUTER
# ==========================================
def generate_master_invoice_html(inv_no, inv_date, veh, items_list, shop_gst="05AHYPG3732A1ZK", 
                          customer_name="", customer_gst="", customer_address="", service_type="NA", customer_code="5648", customer_category="MASS"):
    veh_clean = str(veh).strip().upper()
    if veh_clean == "" or veh_clean == "COUNTER SALE":
        return generate_counter_sale_html(inv_no, inv_date, items_list, shop_gst, customer_name, customer_gst, customer_address, customer_code, customer_category)
    else:
        return generate_jobcard_invoice_html(inv_no, inv_date, veh, items_list, shop_gst, customer_name, customer_gst, customer_address, service_type=service_type)

@st.cache_data(ttl=2)
def fetch_master_data():
    try:
        w = supabase.table("workshop_records").select("*").order("created_at", desc=True).execute().data
        b = supabase.table("workshop_billing").select("*").order("created_at", desc=True).execute().data
        a = supabase.table("staff_attendance").select("*").order("created_at", desc=True).execute().data
        i = supabase.table("shared_inventory").select("*").order("part_name").execute().data
        e = supabase.table("workshop_expenses").select("*").order("created_at", desc=True).execute().data
        s = supabase.table("staff_details").select("*").order("staff_name").execute().data
        return w, b, a, i, e, s
    except: return [], [], [], [], [], []

cars_data, bills_data, att_data, inv_data, exp_data, staff_list = fetch_master_data()

# Dynamic Staff Lists
mechanics = [s['staff_name'] for s in staff_list if s['role'] in ['Mechanic', 'Denter', 'Painter', 'Washing', 'Helper']]
service_advisors = [s['staff_name'] for s in staff_list if s['role'] in ['Advisor', 'Office']]
all_staff_names = [s['staff_name'] for s in staff_list]

# Map Inventory Data for Auto-Fetching
inventory_dict = {str(item['part_number']).strip().upper(): item for item in inv_data} if inv_data else {}

# ==========================================
# 3. SIDEBAR COMMAND CENTER 
# ==========================================
with st.sidebar:
    st.success(f"👤 Logged in as: **{st.session_state['user'].capitalize()}**")
    st.markdown("<h2 style='text-align: center; color: #8b0000;'>Command Center</h2>", unsafe_allow_html=True)
    
    today_str = datetime.now(IST).strftime('%d-%m-%Y')
    st.info(f"📅 **Date:** {today_str}\n\n🕒 **Time:** {datetime.now(IST).strftime('%I:%M %p')}")
    
    st.divider()
    active = len([c for c in cars_data if c['status'] != 'Delivered']) if cars_data else 0
    st.metric("🚗 Vehicles in Shop", active)
    
    today_rev = sum(float(b.get('amount_paid', 0)) for b in bills_data if not b.get('is_estimate') and get_ist(b['created_at'])[0] == today_str and b.get('payment_status') != 'Cancelled') if bills_data else 0
    st.metric("💰 Today's Collection", f"₹{today_rev:,.0f}")
    
    st.divider()
    c_btn1, c_btn2 = st.columns(2)
    if c_btn1.button("🔄 Refresh", use_container_width=True): st.rerun()
    if c_btn2.button("🚪 Logout", use_container_width=True):
        st.session_state['logged_in'] = False
        st.rerun()

# ==========================================
# 4. CORE TABS
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🚗 WORKSHOP & JOBCARD", "🧾 BILLING ENGINE", "💰 PAYROLL & SALARY", "👥 STAFF HQ", "📦 INVENTORY", "💸 EXPENSES & SLIPS", "📊 REPORTS & CRM"
])

# --- TAB 1: WORKSHOP MANAGER, CHECK-IN & JOBCARD UPDATER ---
with tab1:
    sub_tab_ws, sub_tab_hist, sub_tab_jc_print = st.tabs(["⚙️ Live Workshop Board", "🔍 Vehicle Service History", "📋 JobCard Viewer & Live Data Entry"])
    
    with sub_tab_ws:
        col_add, col_up = st.columns([1, 1])
        with col_add:
            st.write("### ➕ Register Entry & Create JobCard")
            with st.form("add_form", clear_on_submit=True):
                f1, f2 = st.columns(2)
                veh = f1.text_input("Vehicle Number").upper()
                name = f2.text_input("Customer Name")
                phone = st.text_input("Customer Phone")
                km_reading = st.number_input("Entry KM Reading", min_value=0, step=1)
                advisor = st.selectbox("Service Advisor", service_advisors if service_advisors else ["Admin"])
                service = st.selectbox("Service Type", ["PMS", "Running Repair", "Body Shop", "Washing Only"])
                mech = st.selectbox("Assign Primary Mechanic", mechanics if mechanics else ["Unassigned"])
                
                st.markdown("---")
                st.write("#### 📝 Initial Demanded Works")
                u_demands = st.text_area("Demanded Works / Client Complaints (Separate items with commas ',')")
                u_parts = st.text_area("Required Parts Checklist (Separate items with commas ',')")
                
                if st.form_submit_button("SAVE ENTRY & COMPILE JOBCARD"):
                    if veh and name:
                        payload = {
                            "customer_name": name, "phone_number": phone, "vehicle_number": veh, 
                            "entry_km": str(km_reading), "service_advisor": advisor, "service_type": service, 
                            "mechanic_name": mech, "status": "Queued"
                        }
                        try:
                            payload["customer_demands"] = u_demands
                            payload["requested_parts"] = u_parts
                            supabase.table("workshop_records").insert(payload).execute()
                        except:
                            payload["service_type"] = f"{service} [Demands: {u_demands} | Parts: {u_parts}]"
                            supabase.table("workshop_records").insert(payload).execute()
                            
                        st.success("Vehicle Entry Added & JobCard Created!"); time.sleep(1); st.rerun()

        with col_up:
            st.write("### 🔄 Update Status")
            if cars_data:
                active_cars = [f"{c['vehicle_number']} ({c['customer_name']})" for c in cars_data if c['status'] != 'Delivered']
                if active_cars:
                    car_to_up = st.selectbox("Select Active Car:", active_cars)
                    new_stat = st.radio("Current Status:", ["Queued", "In Workshop", "Washing", "Ready", "Delivered"], horizontal=True)
                    if st.button("CONFIRM STATUS CHANGE", type="primary"):
                        selected_id = cars_data[[f"{c['vehicle_number']} ({c['customer_name']})" for c in cars_data].index(car_to_up)]['id']
                        updates = {"status": new_stat}
                        if new_stat == "Delivered": updates["delivered_date"] = str(datetime.now(IST).strftime('%Y-%m-%d'))
                        supabase.table("workshop_records").update(updates).eq("id", selected_id).execute()
                        st.rerun()

        st.divider()
        st.write("### 📊 Live Service Board & WhatsApp Alerts")
        if cars_data:
            today_date_str = datetime.now(IST).strftime('%Y-%m-%d')
            for c in cars_data:
                e_date, e_time = get_ist(c['created_at'])
                
                if c['status'] == "Delivered":
                    d_date = c.get('delivered_date', '')
                    if d_date != today_date_str:
                        continue 
                
                c1, c2, c3 = st.columns([3, 1.2, 1.2])
                c1.write(f"**{c['vehicle_number']}** | {c['customer_name']} | KM: {c.get('entry_km', '0')}")
                c1.caption(f"📅 **Entry:** {e_date} {e_time} | SA: {c.get('service_advisor', 'N/A')} | 👨‍🔧 Mech: {c.get('mechanic_name', 'N/A')} | Status: **{c['status']}**")
                
                if c3.button("📋 Open JobCard", key=f"jc_board_{c['id']}"):
                    st.session_state['selected_jc_id'] = c['id']
                    st.success("JobCard loaded into the Entry & Printing tab!")
                
                c_phone = c.get('phone_number', '')
                if c['status'] == "Delivered":
                    msg = f"Thank you for choosing Shri Shyamji Auto Service Center, {c['customer_name']}! Your vehicle {c['vehicle_number']} is delivered. Safe travels!"
                    c2.link_button("🏁 Send 'Thank You'", f"https://wa.me/{c_phone}?text={urllib.parse.quote(msg)}")
                elif c['status'] == "Ready":
                    msg = f"Hello {c['customer_name']}! Your vehicle {c['vehicle_number']} is ready for pickup at Shri Shyamji Auto Service Center."
                    c2.link_button("🟢 Send 'Ready'", f"https://wa.me/{c_phone}?text={urllib.parse.quote(msg)}")
                else:
                    msg = f"Welcome to Shri Shyamji Auto Service Center! Your vehicle {c['vehicle_number']} is registered. Your Service Advisor is {c.get('service_advisor', 'N/A')}."
                    c2.link_button("📱 Send SA Details", f"https://wa.me/{c_phone}?text={urllib.parse.quote(msg)}")
                st.divider()

    with sub_tab_hist:
        st.write("### 🔍 Complete Vehicle Service History & Job Cards")
        search_veh = st.text_input("Enter Vehicle Number to Track History (e.g., UK04X1234):").strip().upper()
        
        if search_veh:
            matched_visits = [c for c in cars_data if c['vehicle_number'] == search_veh]
            matched_bills = [b for b in bills_data if b['vehicle_number'] == search_veh]
            
            if matched_visits:
                st.markdown(f"#### 📜 Summary Profile for **{search_veh}**")
                st.info(f"👤 **Latest Registered Customer:** {matched_visits[0]['customer_name']} | 📱 **Phone:** {matched_visits[0].get('phone_number', 'N/A')}")
                
                vh_col1, vh_col2 = st.columns(2)
                vh_col1.metric("Total Visits on Record", len(matched_visits))
                total_spent = sum(float(b['total_amount']) for b in matched_bills if not b.get('is_estimate') and b.get('payment_status') != 'Cancelled')
                vh_col2.metric("Total Revenue Contributed", f"₹{total_spent:,.2f}")
                
                st.write("#### 🛠️ Chronological Workshop Check-ins & Associated Job Cards")
                for item in matched_visits:
                    v_date, v_time = get_ist(item['created_at'])
                    with st.expander(f"📅 Visit Date: {v_date} (Status: {item['status']}) - Click to Expand Job Card"):
                        st.markdown(f"### 📋 Job Card Details (JC-{item['id']})")
                        st.markdown(f"""
                        * **Service Type:** {item.get('service_type', 'N/A')}
                        * **Odometer (KM) Reading:** {item.get('entry_km', 'N/A')}
                        * **Assigned Mechanic:** {item.get('mechanic_name', 'N/A')}
                        * **Service Advisor:** {item.get('service_advisor', 'N/A')}
                        """)
                        
                        vch_col1, vch_col2 = st.columns(2)
                        with vch_col1:
                            st.info("🗣️ **Customer Demanded Works / Issues:**")
                            demands_list = item.get('customer_demands', '')
                            if demands_list:
                                for idx, d in enumerate(demands_list.split(',')):
                                    if d.strip(): st.write(f"{idx+1}. {d.strip()}")
                            else:
                                st.write("*No complaints recorded.*")
                                
                        with vch_col2:
                            st.success("📦 **Required / Demanded Parts Checklist:**")
                            parts_list = item.get('requested_parts', '')
                            if parts_list:
                                for idx, p in enumerate(parts_list.split(',')):
                                    if p.strip(): st.write(f"{idx+1}. {p.strip()}")
                            else:
                                st.write("*No dynamic parts checklist recorded.*")
                
                st.write("#### 🧾 Invoices Associated with Vehicle")
                if matched_bills:
                    for bill in matched_bills:
                        b_date, _ = get_ist(bill['created_at'])
                        b_status_str = bill.get('payment_status', 'Active')
                        b_title = "📋 Estimate" if bill.get('is_estimate') else f"💰 Tax Invoice ({b_status_str})"
                        with st.expander(f"{b_title} | No: {bill['invoice_number']} | Date: {b_date} | Total: ₹{float(bill['total_amount']):,.2f}"):
                            st.write(f"**Amount Paid:** ₹{float(bill['amount_paid']):,.2f} | **Balance Due:** ₹{float(bill['total_amount'])-float(bill['amount_paid']):,.2f}")
                            if bill.get('invoice_details'):
                                try:
                                    inv_items = json.loads(bill['invoice_details'])
                                    if isinstance(inv_items, dict) and "items" in inv_items:
                                        st.dataframe(pd.DataFrame(inv_items["items"]), use_container_width=True, hide_index=True)
                                    else:
                                        st.dataframe(pd.DataFrame(inv_items), use_container_width=True, hide_index=True)
                                except: st.caption("No dynamic parts breakdown found.")
                else: st.warning("No invoices generated for this vehicle yet.")
            else: st.error("No service history found for this vehicle number.")

    with sub_tab_jc_print:
        st.write("### 📋 Live Data Entry & Update JobCards")
        active_jc_options = {c['id']: f"{c['vehicle_number']} - {c['customer_name']} (JC-{c['id']})" for c in cars_data if c['status'] != 'Delivered'}
        
        if active_jc_options:
            default_index = 0
            if 'selected_jc_id' in st.session_state and st.session_state['selected_jc_id'] in active_jc_options:
                default_index = list(active_jc_options.keys()).index(st.session_state['selected_jc_id'])
                
            selected_id_key = st.selectbox("Select Active Vehicle to View or Enter New Work Data:", options=list(active_jc_options.keys()), format_func=lambda x: active_jc_options[x], index=default_index)
            jc_tgt = next((x for x in cars_data if x['id'] == selected_id_key), None)
            
            if jc_tgt:
                st.markdown("---")
                st.write("#### ✍️ Enter / Modify Job Card Information")
                
                initial_demands_val = jc_tgt.get('customer_demands', '') if jc_tgt.get('customer_demands') else ''
                initial_parts_val = jc_tgt.get('requested_parts', '') if jc_tgt.get('requested_parts') else ''
                
                updated_demands = st.text_area("Customer Demanded Works / Client Complaints:", value=initial_demands_val, key="jc_live_demands_text")
                updated_parts = st.text_area("Required Parts / Materials Checklist:", value=initial_parts_val, key="jc_live_parts_text")
                
                if st.button("💾 UPDATE & SAVE JOBCARD DATA", type="primary", key="jc_save_action_trigger"):
                    try:
                        supabase.table("workshop_records").update({
                            "customer_demands": updated_demands,
                            "requested_parts": updated_parts
                        }).eq("id", jc_tgt['id']).execute()
                        st.success("✅ Job Card entries successfully synced to cloud database!")
                    except:
                        appended_notes = f"{jc_tgt.get('service_type','Running Repair')} [Updated Demands: {updated_demands} | Parts: {updated_parts}]"
                        supabase.table("workshop_records").update({"service_type": appended_notes}).eq("id", jc_tgt['id']).execute()
                        st.success("✅ Job Card entries compiled into service notes profile successfully!")
                        
                    time.sleep(0.5)
                    st.rerun()
                
                st.markdown("---")
                st.write("#### 🖨️ Printable Document Preview Sheet")
                jc_date_str, jc_time_str = get_ist(jc_tgt['created_at'])
                
                compiled_jc_html = generate_jobcard_html(
                    jc_no=f"JC-{jc_tgt['id']}", jc_date=f"{jc_date_str} {jc_time_str}", 
                    veh=jc_tgt['vehicle_number'], name=jc_tgt['customer_name'], 
                    phone=jc_tgt.get('phone_number','N/A'), km=jc_tgt.get('entry_km','0'), 
                    advisor=jc_tgt.get('service_advisor','N/A'), mech=jc_tgt.get('mechanic_name','N/A'), 
                    jc_type=jc_tgt.get('service_type','Running Repair'), 
                    demands=updated_demands, parts=updated_parts
                )
                
                st.download_button(
                    label="📥 DOWNLOAD PRINTABLE JOB CARD",
                    data=compiled_jc_html,
                    file_name=f"JobCard_{jc_tgt['vehicle_number']}.html",
                    mime="text/html",
                    use_container_width=True
                )
                components.html(compiled_jc_html, height=550, scrolling=True)
        else:
            st.info("No active vehicles in the workshop currently. Register a car first to start entry logging.")

# --- TAB 2: BILLING ENGINE ---
with tab2:
    st.write("### 🧾 Invoice Generation & History Engine")
    
    # Initialize State SAFELY
    if "bill_items_df" not in st.session_state:
        st.session_state["bill_items_df"] = [
            {"Type": "Part", "Part_Number": "", "Description": "", "HSN": "8708", "Qty": 1.0, "Rate": 0.0, "Counter_Sale": False}
        ]
    if "edit_mode" not in st.session_state: st.session_state["edit_mode"] = False
    if "target_db_id" not in st.session_state: st.session_state["target_db_id"] = None
    if "current_inv_no" not in st.session_state: st.session_state["current_inv_no"] = ""

    b_mode1, b_mode2 = st.tabs(["⚙️ Create / Edit Document", "🔍 Search, Print & Receipts"])
    
    with b_mode1:
        if st.session_state["edit_mode"]:
            st.warning(f"⚠️ **EDIT MODE ACTIVE:** Modifying Invoice **{st.session_state['current_inv_no']}**")
            if st.button("❌ Cancel Edit & Start Fresh"):
                st.session_state["edit_mode"] = False
                st.session_state["target_db_id"] = None
                st.session_state["current_inv_no"] = ""
                st.session_state["bill_items_df"] = [{"Type": "Part", "Part_Number": "", "Description": "", "HSN": "8708", "Qty": 1.0, "Rate": 0.0, "Counter_Sale": False}]
                st.rerun()

        # Form Inputs
        c1, c2 = st.columns(2)
        b_veh = c1.text_input("Vehicle Number (Leave blank for Counter Sale)", value=st.session_state.get("bill_vnum", ""), key="bill_vnum_input").upper()
        b_type = c2.selectbox("Document Type", ["Tax Invoice", "Estimate", "Pre-Invoice"], key="bill_dtype_input")
        
        c3, c4 = st.columns(2)
        c_name = c3.text_input("Customer Name", value=st.session_state.get("bill_cname", ""), key="bill_cname_input")
        c_gst = c4.text_input("Customer GSTIN", value=st.session_state.get("bill_cgst", ""), key="bill_cgst_input")
        
        c5, c6 = st.columns(2)
        w_gst = c5.text_input("Workshop GSTIN", value="05AHYPG3732A1ZK", key="bill_wgst_input")
        c_addr = c6.text_input("Customer Address", value=st.session_state.get("bill_caddr", ""), key="bill_caddr_input")
        
        # --- CONDITIONAL COUNTER SALE DETAILS ---
        is_counter_sale = not b_veh.strip() or b_veh.strip() == "COUNTER SALE"
        c_code = "5648"
        c_cat = "MASS"
        if is_counter_sale:
            st.markdown("---")
            st.write("#### 📝 Counter Sale Additional Details")
            cc1, cc2 = st.columns(2)
            c_code = cc1.text_input("Customer Code", value=st.session_state.get("bill_ccode", "5648"), key="ccode_input")
            c_cat = cc2.text_input("Customer Category", value=st.session_state.get("bill_ccat", "MASS"), key="ccat_input")

        st.markdown("---")
        st.write("#### 🛒 Itemized Parts & Labor")

        edited_items = st.data_editor(
            st.session_state["bill_items_df"], 
            num_rows="dynamic", 
            use_container_width=True, 
            key="billing_editor",
            column_config={
                "Type": st.column_config.SelectboxColumn("Type", options=["Part", "Labor"]),
                "Part_Number": st.column_config.TextColumn("Part No. (Type & Press Enter)"),
                "Description": st.column_config.TextColumn("Description"),
                "HSN": st.column_config.TextColumn("HSN"),
                "Qty": st.column_config.NumberColumn("Qty", min_value=0.1),
                "Rate": st.column_config.NumberColumn("Rate (₹)", format="%.2f"),
                "Counter_Sale": st.column_config.CheckboxColumn("Counter Sale?") 
            }
        )
        
        # Post-Process Auto-Fetch (Rock Solid, prevents random wipes and forces Part_Number to be string)
        if edited_items is not None:
            current_items = edited_items.to_dict('records') if isinstance(edited_items, pd.DataFrame) else edited_items
            needs_rerun = False
            for item in current_items:
                raw_pn = item.get("Part_Number")
                pn = str(raw_pn).strip().upper() if pd.notna(raw_pn) and raw_pn is not None else ""
                if pn in ["NONE", "NAN", "NULL"]: pn = ""
                item["Part_Number"] = pn  
                
                raw_desc = item.get("Description")
                desc = str(raw_desc).strip() if pd.notna(raw_desc) and raw_desc is not None else ""
                if desc in ["None", "nan", "NaN"]: desc = ""
                item["Description"] = desc
                
                if pn and not desc:
                    if pn in inventory_dict:
                        match = inventory_dict[pn]
                        item["Description"] = match.get("part_name", "")
                        item["Rate"] = float(match.get("selling_price", 0.0))
                        item["HSN"] = "8708"
                        item["Qty"] = 1.0
                        needs_rerun = True
            
            st.session_state["bill_items_df"] = current_items
            if needs_rerun:
                st.rerun()

        # Totals Calculation
        current_items = st.session_state["bill_items_df"]
        p_recv = st.number_input("Amount Received Now (₹)", min_value=0.0, value=float(st.session_state.get("p_recv", 0.0)), key="p_recv_input")
        
        parts_sub = sum(safe_float(r.get('Qty')) * safe_float(r.get('Rate')) for r in current_items if r.get('Type') == "Part")
        labor_sub = sum(safe_float(r.get('Qty')) * safe_float(r.get('Rate')) for r in current_items if r.get('Type') == "Labor")
        
        gst_amount = (parts_sub + labor_sub) * 0.18 if b_type == "Tax Invoice" else 0.0
        grand_total = round(parts_sub + labor_sub + gst_amount)

        st.markdown(f"### 💰 Net Total (Incl. GST): ₹{grand_total:,.2f}")
        
        btn_label = "💾 UPDATE INVOICE" if st.session_state["edit_mode"] else "🚀 GENERATE INVOICE"
        
        if st.button(btn_label, type="primary", use_container_width=True):
            
            if st.session_state["edit_mode"]:
                inv_no = st.session_state["current_inv_no"]
            else:
                now = datetime.now(IST)
                fy_str = f"{str(now.year - 1)[-2:]}/{str(now.year)[-2:]}" if now.month < 4 else f"{str(now.year)[-2:]}/{str(now.year + 1)[-2:]}"
                base_prefix = "SS" if b_type == "Tax Invoice" else ("EST" if b_type == "Estimate" else "PRE")
                full_prefix = f"{base_prefix}{fy_str}-A"
                
                existing_nums = []
                if bills_data:
                    for b in bills_data:
                        inv = b.get('invoice_number', '')
                        if inv.startswith(full_prefix):
                            try: existing_nums.append(int(inv.split('-A')[-1]))
                            except: pass
                next_num = max(existing_nums) + 1 if existing_nums else 1
                inv_no = f"{full_prefix}{str(next_num).zfill(4)}"
            
            final_veh = b_veh if b_veh.strip() else "COUNTER SALE"
            
            # Save Counter Sale details securely inside JSON
            if is_counter_sale:
                details_to_save = {"items": current_items, "c_code": c_code, "c_cat": c_cat}
            else:
                details_to_save = current_items
                
            payload = {
                "invoice_number": inv_no, "vehicle_number": final_veh, "customer_name": c_name,
                "customer_gst": c_gst, "customer_address": c_addr,
                "total_amount": grand_total, "amount_paid": p_recv, 
                "invoice_details": json.dumps(details_to_save), "payment_status": "Active"
            }

            if st.session_state["edit_mode"]:
                supabase.table("workshop_billing").update(payload).eq("id", st.session_state["target_db_id"]).execute()
                st.toast(f"✅ Invoice {inv_no} updated!")
                st.session_state["edit_mode"] = False
            else:
                supabase.table("workshop_billing").insert(payload).execute()
                st.toast(f"✅ Invoice {inv_no} generated!")
            
            time.sleep(1)
            st.rerun()

    # --- SEARCH, EDIT & PRINT MODE ---
    with b_mode2:
        st.write("#### 🔎 Search Past Invoices")
        search_inv = st.text_input("Enter Invoice Number (e.g., SS26/27-A0001):").strip().upper()
        
        if st.button("🔍 FETCH INVOICE"):
            match = next((b for b in bills_data if b['invoice_number'] == search_inv), None)
            if match:
                st.session_state["fetched_invoice"] = match
                st.success(f"Invoice {search_inv} found!")
            else:
                st.error("Invoice not found.")
        
        if "fetched_invoice" in st.session_state:
            match = st.session_state["fetched_invoice"]
            st.json({
                "Invoice No": match['invoice_number'],
                "Customer": match.get('customer_name', ''),
                "Vehicle": match.get('vehicle_number', ''),
                "Total Amount": match.get('total_amount', 0),
                "Amount Paid": match.get('amount_paid', 0),
                "Status": match.get('payment_status', 'Active')
            })

            c_ed, c_pr = st.columns(2)
            
            # 1. EDIT BUTTON
            if c_ed.button("✏️ LOAD FOR EDITING", use_container_width=True):
                st.session_state["edit_mode"] = True
                st.session_state["target_db_id"] = match['id']
                st.session_state["current_inv_no"] = match['invoice_number']
                
                saved_veh = match.get('vehicle_number', '')
                st.session_state["bill_vnum"] = "" if saved_veh == "COUNTER SALE" else saved_veh
                
                st.session_state["bill_cname"] = match.get('customer_name', '')
                st.session_state["bill_cgst"] = match.get('customer_gst', '')
                st.session_state["bill_caddr"] = match.get('customer_address', '')
                st.session_state["p_recv"] = float(match.get('amount_paid', 0.0))
                try:
                    parsed_json = json.loads(match['invoice_details'])
                    if isinstance(parsed_json, dict) and "items" in parsed_json:
                        st.session_state["bill_items_df"] = parsed_json["items"]
                        st.session_state["bill_ccode"] = parsed_json.get("c_code", "5648")
                        st.session_state["bill_ccat"] = parsed_json.get("c_cat", "MASS")
                    else:
                        st.session_state["bill_items_df"] = parsed_json
                        st.session_state["bill_ccode"] = "5648"
                        st.session_state["bill_ccat"] = "MASS"
                except: pass
                st.rerun()
            
            # 2. PRINT INVOICE BUTTON
            parsed_json = json.loads(match.get('invoice_details', '[]'))
            if isinstance(parsed_json, dict) and "items" in parsed_json:
                items_list_for_print = parsed_json["items"]
                c_code_print = parsed_json.get("c_code", "5648")
                c_cat_print = parsed_json.get("c_cat", "MASS")
            else:
                items_list_for_print = parsed_json
                c_code_print = "5648"
                c_cat_print = "MASS"
                
            inv_html = generate_master_invoice_html(
                inv_no=match['invoice_number'], 
                inv_date=str(match['created_at'])[:10], 
                veh=match.get('vehicle_number', ''), 
                items_list=items_list_for_print, 
                shop_gst="05AHYPG3732A1ZK", 
                customer_name=match.get('customer_name', ''), 
                customer_gst=match.get('customer_gst', ''), 
                customer_address=match.get('customer_address', ''),
                service_type="BODY REPAIR" if "BODY" in match.get('service_type', '').upper() else "PMS / REPAIR",
                customer_code=c_code_print,
                customer_category=c_cat_print
            )
            c_pr.download_button("🖨️ DOWNLOAD INVOICE PDF/HTML", data=inv_html, file_name=f"{match['invoice_number']}.html", mime="text/html", use_container_width=True)
            
            # 3. CUSTOM RECEIPT GENERATOR
            st.markdown("---")
            st.write("#### 🧾 Generate Cash Receipt")
            r_col1, r_col2 = st.columns(2)
            rcpt_paid_by = r_col1.text_input("Amount Received From:", value=match.get('customer_name', ''))
            rcpt_auth = r_col2.text_input("Authorized Signatory (Owner):", value="Shri Shyamji Auto Service Center")
            
            rcpt_html = generate_receipt_html(
                receipt_date=datetime.now(IST).strftime('%d-%m-%Y'),
                inv_no=match['invoice_number'], 
                veh=match.get('vehicle_number', ''),
                customer_name=rcpt_paid_by, 
                total_amount=float(match['total_amount']),
                paid_amount=float(match.get('amount_paid', 0)),
                shop_gst="05AHYPG3732A1ZK"
            )
            st.download_button("🧾 DOWNLOAD CASH RECEIPT", data=rcpt_html, file_name=f"Receipt_{match['invoice_number']}.html", mime="text/html", use_container_width=True)

# --- TAB 3: PAYROLL & SALARY ---
with tab3:
    st.write("### 💰 Smart Payroll & Deductions Engine")
    
    sub1, sub2, sub3 = st.tabs(["⚙️ Monthly Calculator", "📜 Staff Ledger History", "✍️ Manual Entry"])
    
    with sub1:
        c_m1, c_y1, c_m2 = st.columns([1, 1, 2])
        month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        prev_month_idx = (datetime.now().month - 2) % 12 
        sel_month_name = c_m1.selectbox("Evaluating Month", month_names, index=prev_month_idx)
        sel_month_num = str(month_names.index(sel_month_name) + 1).zfill(2)
        sel_year = c_y1.number_input("Evaluating Year", min_value=2020, max_value=2100, value=datetime.now().year)
        special_holidays = c_m2.multiselect("Special Bonus Holidays (Ignored)", [f"{i:02d}-{sel_month_num}-{sel_year}" for i in range(1, 32)])

        if staff_list:
            salary_data = []
            for s in staff_list:
                s_name, s_base, s_type = s['staff_name'], s['base_salary'], s.get('pay_type', 'Monthly')
                s_att = [a for a in att_data if a['mechanic_name'] == s_name and f"-{sel_month_num}-{sel_year}" in a['date']]
                
                full_leaves = len([a for a in s_att if a['status'] == "Absent" and a['date'] not in special_holidays])
                half_leaves = len([a for a in s_att if a['status'] == "Half-Day" and a['date'] not in special_holidays])
                total_leaves = full_leaves + (half_leaves * 0.5)
                
                target_str = f"{sel_month_name} {sel_year}"
                adv = sum(float(e['amount']) for e in exp_data if e['category'] == "Staff Salary/Advance" and s_name.lower() in e['description'].lower() and target_str in e['description'])
                salary_data.append({"Name": s_name, "Pay Cycle": s_type, "Base (₹)": s_base, "Period (Days)": (7 if s_type == "Weekly (Hafta)" else 30), "Leaves Taken": total_leaves, "Advances Taken": adv, "Incentive / Overtime (₹)": 0.0, "Deductions (₹)": 0.0, "Deduction Reason": ""})

            df_sal = pd.DataFrame(salary_data)
            edited_sal = st.data_editor(df_sal, use_container_width=True, hide_index=True)
            
            for index, row in edited_sal.iterrows():
                period, leaves, base = float(row['Period (Days)']), float(row['Leaves Taken']), float(row['Base (₹)'])
                extra_payable = 4 if leaves < 7 else 0
                
                payable_days = period - leaves + extra_payable
                gross = (base / period) * payable_days if period > 0 else 0
                net = gross - float(row['Advances Taken']) + float(row['Incentive / Overtime (₹)']) - float(row['Deductions (₹)'])
                
                c1, c2, c3, c4 = st.columns([2, 1.2, 1.2, 2])
                c1.info(f"👤 **{row['Name']}**")
                c1.caption(f"📅 Leaves: {leaves} | Addl. Holiday: {extra_payable}")
                
                c2.metric("Gross", f"₹{gross:,.0f}")
                c3.metric("Net Due", f"₹{max(net, 0):,.0f}")
                
                with c4:
                    with st.expander("💳 Payment Options"):
                        p_date = st.date_input("Date", key=f"d_{row['Name']}")
                        p_mode = st.selectbox("Mode", ["Cash", "Online"], key=f"m_{row['Name']}")
                        paid_by = st.text_input("Paid By Whom", value=st.session_state['user'].capitalize(), key=f"by_{row['Name']}")
                        pay_amt = st.number_input("Amount to Pay (₹)", min_value=0.0, max_value=float(max(net, 50000)), value=float(max(net, 0)), key=f"amt_{row['Name']}")
                        
                        if st.button("✅ Log Payment", key=f"log_{row['Name']}"):
                            desc = f"Salary: {row['Name']} | {sel_month_name} {sel_year} | Mode: {p_mode} | Paid By: {paid_by} (Auth-Slip System)"
                            supabase.table("workshop_expenses").insert({"date": p_date.strftime('%d-%m-%Y'), "category": "Staff Salary/Advance", "amount": pay_amt, "description": desc, "authorized_by": paid_by}).execute()
                            st.toast("✅ Payment Logged!"); time.sleep(0.5); st.rerun()

    with sub2:
        st.write("### 📜 Staff Payment Ledger")
        sel_staff = st.selectbox("Select Staff Member to view history:", all_staff_names)
        staff_expenses = [e for e in exp_data if e['category'] == "Staff Salary/Advance" and sel_staff.lower() in e['description'].lower()]
        if staff_expenses:
            df_hist = pd.DataFrame(staff_expenses)[['date', 'amount', 'description']]
            df_hist.rename(columns={'date': 'Date', 'amount': 'Paid (₹)', 'description': 'Details'}, inplace=True)
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
            st.metric("Total Paid to Date", f"₹{df_hist['Paid (₹)'].sum():,.2f}")
            
    with sub3:
        st.write("### ✍️ Manual Entry")
        st.info("Log payments manually if you missed them in the calculator.")
           
# --- TAB 4: STAFF HQ ---
with tab4:
    st.write("### 👥 Human Resources & Staff Management")
    c1, c2 = st.columns([1, 2])
    with c1:
        st.write("#### ➕ Recruit New Staff")
        with st.form("recruit_form", clear_on_submit=True):
            new_name = st.text_input("Full Name")
            new_role = st.selectbox("Role", ["Mechanic", "Advisor", "Denter", "Painter", "Office", "Washing", "Helper"])
            new_pay_type = st.radio("Pay Frequency", ["Monthly", "Weekly (Hafta)"], horizontal=True)
            new_sal = st.number_input("Base Salary (per period)", min_value=0, value=15000)
            if st.form_submit_button("ADD TO TEAM"):
                if new_name:
                    supabase.table("staff_details").insert({"staff_name": new_name, "role": new_role, "base_salary": new_sal, "pay_type": new_pay_type}).execute()
                    st.success(f"{new_name} Joined!"); time.sleep(1); st.rerun()

    with c2:
        st.write("#### 📋 Active Staff & Base Salaries")
        if staff_list:
            df_staff = pd.DataFrame(staff_list)
            edited_staff = st.data_editor(df_staff[['id', 'staff_name', 'role', 'pay_type', 'base_salary']], hide_index=True, use_container_width=True)
            if st.button("💾 SAVE CHANGES"):
                for _, row in edited_staff.iterrows():
                    supabase.table("staff_details").update({"base_salary": row['base_salary'], "pay_type": row['pay_type']}).eq("id", row['id']).execute()
                st.success("Updated!"); time.sleep(1); st.rerun()
                
    st.divider()
    st.write("### 📅 Daily Attendance Logger")
    with st.form("attendance_form"):
        staff_member = st.selectbox("Select Staff", all_staff_names)
        attendance_status = st.radio("Status", ["Present", "Absent", "Half-Day"], horizontal=True)
        if st.form_submit_button("LOG ATTENDANCE"):
            supabase.table("staff_attendance").insert({"mechanic_name": staff_member, "date": datetime.now(IST).strftime('%d-%m-%Y'), "status": attendance_status}).execute()
            st.success(f"Logged for {staff_member}"); st.rerun()

# --- TAB 5: INVENTORY ---
with tab5:
    st.write("### 📦 Central Master Inventory")
    with st.expander("➕ Inward Stock Form (Smart Merge)", expanded=False):
        with st.form("inv_form", clear_on_submit=True):
            c_name, c_num, c_brand, c_qty, c_price = st.columns([2, 1, 1, 1, 1])
            p_name = c_name.text_input("Part Name")
            p_num = c_num.text_input("Part Number (Exact Match)")
            p_brand = c_brand.selectbox("Type / Brand", ["Maruti", "Non Maruti"])
            p_qty = c_qty.number_input("Qty", min_value=0)
            p_price = c_price.number_input("Price (₹)", min_value=0.0)
            if st.form_submit_button("ADD TO CLOUD"):
                check_db = supabase.table("shared_inventory").select("*").eq("part_number", p_num.strip().upper()).execute()
                if check_db.data: 
                    supabase.table("shared_inventory").update({"stock_qty": int(check_db.data[0].get('stock_qty', 0)) + int(p_qty), "selling_price": p_price, "part_brand": p_brand}).eq("id", check_db.data[0]['id']).execute()
                else: 
                    supabase.table("shared_inventory").insert({"part_name": p_name, "part_number": p_num.strip().upper(), "stock_qty": p_qty, "selling_price": p_price, "part_brand": p_brand}).execute()
                st.rerun()

    if inv_data:
        f_c1, f_c2 = st.columns(2)
        search_part = f_c1.text_input("🔍 Search Live Catalog:")
        brand_filter = f_c2.selectbox("Filter by Type:", ["All Types", "Maruti", "Non Maruti"])
        
        df_inv = pd.DataFrame(inv_data)
        if search_part: 
            df_inv = df_inv[df_inv['part_name'].str.contains(search_part, case=False, na=False) | df_inv['part_number'].str.contains(search_part, case=False, na=False)]
        if brand_filter != "All Types":
            df_inv = df_inv[df_inv['part_brand'] == brand_filter]
            
        st.dataframe(df_inv[['part_number', 'part_name', 'part_brand', 'stock_qty', 'selling_price']], use_container_width=True, hide_index=True)

# --- TAB 6: MASTER EXPENSE LEDGER & PAY SLIP ENGINE ---
with tab6:
    if 'current_slip_html' not in st.session_state: st.session_state['current_slip_html'] = None
    if 'current_slip_id' not in st.session_state: st.session_state['current_slip_id'] = ""

    col_log, col_view = st.columns([1, 1.8])
    with col_log:
        st.write("### 💸 Log New Expense")
        with st.form("exp_form_ledger", clear_on_submit=True):
            e_cat = st.selectbox("Category", ["Spare Parts Purchase", "Staff Salary/Advance", "Electricity & Utilities", "Tea, Snacks & Meals", "Workshop Maintenance", "Rent", "Other"])
            e_amt = st.number_input("Amount (₹)", min_value=0.0)
            e_desc = st.text_input("Notes / Description")
            e_auth = st.text_input("Authorized/Paid By Whom", value=st.session_state['user'].capitalize())
            e_date = st.date_input("Date of Expense")
            
            if st.form_submit_button("LOG EXPENSE & COMPILE SLIP"):
                if e_amt > 0:
                    res = supabase.table("workshop_expenses").insert({
                        "date": e_date.strftime('%d-%m-%Y'), "category": e_cat, 
                        "amount": e_amt, "description": e_desc, "authorized_by": e_auth
                    }).execute()
                    
                    new_id = res.data[0]['id'] if res.data else datetime.now().strftime('%M%S')
                    st.session_state['current_slip_id'] = f"EXPS-{new_id}"
                    st.session_state['current_slip_html'] = generate_expense_slip_html(
                        slip_id=f"EXPS-{new_id}", date=e_date.strftime('%d-%m-%Y'),
                        cat=e_cat, amt=e_amt, desc=e_desc, paid_by=e_auth
                    )
                    st.toast("✅ Expense Logged Natively!"); time.sleep(0.5); st.rerun()

    with col_view:
        st.write("### 📜 Master Expense Ledger")
        if exp_data:
            df_exp = pd.DataFrame(exp_data)
            f1, f2 = st.columns(2)
            month_filter = f1.selectbox("Filter by Month", ["All Time", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], index=0)
            unique_cats = df_exp['category'].dropna().unique().tolist()
            cat_filter = f2.selectbox("Filter by Category", ["All Categories"] + unique_cats)
            
            filtered_exp = df_exp.copy()
            if month_filter != "All Time": filtered_exp = filtered_exp[filtered_exp['date'].str.contains(f"-{month_filter}-", na=False)]
            if cat_filter != "All Categories": filtered_exp = filtered_exp[filtered_exp['category'] == cat_filter]
            
            total_filtered = filtered_exp['amount'].astype(float).sum()
            st.markdown(f"<div style='background-color:#ffebee; padding:15px; border-radius:8px; border-left:5px solid #d32f2f; margin-bottom:10px;'><h4>Total Expenses: ₹{total_filtered:,.2f}</h4></div>", unsafe_allow_html=True)
            
            st.write("#### 🖨️ Generate Slip for Previous Expenses")
            expense_titles = [f"ID: {row['id']} | {row['date']} | {row['category']} | ₹{row['amount']}" for _, row in filtered_exp.iterrows()]
            selected_exp_row = st.selectbox("Choose a record below to display payment slip:", ["-- Select Record --"] + expense_titles)
            
            if selected_exp_row != "-- Select Record --":
                tgt_id = int(selected_exp_row.split("|")[0].replace("ID:", "").strip())
                match_exp = next((x for x in exp_data if x['id'] == tgt_id), None)
                if match_exp:
                    st.session_state['current_slip_id'] = f"EXPS-{match_exp['id']}"
                    st.session_state['current_slip_html'] = generate_expense_slip_html(
                        slip_id=f"EXPS-{match_exp['id']}", date=match_exp['date'],
                        cat=match_exp['category'], amt=match_exp['amount'], 
                        desc=match_exp.get('description',''), paid_by=match_exp.get('authorized_by','Admin')
                    )

    if st.session_state['current_slip_html'] is not None:
        st.markdown("---")
        st.write("### 🖨️ Active Payment Voucher")
        sc1, sc2 = st.columns([1, 2])
        sc1.download_button(
            label="📥 DOWNLOAD PRINTABLE EXPENSE SLIP",
            data=st.session_state['current_slip_html'],
            file_name=f"ExpenseSlip_{st.session_state['current_slip_id']}.html",
            mime="text/html",
            use_container_width=True
        )
        if sc2.button("🧹 Clear Voucher Workspace View"):
            st.session_state['current_slip_html'] = None
            st.session_state['current_slip_id'] = ""
            st.rerun()
        components.html(st.session_state['current_slip_html'], height=350, scrolling=True)

# --- TAB 7: REPORTS & CRM ---
with tab7:
    st.write("### 📊 Business P&L")
    total_col = sum(float(b.get('amount_paid', 0)) for b in bills_data if not b.get('is_estimate') and b.get('payment_status') != 'Cancelled')
    total_exp = sum(float(e.get('amount', 0)) for e in exp_data)
    m1, m2, m3 = st.columns(3)
    m1.metric("Gross Revenue", f"₹{total_col:,.2f}")
    m2.metric("Total Expenses", f"₹{total_exp:,.2f}")
    m3.metric("Net Profit", f"₹{total_col - total_exp:,.2f}")
    
    st.divider()
    st.write("#### 📥 Secure Excel Downloads")
    c_dl1, c_dl2, c_dl3 = st.columns(3)
    if cars_data: c_dl1.download_button("🚗 Workshop Flow", data=to_excel(pd.DataFrame(cars_data)), file_name="Workshop_Ledger.xlsx", use_container_width=True)
    if bills_data: c_dl2.download_button("💰 Billing Ledger", data=to_excel(pd.DataFrame(bills_data)), file_name="Billing_Ledger.xlsx", use_container_width=True)
    if exp_data: c_dl3.download_button("💸 Expense Log", data=to_excel(pd.DataFrame(exp_data)), file_name="Expense_Ledger.xlsx", use_container_width=True)
    
    st.divider()
    c_fb, c_ins = st.columns(2)
    with c_fb:
        st.write("### ⭐ 2-Day Feedback Desk")
        if cars_data:
            target_date = str((datetime.now(IST) - timedelta(days=2)).date())
            feedback_list = [c for c in cars_data if c['status'] == "Delivered" and c.get('delivered_date') == target_date and c.get('feedback_sent') != 'Yes']
            for f in feedback_list:
                st.warning(f"🚘 **{f['vehicle_number']}** ({f['customer_name']})")
                fb_msg = f"Dear {f['customer_name']}, it's been 2 days since your vehicle {f['vehicle_number']} was serviced at Shri Shyamji Auto Service Center. We hope it's running smoothly!"
                st.link_button("📱 Send WhatsApp Feedback", f"https://wa.me/{f.get('phone_number', '')}?text={urllib.parse.quote(fb_msg)}")
    with c_ins:
        st.write("### 🛡️ Insurance Alerts (Expires within 45 Days)")
        if cars_data:
            today_date_obj = datetime.now(IST).date()
            for i in [c for c in cars_data if c.get('insurance_expiry') and c['insurance_expiry'].strip() != ""]:
                try:
                    exp_obj = datetime.strptime(i['insurance_expiry'].strip(), '%d-%m-%Y').date()
                    days_left = (exp_obj - today_date_obj).days
                    if 0 <= days_left <= 45:
                        st.warning(f"⏳ **DUE IN {days_left} DAYS:** {i['vehicle_number']} ({i['customer_name']})")
                        ins_msg = f"Hello {i['customer_name']}, your vehicle ({i['vehicle_number']}) insurance expires on {i['insurance_expiry']}. Let us know if we can help you renew it."
                        st.link_button("📱 Send Reminder", f"https://wa.me/{i.get('phone_number', '')}?text={urllib.parse.quote(ins_msg)}")
                    elif days_left < 0:
                        st.error(f"🚨 **EXPIRED:** {i['vehicle_number']} ({i['customer_name']})")
                except: pass
