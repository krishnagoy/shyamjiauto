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
    .stApp { background-color: #f0f2f6; }
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
        <p style="margin:5px 0 0 0; font-size: 16px; opacity: 0.9;">Master ERP, Payroll & CRM System</p>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 2. DATABASE SETUP & HTML GENERATORS
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
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def generate_jobcard_html(jc_no, jc_date, veh, name, phone, km, advisor, mech, stype, demands, parts):
    demand_rows = "".join([f"<tr><td style='border: 1px solid #ddd; padding: 8px;'>{idx+1}</td><td style='border: 1px solid #ddd; padding: 8px;'>{d.strip()}</td><td style='border: 1px solid #ddd; padding: 8px; color: green; font-weight:bold;'>Assigned</td></tr>" for idx, d in enumerate(demands.split(',')) if d.strip()]) if demands.strip() else "<tr><td colspan='3' style='border: 1px solid #ddd; padding: 8px; text-align:center; color:#777;'>No structural demanding works cataloged yet.</td></tr>"
    parts_rows = "".join([f"<tr><td style='border: 1px solid #ddd; padding: 8px;'>{idx+1}</td><td style='border: 1px solid #ddd; padding: 8px;'>{p.strip()}</td></tr>" for idx, p in enumerate(parts.split(',')) if p.strip()]) if parts.strip() else "<tr><td colspan='2' style='border: 1px solid #ddd; padding: 8px; text-align:center; color:#777;'>No structural initial parts checklist cataloged yet.</td></tr>"
    
    return f"""
    <html>
    <body style="font-family: sans-serif; color: #333; padding: 20px; border: 1px solid #ccc; max-width: 800px; margin: auto;">
        <div style="text-align:center; border-bottom: 3px solid #8b0000; padding-bottom:10px;">
            <h2 style="margin:0; color:#8b0000; letter-spacing: 1px;">SHRI SHYAMJI AUTO SERVICE CENTER</h2>
            <p style="margin:4px 0; font-size: 13px;">Near Ambedkar Chowk, Bareilly Road, Kichha, Uttarakhand</p>
            <h3 style="margin:10px 0 0 0; background:#8b0000; color:white; padding:5px; border-radius:3px;">WORKSHOP JOB CARD</h3>
        </div>
        <table style="width:100%; margin-top:15px; font-size:13px; border-collapse:collapse;" border="0">
            <tr><td><b>Job Card No:</b> {jc_no}</td><td style="text-align:right;"><b>Date/Time:</b> {jc_date}</td></tr>
            <tr><td><b>Vehicle No:</b> {veh}</td><td style="text-align:right;"><b>Odometer Reading:</b> {km} KM</td></tr>
            <tr><td><b>Customer Name:</b> {name}</td><td style="text-align:right;"><b>Phone:</b> {phone}</td></tr>
            <tr><td><b>Service Advisor:</b> {advisor}</td><td style="text-align:right;"><b>Assigned Mechanic:</b> {mech}</td></tr>
            <tr><td><b>Service Type:</b> {stype}</td><td></td></tr>
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

def generate_invoice_html(inv_no, inv_date, veh, parts, labor, gst, total, paid, doc_type="Tax Invoice", items_list=None, shop_gst="", db_status="", customer_name="", customer_gst="", customer_address=""):
    balance = total - paid
    title_text = "TAX INVOICE" if doc_type == "Tax Invoice" else doc_type.upper()
    if db_status == "Cancelled": title_text = "CANCELLED DOCUMENT"
    
    table_rows = ""
    if items_list:
        for idx, item in enumerate(items_list):
            amt = float(item.get('Qty', 1)) * float(item.get('Rate', 0))
            p_num = item.get('Part_Number', '')
            hsn = item.get('HSN', '')
            table_rows += f"<tr><td>{idx+1}</td><td>{item.get('Type','')}</td><td>{item.get('Description','')}</td><td>{p_num}</td><td>{hsn}</td><td>{item.get('Qty',1)}</td><td>{float(item.get('Rate',0)):.2f}</td><td style='text-align:right;'>{amt:.2f}</td></tr>"

    return f"""
    <html>
    <body style="font-family: sans-serif; color: #333; padding: 30px; border: 1px solid #eee;">
        <div style="text-align:center; border-bottom: 3px solid #1a237e; padding-bottom:10px;">
            <h1 style="margin:0; color:#1a237e; letter-spacing: 2px;">SHRI SHYAMJI AUTO SERVICE CENTER</h1>
            <p style="margin:5px 0; font-size: 14px;">Near Ambedkar Chowk, Bareilly Road, Kichha, Uttarakhand | <b>GSTIN: {shop_gst}</b></p>
        </div>
        <h3 style="text-align:center; margin-top: 15px;">{title_text}</h3>
        <div style="display:flex; justify-content:space-between; margin-top:20px; font-size: 14px;">
            <div><b>BILL TO:</b><br>{customer_name}<br>{customer_address}<br><b>PARTY GSTIN:</b> {customer_gst}</div>
            <div style="text-align:right;"><b>INV NO:</b> {inv_no}<br><b>DATE:</b> {inv_date}<br><b>VEHICLE:</b> {veh}</div>
        </div>
        <table border="1" style="width:100%; border-collapse:collapse; margin-top:20px; font-size: 12px;">
            <tr style="background:#1a237e; color: white;"><th>#</th><th>TYPE</th><th>DESCRIPTION</th><th>PART NO.</th><th>HSN</th><th>QTY</th><th>RATE</th><th style="text-align:right;">AMOUNT</th></tr>
            {table_rows}
        </table>
        <div style="text-align:right; margin-top:20px; font-size: 14px;">
            <p>Parts Total: ₹{parts:.2f} | Labor Total: ₹{labor:.2f}</p>
            <p>GST (18%): ₹{gst:.2f}</p>
            <h2 style="color:#1a237e; border-top: 2px solid #1a237e; display:inline-block; padding-top:5px;">GRAND TOTAL: ₹{total:.2f}</h2>
            <p><b>Paid: ₹{paid:.2f} | Balance: ₹{balance:.2f}</b></p>
        </div>
        <div style="margin-top:50px; display:flex; justify-content:space-between; font-size:10px;">
            <div style="border-top: 1px solid #333; width: 150px; text-align: center;">Customer Signature</div>
            <div style="border-top: 1px solid #333; width: 200px; text-align: center;">Authorized Signatory for<br><b>Shri Shyamji Auto Service Center</b></div>
        </div>
    </body>
    </html>
    """

@st.set_data_attributes if False else st.cache_data(ttl=2)
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
                        supabase.table("workshop_records").insert({
                            "customer_name": name, "phone_number": phone, "vehicle_number": veh, 
                            "entry_km": str(km_reading), "service_advisor": advisor, "service_type": service, 
                            "mechanic_name": mech, "status": "Queued",
                            "customer_demands": u_demands, "requested_parts": u_parts
                        }).execute()
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
                        if new_stat == "Delivered": updates["delivered_date"] = str(datetime.now(IST).date())
                        supabase.table("workshop_records").update(updates).eq("id", selected_id).execute()
                        st.rerun()

        st.divider()
        st.write("### 📊 Live Service Board & WhatsApp Alerts")
        if cars_data:
            for c in cars_data:
                e_date, e_time = get_ist(c['created_at'])
                if c['status'] == "Delivered" and e_date != datetime.now(IST).strftime('%d-%m-%Y'): continue 
                
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
        st.write("### 🔍 Complete Vehicle Service History")
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
                
                st.write("#### 🛠️ Chronological Workshop Check-ins")
                for item in matched_visits:
                    v_date, v_time = get_ist(item['created_at'])
                    with st.expander(f"📅 Visit Date: {v_date} (Status: {item['status']})"):
                        st.markdown(f"""
                        * **Service Type:** {item.get('service_type', 'N/A')}
                        * **Odometer (KM) Reading:** {item.get('entry_km', 'N/A')}
                        * **Assigned Mechanic:** {item.get('mechanic_name', 'N/A')}
                        * **Service Advisor:** {item.get('service_advisor', 'N/A')}
                        * **Demanded Works Recorded:** {item.get('customer_demands', 'None')}
                        * **Parts Checklist:** {item.get('requested_parts', 'None')}
                        """)
                
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
                                    st.dataframe(pd.DataFrame(inv_items), use_container_width=True, hide_index=True)
                                except: st.caption("No dynamic parts breakdown found.")
                else: st.warning("No invoices generated for this vehicle yet.")
            else: st.error("No service history found for this vehicle number.")

    with sub_tab_jc_print:
        st.write("### 📋 Live Data Entry & Update JobCards")
        
        # Pull dynamic dropdown mapping active workshop lines
        active_jc_options = {c['id']: f"{c['vehicle_number']} - {c['customer_name']} (JC-{c['id']})" for c in cars_data if c['status'] != 'Delivered'}
        
        if active_jc_options:
            default_index = 0
            if 'selected_jc_id' in st.session_state and st.session_state['selected_jc_id'] in active_jc_options:
                default_index = list(active_jc_options.keys()).index(st.session_state['selected_jc_id'])
                
            selected_id_key = st.selectbox("Select Active Vehicle to View or Enter New Work Data:", options=list(active_jc_options.keys()), format_func=lambda x: active_jc_options[x], index=default_index)
            
            # Fetch fresh real-time record object state from memory mapping
            jc_tgt = next((x for x in cars_data if x['id'] == selected_id_key), None)
            
            if jc_tgt:
                st.markdown("---")
                st.write("#### ✍️ Enter / Modify Job Card Information")
                
                # Fields prefilled allowing live append and modify actions
                updated_demands = st.text_area("Customer Demanded Works / Client Complaints:", value=jc_tgt.get('customer_demands', ''))
                updated_parts = st.text_area("Required Parts / Materials Checklist:", value=jc_tgt.get('requested_parts', ''))
                
                if st.button("💾 UPDATE & SAVE JOBCARD DATA", type="primary"):
                    supabase.table("workshop_records").update({
                        "customer_demands": updated_demands,
                        "requested_parts": updated_parts
                    }).eq("id", jc_tgt['id']).execute()
                    st.success("✅ Job Card entries successfully synced to cloud database!"); time.sleep(0.5); st.rerun()
                
                st.markdown("---")
                st.write("#### 🖨️ Printable Document Preview Sheet")
                jc_date_str, jc_time_str = get_ist(jc_tgt['created_at'])
                
                compiled_jc_html = generate_jobcard_html(
                    jc_no=f"JC-{jc_tgt['id']}", jc_date=f"{jc_date_str} {jc_time_str}", 
                    veh=jc_tgt['vehicle_number'], name=jc_tgt['customer_name'], 
                    phone=jc_tgt.get('phone_number','N/A'), km=jc_tgt.get('entry_km','0'), 
                    advisor=jc_tgt.get('service_advisor','N/A'), mech=jc_tgt.get('mechanic_name','N/A'), 
                    stype=jc_tgt.get('service_type','Running Repair'), 
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
    
    if 'current_html' not in st.session_state: st.session_state['current_html'] = None
    if 'current_inv_no' not in st.session_state: st.session_state['current_inv_no'] = ""
    if 'edit_mode' not in st.session_state: st.session_state['edit_mode'] = False
    if 'target_db_id' not in st.session_state: st.session_state['target_db_id'] = None
    if 'current_bill_status' not in st.session_state: st.session_state['current_bill_status'] = "Active"

    b_mode1, b_mode2 = st.tabs(["⚙️ Create / Edit Document", "🔍 Search & Update Past Invoices"])
    
    with b_mode2:
        st.write("#### 🔎 Locate Document from Database")
        search_inv_no = st.text_input("Enter exact Invoice Number to Fetch & Edit (e.g., SS-2605...):").strip()
        if st.button("🔍 FETCH INVOICE FOR EDITING"):
            if search_inv_no:
                match = next((b for b in bills_data if b['invoice_number'] == search_inv_no), None)
                if match:
                    st.session_state['current_inv_no'] = match['invoice_number']
                    st.session_state['target_db_id'] = match['id']
                    st.session_state['edit_mode'] = True
                    st.session_state['current_bill_status'] = match.get('payment_status', 'Active')
                    
                    st.session_state['bill_vnum'] = match.get('vehicle_number', '')
                    st.session_state['bill_cname'] = match.get('customer_name', '')
                    st.session_state['bill_cgst'] = match.get('customer_gst', '')
                    st.session_state['bill_caddr'] = match.get('customer_address', '')
                    st.session_state['bill_precv'] = float(match.get('amount_paid', 0))
                    
                    try: st.session_state["bill_items_df"] = json.loads(match['invoice_details'])
                    except: st.session_state["bill_items_df"] = [{"Type": "Part", "Part_Number": "", "Description": "", "HSN": "2710", "Qty": 1.0, "Rate": 0.0}]
                    
                    st.success(f"Loaded {search_inv_no}! Switch tabs to finish modifications or Cancel it.")
                else: st.error("Invoice number not found.")

    with b_mode1:
        if st.session_state['edit_mode']:
            st.warning(f"⚠️ **EDIT MODE ACTIVE:** Modifying existing invoice number: **{st.session_state['current_inv_no']}** | Status: **{st.session_state['current_bill_status']}**")
            if st.button("Cancel Edit Mode & Reset Form"):
                st.session_state['edit_mode'] = False
                st.session_state['current_html'] = None
                st.session_state['current_inv_no'] = ""
                st.session_state['target_db_id'] = None
                st.session_state['current_bill_status'] = "Active"
                st.session_state["bill_items_df"] = [{"Type": "Part", "Part_Number": "", "Description": "", "HSN": "2710", "Qty": 1.0, "Rate": 0.0}]
                st.rerun()

        c1, c2 = st.columns(2)
        b_veh = c1.text_input("Vehicle Number", key="bill_vnum").upper()
        b_type = c2.selectbox("Document Type", ["Tax Invoice", "Estimate", "Pre-Invoice"], key="bill_dtype")
        
        c3, c4 = st.columns(2)
        c_name = c3.text_input("Customer / Billing Name", key="bill_cname")
        c_gst = c4.text_input("Customer GSTIN (Optional)", key="bill_cgst")
        
        c5, c6 = st.columns(2)
        w_gst = c5.text_input("Workshop GSTIN", value="05AAIFS1234M1Z1", key="bill_wgst")
        c_addr = c6.text_input("Customer Address (Optional)", key="bill_caddr")
        
        l_disc = st.number_input("Labor Discount (%)", min_value=0.0, max_value=100.0, value=0.0, key="bill_ldisc")
        
        st.write("#### 🛒 Itemized Parts & Labor")
        if "bill_items_df" not in st.session_state:
            st.session_state["bill_items_df"] = [{"Type": "Part", "Part_Number": "", "Description": "", "HSN": "2710", "Qty": 1.0, "Rate": 0.0}]
            
        edited_items = st.data_editor(st.session_state["bill_items_df"], num_rows="dynamic", use_container_width=True, key="billing_editor")
        
        processed_items = []
        if edited_items:
            for row in edited_items:
                part_no_clean = str(row.get('Part_Number', '')).strip().upper()
                if row.get('Type') == "Part" and part_no_clean in inventory_dict:
                    inv_match = inventory_dict[part_no_clean]
                    if not row.get('Description'): row['Description'] = inv_match.get('part_name', '')
                    if float(row.get('Rate', 0.0)) == 0.0: row['Rate'] = float(inv_match.get('selling_price', 0.0))
                processed_items.append(row)
            st.session_state["bill_items_df"] = processed_items

        p_recv = st.number_input("Amount Received Now (₹)", min_value=0.0, value=0.0, key="bill_precv")
        p_due = st.date_input("Payment Due Date (If Pending Balance)", key="bill_pdue")
        
        parts_sub = sum(float(r['Qty']) * float(r['Rate']) for r in processed_items if r.get('Type') == "Part")
        labor_sub = sum(float(r['Qty']) * float(r['Rate']) for r in processed_items if r.get('Type') == "Labor")
        labor_after_disc = labor_sub * (1 - l_disc / 100)
        gst_amount = (parts_sub + labor_after_disc) * 0.18
        grand_total = round(parts_sub + labor_after_disc + gst_amount)

        st.markdown(f"### 💰 Net Document Total: ₹{grand_total:,.2f}")
        
        if st.session_state['edit_mode']:
            b_col1, b_col2 = st.columns(2)
            with b_col1:
                submit_bill = st.button("💾 SAVE & UPDATE CLOUD RECORD", type="primary", use_container_width=True)
                submit_cancel = False
            with b_col2:
                submit_cancel = st.button("🚨 CANCEL THIS INVOICE (RESTORE STOCK)", use_container_width=True)
                submit_bill = False
        else:
            submit_bill = st.button("🚀 GENERATE NEW INVOICE NUMBER", type="primary", use_container_width=True)
            submit_cancel = False

        if submit_bill:
            if b_veh and processed_items:
                if st.session_state['edit_mode']:
                    inv_no = st.session_state['current_inv_no']
                else:
                    prefix = "PRE-" if b_type == "Pre-Invoice" else ("EST-" if b_type == "Estimate" else "SS-")
                    inv_no = f"{prefix}{datetime.now().strftime('%y%m%d%H%M')}"
                
                if b_type == "Tax Invoice" and not st.session_state['edit_mode']:
                    for item in processed_items:
                        if item.get('Type') == "Part" and item.get('Part_Number') != "-":
                            pn = str(item.get('Part_Number')).strip().upper()
                            if pn in inventory_dict:
                                new_stock = float(inventory_dict[pn].get('stock_qty', 0)) - float(item.get('Qty', 0))
                                supabase.table("shared_inventory").update({"stock_qty": int(new_stock)}).eq("id", inventory_dict[pn]['id']).execute()
                
                payload = {
                    "invoice_number": inv_no,
                    "vehicle_number": b_veh,
                    "customer_name": c_name,
                    "customer_gst": c_gst,
                    "customer_address": c_addr,
                    "total_amount": grand_total,
                    "amount_paid": p_recv,
                    "due_date": str(p_due),
                    "is_estimate": (b_type != "Tax Invoice"),
                    "invoice_details": json.dumps(processed_items),
                    "payment_status": "Active"
                }
                
                if st.session_state['edit_mode']:
                    supabase.table("workshop_billing").update(payload).eq("id", st.session_state['target_db_id']).execute()
                    st.toast("Invoice updated securely!")
                else:
                    supabase.table("workshop_billing").insert(payload).execute()
                    st.toast("Invoice saved to cloud history!")
                
                formatted_date = datetime.now(IST).strftime('%d-%m-%Y')
                st.session_state['current_html'] = generate_invoice_html(
                    inv_no=inv_no, inv_date=formatted_date, veh=b_veh, 
                    parts=parts_sub, labor=labor_after_disc, gst=gst_amount, 
                    total=grand_total, paid=p_recv, doc_type=b_type, 
                    items_list=processed_items, shop_gst=w_gst, db_status="Active", 
                    customer_name=c_name, customer_gst=c_gst, customer_address=c_addr
                )
                st.session_state['current_inv_no'] = inv_no
                st.session_state['edit_mode'] = False
                st.rerun()

        if submit_cancel:
            if st.session_state['target_db_id']:
                supabase.table("workshop_billing").update({"payment_status": "Cancelled"}).eq("id", st.session_state['target_db_id']).execute()
                
                if b_type == "Tax Invoice":
                    for item in processed_items:
                        if item.get('Type') == "Part" and item.get('Part_Number') != "-":
                            pn = str(item.get('Part_Number')).strip().upper()
                            if pn in inventory_dict:
                                add_back_stock = float(inventory_dict[pn].get('stock_qty', 0)) + float(item.get('Qty', 0))
                                supabase.table("shared_inventory").update({"stock_qty": int(add_back_stock)}).eq("id", inventory_dict[pn]['id']).execute()
                
                formatted_date = datetime.now(IST).strftime('%d-%m-%Y')
                st.session_state['current_html'] = generate_invoice_html(
                    inv_no=st.session_state['current_inv_no'], inv_date=formatted_date, veh=b_veh, 
                    parts=parts_sub, labor=labor_after_disc, gst=gst_amount, 
                    total=grand_total, paid=p_recv, doc_type=b_type, 
                    items_list=processed_items, shop_gst=w_gst, db_status="Cancelled", 
                    customer_name=c_name, customer_gst=c_gst, customer_address=c_addr
                )
                st.session_state['edit_mode'] = False
                st.session_state['current_bill_status'] = "Cancelled"
                st.success(f"🚨 Invoice {st.session_state['current_inv_no']} marked as Cancelled & stock restored!")
                time.sleep(1)
                st.rerun()

    if st.session_state['current_html'] is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<h2 style='color:#8b0000; text-align:center;'>📄 Document Panel Sheet</h2>", unsafe_allow_html=True)
        
        met1, met2 = st.columns(2)
        met1.metric(label="Active Document Number", value=st.session_state['current_inv_no'])
        met2.metric(label="Final Document Net Amount", value=f"₹{grand_total:,.2f}")
        
        act_col1, act_col2 = st.columns([1, 1])
        act_col1.download_button(
            label="📥 DOWNLOAD PRINTABLE HTML BILL",
            data=st.session_state['current_html'],
            file_name=f"Invoice_{st.session_state['current_inv_no']}.html",
            mime="text/html",
            use_container_width=True
        )
        
        if act_col2.button("🧹 Clear Workspace & Start Fresh", use_container_width=True):
            st.session_state['current_html'] = None
            st.session_state['current_inv_no'] = ""
            st.session_state['target_db_id'] = None
            st.session_state['edit_mode'] = False
            st.session_state['current_bill_status'] = "Active"
            st.session_state["bill_items_df"] = [{"Type": "Part", "Part_Number": "", "Description": "", "HSN": "2710", "Qty": 1.0, "Rate": 0.0}]
            st.rerun()
            
        components.html(st.session_state['current_html'], height=600, scrolling=True)

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

