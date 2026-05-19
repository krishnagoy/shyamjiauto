import streamlit as st
from supabase import create_client
import urllib.parse
from datetime import datetime, timezone, timedelta
import pandas as pd
import json
import time
import io

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

# Inside your generate_invoice_html function, ensure the loop looks like this:
# Place this at the top of your file with your other functions
def generate_invoice_rows(items_list):
    html_rows = ""
    for item in items_list:
        item_type = item.get('Type', 'N/A') 
        html_rows += f"""
        <tr>
            <td style="border: 1px solid #ddd; padding: 8px;">{item_type}</td> 
            <td style="border: 1px solid #ddd; padding: 8px;">{item.get('Description', '')}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">{item.get('Part_Number', '-')}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">{item.get('Qty', 0)}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">₹{float(item.get('Rate', 0)):,.2f}</td>
        </tr>
        """
    return html_rows

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
            <h1 style="margin:0; color:#1a237e; letter-spacing: 2px;">GALAXY AUTOMOBILES</h1>
            <p style="margin:5px 0; font-size: 14px;">Barielly Road, Kichha, Uttarakhand | <b>GSTIN: {shop_gst}</b></p>
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
            <div style="border-top: 1px solid #333; width: 200px; text-align: center;">Authorized Signatory for<br><b>Galaxy Automobiles</b></div>
        </div>
    </body>
    </html>
    """
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
    
    today_rev = sum(float(b.get('amount_paid', 0)) for b in bills_data if not b.get('is_estimate') and get_ist(b['created_at'])[0] == today_str) if bills_data else 0
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
    "🚗 WORKSHOP", "🧾 BILLING", "💰 PAYROLL & SALARY", "👥 STAFF HQ", "📦 INVENTORY", "💸 EXPENSES", "📊 REPORTS & CRM"
])

# --- TAB 1: WORKSHOP MANAGER ---
with tab1:
    col_add, col_up = st.columns([1, 1])
    with col_add:
        st.write("### ➕ Register New Entry")
        with st.form("add_form", clear_on_submit=True):
            f1, f2 = st.columns(2)
            veh = f1.text_input("Vehicle Number").upper()
            name = f2.text_input("Customer Name")
            phone = st.text_input("Customer Phone")
            km_reading = st.number_input("Entry KM Reading", min_value=0, step=1)
            advisor = st.selectbox("Service Advisor", service_advisors if service_advisors else ["Admin"])
            service = st.selectbox("Service Type", ["PMS", "Running Repair", "Body Shop", "Washing Only"])
            mech = st.selectbox("Assign Primary Mechanic", mechanics if mechanics else ["Unassigned"])
            
            if st.form_submit_button("SAVE NEW ENTRY"):
                if veh and name:
                    supabase.table("workshop_records").insert({"customer_name": name, "phone_number": phone, "vehicle_number": veh, "entry_km": str(km_reading), "service_advisor": advisor, "service_type": service, "mechanic_name": mech, "status": "Queued"}).execute()
                    st.success("Registered!"); time.sleep(1); st.rerun()

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
            
            c1, c2 = st.columns([3, 1.5])
            c1.write(f"**{c['vehicle_number']}** | {c['customer_name']} | KM: {c.get('entry_km', '0')}")
            c1.caption(f"📅 **Entry:** {e_date} {e_time} | SA: {c.get('service_advisor', 'N/A')} | 👨‍🔧 Mech: {c.get('mechanic_name', 'N/A')} | Status: **{c['status']}**")
            
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

# --- TAB 2: BILLING ---
# --- TAB 2: BILLING ---
with tab2:
    st.write("### 🧾 Invoice Generation Engine")
    
    # 1. Input Section
    c1, c2 = st.columns(2)
    b_veh = c1.text_input("Vehicle Number").upper()
    b_type = c2.selectbox("Document Type", ["Tax Invoice", "Estimate", "Pre-Invoice"])
    
    c3, c4 = st.columns(2)
    c_name = c3.text_input("Customer / Billing Name")
    c_gst = c4.text_input("Customer GSTIN (Optional)")
    
    c5, c6 = st.columns(2)
    w_gst = c5.text_input("Workshop GSTIN")
    c_addr = c6.text_input("Customer Address (Optional)")
    
    # 2. Labor Discount
    l_disc = st.number_input("Labor Discount (%)", min_value=0.0, max_value=100.0, value=0.0)

    # 3. Item Editor
    st.write("#### 🛒 Itemized Parts & Labor")
    df_items = pd.DataFrame([{"Type": "Part", "Description": "", "Part_Number": "-", "HSN": "2710", "Qty": 1.0, "Rate": 0.0}])
    edited_items = st.data_editor(df_items, num_rows="dynamic", use_container_width=True)
    items_list = edited_items.to_dict('records') if edited_items is not None else []
    
    # 4. Calculations
    parts_sub = sum(float(r['Qty']) * float(r['Rate']) for r in items_list if r.get('Type') == "Part")
    labor_sub = sum(float(r['Qty']) * float(r['Rate']) for r in items_list if r.get('Type') == "Labor")
    labor_after_disc = labor_sub * (1 - l_disc / 100)
    gst_amount = (parts_sub + labor_after_disc) * 0.18
    grand_total = round(parts_sub + labor_after_disc + gst_amount)

    # 5. Payment Details
    p_recv = st.number_input("Amount Received Now (₹)", min_value=0.0, value=0.0)
    p_due = st.date_input("Payment Due Date (If Pending Balance)")

    st.success(f"### GRAND TOTAL: ₹{grand_total:,.2f}")
    st.caption(f"Parts: ₹{parts_sub:,.2f} | Labor (After Disc): ₹{labor_after_disc:,.2f} | GST: ₹{gst_amount:,.2f}")

    # 6. Save & Generate
    if st.button("🚀 GENERATE DOCUMENT & SYNC"):
        if b_veh and items_list:
            # Set Prefix
            prefix = "PRE-" if b_type == "Pre-Invoice" else ("EST-" if b_type == "Estimate" else "SS-")
            inv_no = f"{prefix}{datetime.now().strftime('%y%m%d%H%M')}"
            
            # Inventory Sync (Only for Tax Invoices)
            if b_type == "Tax Invoice":
                for item in items_list:
                    if item.get('Type') == "Part" and item.get('Part_Number') != "-":
                        check_inv = supabase.table("shared_inventory").select("*").eq("part_number", str(item.get('Part_Number')).strip().upper()).execute()
                        if check_inv.data: 
                            supabase.table("shared_inventory").update({"stock_qty": float(check_inv.data[0].get('stock_qty', 0)) - float(item.get('Qty', 0))}).eq("id", check_inv.data[0]['id']).execute()
            
            # Database Save
            supabase.table("workshop_billing").insert({
                "invoice_number": inv_no,
                "vehicle_number": b_veh,
                "customer_name": c_name,
                "customer_gst": c_gst,
                "customer_address": c_addr,
                "total_amount": grand_total,
                "amount_paid": p_recv,
                "due_date": str(p_due),
                "is_estimate": (b_type != "Tax Invoice"),
                "invoice_details": json.dumps(items_list)
            }).execute()
            
            st.success(f"✅ {b_type} ({inv_no}) saved successfully!")
            
            # Preview
            st.write("#### 📄 Invoice Preview")
            html_rows = ""
            for item in items_list:
                html_rows += f"<tr><td>{item.get('Type')}</td><td>{item.get('Description')}</td><td>{item.get('Qty')}</td><td>₹{float(item.get('Rate', 0)):,.2f}</td></tr>"
            st.markdown(f"<table style='width:100%; border-collapse: collapse;'><tr><th>Type</th><th>Desc</th><th>Qty</th><th>Rate</th></tr>{html_rows}</table>", unsafe_allow_html=True)
            time.sleep(1)
            st.rerun()
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
                
                # --- WORKSHOP HOLIDAY POLICY: 4 days credit if leaves >= 7 ---
                extra_payable = 4 if leaves < 7 else 0
                
                payable_days = period - leaves + extra_payable
                gross = (base / period) * payable_days if period > 0 else 0
                net = gross - float(row['Advances Taken']) + float(row['Incentive / Overtime (₹)']) - float(row['Deductions (₹)'])
                
                # --- WIDER COLUMN RATIO FOR 5+ DIGIT DISPLAY ---
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
                            desc = f"Salary: {row['Name']} | {sel_month_name} {sel_year} | Mode: {p_mode} | Paid By: {paid_by}"
                            supabase.table("workshop_expenses").insert({"date": p_date.strftime('%d-%m-%Y'), "category": "Staff Salary/Advance", "amount": pay_amt, "description": desc}).execute()
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
           
# --- TAB 4: STAFF HQ (RECRUITMENT & SALARY SAVING) ---
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

# --- TAB 6: MASTER EXPENSE LEDGER ---
with tab6:
    col_log, col_view = st.columns([1, 2])
    with col_log:
        st.write("### 💸 Log New Expense")
        with st.form("exp_form_ledger", clear_on_submit=True):
            e_cat = st.selectbox("Category", ["Spare Parts Purchase", "Staff Salary/Advance", "Electricity & Utilities", "Tea, Snacks & Meals", "Workshop Maintenance", "Rent", "Other"])
            e_amt = st.number_input("Amount (₹)", min_value=0.0)
            e_desc = st.text_input("Notes")
            e_date = st.date_input("Date of Expense")
            if st.form_submit_button("LOG EXPENSE"):
                if e_amt > 0:
                    supabase.table("workshop_expenses").insert({"date": e_date.strftime('%d-%m-%Y'), "category": e_cat, "amount": e_amt, "description": e_desc}).execute()
                    st.success("✅ Expense Logged Securely!"); time.sleep(1); st.rerun()

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
            st.markdown(f"<div style='background-color:#ffebee; padding:15px; border-radius:8px; border-left:5px solid #d32f2f;'><h4>Total Expenses: ₹{total_filtered:,.2f}</h4></div>", unsafe_allow_html=True)
            st.dataframe(filtered_exp[['date', 'category', 'amount', 'description']].rename(columns={'date': 'Date', 'category': 'Category', 'amount': 'Amount (₹)', 'description': 'Description'}), use_container_width=True, hide_index=True)

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

