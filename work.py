import streamlit as st
from supabase import create_client
import urllib.parse
from datetime import datetime, timezone, timedelta
import pandas as pd
import json
import time
import io

# ==========================================
# 1. ULTRA-PREMIUM UI & CSS INJECTION
# ==========================================
st.set_page_config(page_title="Shyamji ERP", layout="wide", page_icon="🛠️", initial_sidebar_state="expanded")

# This CSS hides default Streamlit branding and creates a custom software feel
st.markdown("""
    <style>
    /* Hide Streamlit Default Header, Footer, and Menu */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Change main background to soft gray for contrast */
    .stApp { background-color: #f0f2f6; }
    
    /* Style Tabs to look like software tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 5px 5px 0 0; padding: 10px 20px; box-shadow: 0px -2px 5px rgba(0,0,0,0.05); }
    .stTabs [aria-selected="true"] { background-color: #8b0000; color: white !important; border-bottom: none; font-weight: bold; }
    
    /* Style Buttons */
    .stButton>button { width: 100%; border-radius: 6px; background-color: #8b0000; color: white; font-weight: 600; border: none; padding: 10px; transition: all 0.3s ease; }
    .stButton>button:hover { background-color: #660000; box-shadow: 0 4px 8px rgba(0,0,0,0.2); border: none; color: white;}
    
    /* Metric Cards */
    [data-testid="stMetric"] { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #8b0000; }
    
    /* Form Backgrounds */
    [data-testid="stForm"] { background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); border: 1px solid #e0e0e0; }
    </style>
    """, unsafe_allow_html=True)

# Custom HTML Header Banner
st.markdown("""
    <div style="background: linear-gradient(135deg, #8b0000 0%, #d32f2f 100%); padding: 25px; border-radius: 10px; color: white; text-align: center; margin-bottom: 25px; margin-top: -40px; box-shadow: 0 4px 15px rgba(139,0,0,0.3);">
        <h1 style="margin:0; font-size: 36px; font-weight: 800; letter-spacing: 1px; color: white;">🛠️ SHRI SHYAMJI AUTO SERVICE CENTER</h1>
        <p style="margin:5px 0 0 0; font-size: 16px; opacity: 0.9;">Master Enterprise Resource Planning (ERP) System</p>
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
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# --- PROFESSIONAL INVOICE GENERATOR ---
def generate_invoice_html(inv_no, inv_date, veh, parts, labor, gst, total, paid, is_estimate=False, items_list=None, shop_gst="", db_status="", customer_gst="", customer_address="", customer_name=""):
    doc_title = "ESTIMATE / QUOTATION" if is_estimate else "TAX INVOICE / BILL OF SUPPLY"
    balance = total - paid
    
    if db_status == "Cancelled":
        doc_title = "CANCELLED DOCUMENT"
        status = "CANCELLED (VOID)"
    else:
        status = "Pending Approval" if is_estimate else ("Paid (Settled)" if balance <= 0 else "Partial/Pending Payment")
    
    table_rows = ""
    if items_list and len(items_list) > 0:
        for idx, item in enumerate(items_list):
            amt = float(item.get('Qty', 1)) * float(item.get('Rate', 0))
            p_num = item.get('Part_Number', '')
            hsn = item.get('HSN', '')
            table_rows += f"<tr><td>{idx+1}</td><td>{item.get('Type', '')}</td><td>{item.get('Description','')}</td><td>{p_num}</td><td>{hsn}</td><td>{item.get('Qty',1)}</td><td>{float(item.get('Rate',0)):.2f}</td><td style='text-align:right;'>{amt:.2f}</td></tr>"

    gst_html = f"<p><strong>Your GSTIN:</strong> {shop_gst}</p>" if shop_gst else ""
    cust_info_html = f"<p><strong>Billed To:</strong> {customer_name}</p>" if customer_name else ""
    cust_gst_html = f"<p><strong>Party GSTIN:</strong> {customer_gst}</p>" if customer_gst else ""
    cust_addr_html = f"<p><strong>Address:</strong> {customer_address}</p>" if customer_address else ""

    return f"""
    <html>
    <body style="font-family: Arial, sans-serif; padding: 40px; color: #333; border: 1px solid #ddd;">
        <div style="text-align: center; border-bottom: 3px solid #8b0000; padding-bottom: 15px; margin-bottom: 30px;">
            <div style="font-size: 18px; font-weight: bold; color: #8b0000; text-transform: uppercase;">Maruti Authorised Service Station</div>
            <h1 style="margin: 5px 0; font-size: 28px; text-transform: uppercase;">SHRI SHYAMJI AUTO SERVICE CENTER</h1>
            <p style="margin: 0; font-size: 14px; color: #555;">Barielly Road, Near Ambedkar Chowk, Kichha, Uttarakhand</p>
            <p style="margin: 0; font-size: 14px; color: #555;">Service: 9837133377, 9837833377 | Insurance: 9012520000</p>
        </div>
        <h3 style="text-align:center; text-decoration: underline;">{doc_title}</h3>
        <div style="display: flex; justify-content: space-between; margin-bottom: 20px; font-size: 14px;">
            <div><p><strong>Ref No:</strong> {inv_no}</p><p><strong>Date:</strong> {inv_date}</p>{gst_html}</div>
            <div style="text-align: right;">{cust_info_html}<p><strong>Vehicle Number:</strong> {veh}</p>{cust_gst_html}{cust_addr_html}</div>
        </div>
        <table border="1" style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px;">
            <tr style="background-color: #8b0000; color: white;"><th>#</th><th>TYPE</th><th>DESCRIPTION</th><th>PART NO.</th><th>HSN CODE</th><th>QTY</th><th>RATE</th><th style="text-align:right;">AMOUNT</th></tr>
            {table_rows}
        </table>
        <table style="width: 50%; margin-left: auto; margin-top: 20px; font-size: 14px;">
            <tr><td><strong>Parts Total:</strong></td><td style="text-align:right;">₹{parts:.2f}</td></tr>
            <tr><td><strong>Labor Total:</strong></td><td style="text-align:right;">₹{labor:.2f}</td></tr>
            <tr><td><strong>GST (18%):</strong></td><td style="text-align:right;">₹{gst:.2f}</td></tr>
        </table>
        <div style="margin-top: 20px; border-top: 2px solid #333; padding-top: 15px; text-align: right; font-size: 16px;">
            <h2 style="color: #8b0000; margin: 0 0 10px 0;">Grand Total: ₹{total:.2f}</h2>
            {f'<p style="margin:0;"><strong>Amount Paid / Advance:</strong> ₹{paid:.2f}</p><p style="margin:5px 0;"><strong>Balance Due:</strong> ₹{balance:.2f}</p>' if not is_estimate and db_status != 'Cancelled' else ''}
            <p style="color:{'red' if db_status == 'Cancelled' else 'black'}; margin-top: 10px;"><strong>Status:</strong> {status}</p>
        </div>
        <div style="margin-top: 60px; display: flex; justify-content: space-between; font-size: 12px; color: #777;">
            <div style="border-top: 1px solid #777; width: 200px; text-align: center; padding-top: 5px;">Customer Signature</div>
            <div style="border-top: 1px solid #777; width: 200px; text-align: center; padding-top: 5px;">Authorized Signatory<br><b>Shri Shyamji Auto</b></div>
        </div>
    </body>
    </html>
    """

@st.cache_data(ttl=2)
def fetch_all_data():
    try:
        w = supabase.table("workshop_records").select("*").order("created_at", desc=True).execute().data
        b = supabase.table("workshop_billing").select("*").order("created_at", desc=True).execute().data
        a = supabase.table("staff_attendance").select("*").order("created_at", desc=True).execute().data
        i = supabase.table("shared_inventory").select("*").order("part_name").execute().data
        e = supabase.table("workshop_expenses").select("*").order("created_at", desc=True).execute().data
        return w, b, a, i, e
    except: return [], [], [], [], []

cars_data, bills_data, att_data, inv_data, exp_data = fetch_all_data()

mechanics = ["Imran", "Rajesh", "Yusuf", "Chandu", "Arman", "Yunnish", "Saif Electrician"]
service_advisors = ["Arun", "Mandeep", "Sanjay", "Admin"]
sa_contacts = {"Arun": "919027831842", "Mandeep": "919756016402", "Sanjay": "919837133377", "Admin": "919837133377"}

# ==========================================
# 3. SIDEBAR DASHBOARD
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #8b0000;'>Command Center</h2>", unsafe_allow_html=True)
    st.info(f"📅 **Date:** {datetime.now(IST).strftime('%d-%m-%Y')}\n\n🕒 **Time:** {datetime.now(IST).strftime('%I:%M %p')}")
    if st.button("🔄 FORCE REFRESH DATA"): st.rerun()
    st.divider()
    if cars_data:
        active = len([c for c in cars_data if c['status'] != 'Delivered'])
        st.metric("🚗 Vehicles in Shop", active)
    if bills_data:
        today = datetime.now(IST).strftime('%d-%m-%Y')
        today_rev = sum(float(b.get('amount_paid', 0)) for b in bills_data if not b.get('is_estimate') and get_ist(b['created_at'])[0] == today)
        st.metric("💰 Today's Collection", f"₹{today_rev:,.0f}")

# ==========================================
# 4. CORE TABS (All 8 Restored)
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "🚗 WORKSHOP FLOW", "🧾 BILLING CENTER", "💸 EXPENSES", "🛡️ CRM & ALERTS", "📊 P&L REPORTS", "👥 STAFF HQ", "🕒 TIMELINE", "📦 MASTER INVENTORY"
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
            phone = st.text_input("Customer Phone (Include 91)")
            km_reading = st.number_input("Entry KM Reading", min_value=0, step=1)
            advisor = st.selectbox("Service Advisor", service_advisors)
            service = st.selectbox("Service Type", ["PMS", "Free Checkup Camp", "Running Repair", "Body Shop"])
            c_m1, c_m2 = st.columns(2)
            with c_m1: mech = st.selectbox("Assign Primary Mechanic", mechanics)
            with c_m2: mech2 = st.selectbox("Assign 2nd Mechanic (Optional)", ["None"] + mechanics)
            ins_date = st.text_input("Insurance Expiry (DD-MM-YYYY) - Optional")
            
            if st.form_submit_button("SAVE NEW ENTRY"):
                if veh and name:
                    supabase.table("workshop_records").insert({
                        "customer_name": name, "phone_number": phone, "vehicle_number": veh, "entry_km": str(km_reading), 
                        "service_advisor": advisor, "service_type": service, "mechanic_name": mech, 
                        "mechanic_2_name": (mech2 if mech2 != "None" else ""), "status": "Queued", "insurance_expiry": ins_date
                    }).execute()
                    st.success("Vehicle Registered Successfully!"); time.sleep(1); st.rerun()
                else: st.error("Vehicle Number and Customer Name required.")

    with col_up:
        st.write("### 🔄 Update Vehicle Status")
        if cars_data:
            active_cars = [f"{c['vehicle_number']} ({c['customer_name']})" for c in cars_data if c['status'] != 'Delivered']
            if active_cars:
                car_to_up = st.selectbox("Select Active Car:", active_cars)
                selected_id = cars_data[[f"{c['vehicle_number']} ({c['customer_name']})" for c in cars_data].index(car_to_up)]['id']
                new_stat = st.radio("Current Status:", ["Queued", "In Workshop", "Washing", "Ready", "Delivered"], horizontal=True)
                if st.button("CONFIRM STATUS CHANGE", type="primary"):
                    if new_stat == "Delivered":
                        supabase.table("workshop_records").update({"status": new_stat, "delivered_date": str(datetime.now(IST).date())}).eq("id", selected_id).execute()
                    else:
                        supabase.table("workshop_records").update({"status": new_stat}).eq("id", selected_id).execute()
                    st.rerun()
            else:
                st.info("No active vehicles currently in the workshop.")

    st.divider()
    st.write("### 📊 Live Service Board")
    if cars_data:
        for c in cars_data:
            e_date, e_time = get_ist(c['created_at'])
            if c['status'] == "Delivered" and e_date != datetime.now(IST).strftime('%d-%m-%Y'): continue 
            
            c1, c2 = st.columns([3, 1.5])
            m1 = c.get('mechanic_name', 'Unassigned')
            m2 = c.get('mechanic_2_name', '')
            mech_display = f"{m1} & {m2}" if m2 else m1
            c1.write(f"**{c['vehicle_number']}** | {c['customer_name']} | KM: {c.get('entry_km', '0')}")
            c1.caption(f"📅 **Entry:** {e_date} {e_time} | SA: {c.get('service_advisor', 'N/A')} | 👨‍🔧 Mech: {mech_display} | Status: **{c['status']}**")
            
            c_phone = c.get('phone_number', '')
            if c['status'] == "Delivered":
                msg = f"Thank you for choosing Shri Shyamji Auto Service Center, {c['customer_name']}! Your vehicle {c['vehicle_number']} is delivered. Safe travels!"
                c2.link_button("🏁 Send 'Thank You'", f"https://wa.me/{c_phone}?text={urllib.parse.quote(msg)}")
            elif c['status'] == "Ready":
                msg = f"Hello {c['customer_name']}! Your vehicle {c['vehicle_number']} is ready for pickup at Shri Shyamji Auto Service Center."
                c2.link_button("🟢 Send 'Ready'", f"https://wa.me/{c_phone}?text={urllib.parse.quote(msg)}")
            else:
                sa_ph = sa_contacts.get(c.get('service_advisor', 'N/A'), "919837133377") 
                msg = f"Welcome to Shri Shyamji Auto Service Center! Your vehicle {c['vehicle_number']} is registered. Your Service Advisor is {c.get('service_advisor', 'N/A')} ({sa_ph.replace('91', '', 1)})."
                c2.link_button("📱 Send SA Details", f"https://wa.me/{c_phone}?text={urllib.parse.quote(msg)}")
            st.divider()

# --- TAB 2: ITEMIZED BILLING & AUTO-DEDUCT ---
with tab2:
    st.write("### 🧾 Invoice Generation Engine")
    
    with st.container():
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            b_veh = st.text_input("Vehicle Number for Bill").upper()
            match_name = next((c['customer_name'] for c in cars_data if c['vehicle_number'] == b_veh), "") if b_veh and cars_data else ""
            customer_name_input = st.text_input("Customer Name", value=match_name)
            customer_gst_input = st.text_input("Customer GST No. (Optional)")
        with col_info2:
            shop_gst_input = st.text_input("Workshop GST No.", value="05XXXXX")
            customer_addr_input = st.text_input("Customer Address (Optional)")
            b_type = st.selectbox("Document Type", ["Tax Invoice", "Estimate"])

    st.write("#### 🛒 Part Details & Labor Breakdown")
    df_items = pd.DataFrame([
        {"Type": "Part", "Description": "Engine Oil 5W30", "Part_Number": "16510M68K10", "HSN": "2710", "Qty": 1.0, "Rate": 1500.0},
        {"Type": "Labor", "Description": "General Service & Washing", "Part_Number": "-", "HSN": "9987", "Qty": 1.0, "Rate": 800.0}
    ])
    
    edited_items = st.data_editor(
        df_items, num_rows="dynamic", use_container_width=True,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Part", "Labor"], required=True),
            "Description": st.column_config.TextColumn("Description / Part Name", required=True),
            "Part_Number": st.column_config.TextColumn("Part No. (Exact Match for Auto-Deduct)"),
            "HSN": st.column_config.TextColumn("HSN Code (For GST)"),
            "Qty": st.column_config.NumberColumn("Qty", min_value=0.1, format="%.1f"),
            "Rate": st.column_config.NumberColumn("Rate (₹)", min_value=0.0, format="%.2f")
        }
    )
    
    parts_sub = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items.iterrows() if r['Type'] == "Part")
    labor_sub = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items.iterrows() if r['Type'] == "Labor")
                
    st.divider()
    c_disc, c_gst = st.columns(2)
    with c_disc: labor_discount = st.number_input("Labor Discount (%)", min_value=0, max_value=100, value=0)
    with c_gst: apply_gst = st.checkbox("Apply 18% GST on Grand Total", value=True)

    final_labor = labor_sub - (labor_sub * (labor_discount / 100))
    gst_amount = (parts_sub + final_labor) * 0.18 if apply_gst else 0.0
    grand_total = round(parts_sub + final_labor + gst_amount)

    st.markdown(f"""
        <div style="background-color: #fce4e4; padding: 15px; border-radius: 8px; border: 1px solid #f5c6c6; text-align: center; margin-bottom: 20px;">
            <h3 style="margin:0; color: #8b0000;">GRAND TOTAL: ₹{grand_total:,.2f}</h3>
            <p style="margin:5px 0 0 0; font-size: 14px; color: #a52a2a;">Parts: ₹{parts_sub:,.2f} | Labor: ₹{final_labor:,.2f} | GST: ₹{gst_amount:,.2f}</p>
        </div>
    """, unsafe_allow_html=True)

    with st.form("save_invoice_form"):
        c_pay1, c_pay2 = st.columns(2)
        with c_pay1:
            advance_paid = st.number_input("Amount Paid Now / Advance (₹)", min_value=0.0, max_value=float(grand_total), step=100.0)
        with c_pay2:
            b_mode = st.selectbox("Payment Method", ["Cash", "UPI / PhonePe", "Card", "Pending (No Payment)"])
            due_date_input = st.date_input("Payment Due Date (If Pending)")
        
        if st.form_submit_button("🚀 GENERATE DOCUMENT & SYNC INVENTORY"):
            if b_veh:
                inv_no = f"{'EST-' if b_type == 'Estimate' else 'SS-'}{datetime.now().strftime('%y%m%d%H%M')}"
                is_final_bill = (b_type == "Tax Invoice")
                items_list = edited_items.to_dict('records')
                
                can_proceed, error_messages, deduction_tasks = True, [], []

                # --- STOCK GUARD LOGIC ---
                for item in items_list:
                    if item.get('Type') == "Part":
                        p_num_input = str(item.get('Part_Number', '')).strip().upper()
                        qty_needed = float(item.get('Qty', 0))
                        
                        if is_final_bill and not p_num_input:
                            can_proceed = False; error_messages.append(f"❌ Missing Part Number for {item.get('Description')}.")
                            continue

                        if p_num_input and p_num_input != "-":
                            check_inv = supabase.table("shared_inventory").select("*").eq("part_number", p_num_input).execute()
                            if check_inv.data:
                                existing_part = check_inv.data[0]
                                current_stock = float(existing_part.get('stock_qty', 0))
                                
                                if is_final_bill:
                                    if current_stock < qty_needed:
                                        can_proceed = False; error_messages.append(f"❌ Out of Stock: {p_num_input} (Have: {current_stock}, Need: {qty_needed})")
                                    else:
                                        deduction_tasks.append({"id": existing_part['id'], "new_qty": current_stock - qty_needed, "num": p_num_input})
                            else:
                                if is_final_bill:
                                    can_proceed = False; error_messages.append(f"❌ Invalid Part No: {p_num_input} not found in Inventory.")

                if not can_proceed:
                    for err in error_messages: st.error(err)
                    st.warning("⚠️ Fix Part Numbers/Stock or change Document Type to 'Estimate' to bypass stock checks.")
                else:
                    if is_final_bill:
                        for task in deduction_tasks:
                            supabase.table("shared_inventory").update({"stock_qty": task['new_qty']}).eq("id", task['id']).execute()
                
                    if not is_final_bill: b_status = "Estimate"
                    elif b_mode == "Pending (No Payment)": b_status = "Pending"
                    elif advance_paid >= grand_total: b_status = f"Paid ({b_mode})"
                    else: b_status = f"Partial ({b_mode})"

                    supabase.table("workshop_billing").insert({
                        "invoice_number": inv_no, "vehicle_number": b_veh, "customer_name": customer_name_input, 
                        "total_amount": grand_total, "amount_paid": advance_paid if is_final_bill else 0.0, 
                        "parts_cost": parts_sub, "final_labor": final_labor, "gst_amount": gst_amount, 
                        "payment_status": b_status, "is_estimate": not is_final_bill, "due_date": str(due_date_input) if b_status in ["Pending", "Partial"] else "", 
                        "shop_gst": shop_gst_input, "customer_gst": customer_gst_input, "customer_address": customer_addr_input, 
                        "invoice_details": json.dumps(items_list)
                    }).execute()
                    
                    st.success(f"Document {inv_no} saved successfully!"); time.sleep(2); st.rerun()
            else: st.error("Please enter a Vehicle Number.")

    st.divider()
    
    st.write("### 🛠️ Modify / Reprint / Cancel Document")
    search_inv = st.text_input("Enter Document No. (e.g., SS-2405...)").strip().upper()
    
    if search_inv and bills_data:
        target = next((b for b in bills_data if b['invoice_number'] == search_inv), None)
        if target:
            if target.get('payment_status') == 'Cancelled':
                st.error(f"❌ Document **{target['invoice_number']}** is Cancelled.")
            else:
                c_edit, c_canc = st.columns([3, 1])
                with c_canc:
                    st.write("") 
                    if st.button("❌ CANCEL DOCUMENT", type="primary", use_container_width=True):
                        supabase.table("workshop_billing").update({"payment_status": "Cancelled"}).eq("id", target['id']).execute()
                        st.rerun()
                
                with c_edit:
                    with st.expander(f"✏️ Edit Details for {target['invoice_number']}", expanded=True):
                        items = json.loads(target.get('invoice_details', '[]')) if target.get('invoice_details') else [{"Type": "Part", "Description": "Legacy Item", "Part_Number": "", "HSN": "", "Qty": 1.0, "Rate": float(target.get('parts_cost',0))}]
                        
                        with st.form(f"edit_form_{target['id']}"):
                            n_veh = st.text_input("Vehicle Number", value=target['vehicle_number'])
                            n_name = st.text_input("Customer Name", value=target.get('customer_name', ''))
                            
                            edited_items_modify = st.data_editor(pd.DataFrame(items), num_rows="dynamic", use_container_width=True)
                            n_paid = st.number_input("Amount Paid (₹)", value=float(target.get('amount_paid', 0)), step=100.0)
                            
                            if st.form_submit_button("💾 Save Updates"):
                                p_mod = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items_modify.iterrows() if r.get('Type') == "Part")
                                l_mod = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items_modify.iterrows() if r.get('Type') == "Labor")
                                g_mod = p_mod + l_mod + (p_mod + l_mod) * 0.18 if float(target.get('gst_amount', 0)) > 0 else 0
                                
                                supabase.table("workshop_billing").update({
                                    "vehicle_number": n_veh, "customer_name": n_name, "invoice_details": json.dumps(edited_items_modify.to_dict('records')),
                                    "amount_paid": n_paid
                                }).eq("id", target['id']).execute()
                                st.success("Updated!"); time.sleep(1); st.rerun()

            date_str, _ = get_ist(target['created_at'])
            html_inv_search = generate_invoice_html(
                inv_no=target['invoice_number'], inv_date=date_str, veh=target['vehicle_number'], 
                parts=float(target.get('parts_cost', 0)), labor=float(target.get('final_labor', 0)), 
                gst=float(target.get('gst_amount', 0)), total=float(target['total_amount']), paid=float(target.get('amount_paid', 0)), 
                is_estimate=target.get('is_estimate', False), items_list=json.loads(target.get('invoice_details', '[]')), 
                shop_gst=target.get('shop_gst', ''), db_status=target.get('payment_status', ''), customer_name=target.get('customer_name', '')
            )
            st.download_button("🖨️ Print / Download Document", data=html_inv_search, file_name=f"{target['invoice_number']}.html", mime="text/html")

    st.divider()
    st.write("### ⚠️ PENDING PAYMENTS MANAGER")
    if bills_data:
        for p in bills_data:
            stat = p.get('payment_status', '')
            is_est = p.get('is_estimate', False)
            
            if "Paid" in stat or stat in ["Cash", "UPI", "Card"] or stat == "Cancelled" or is_est: continue 
                
            total_amt = float(p.get('total_amount', 0))
            paid_amt = float(p.get('amount_paid', 0))
            balance = total_amt - paid_amt
            inv_num = p.get('invoice_number', 'N/A')
            due = p.get('due_date', 'N/A')
            
            st.error(f"🚨 **PENDING:** {p['vehicle_number']} ({p.get('customer_name', 'N/A')}) | Invoice: {inv_num}")
            st.write(f"Balance Due: **₹{balance:.2f}** | **Due Date: {due}**")
            
            c_pay, c_print = st.columns([2, 1])
            with c_pay:
                with st.form(key=f"pay_form_{p['id']}"):
                    clear_mode = st.selectbox("Payment Method", ["Cash", "UPI", "Card"])
                    add_payment = st.number_input("Add Payment (₹)", min_value=0.0, max_value=balance, step=100.0)
                    if st.form_submit_button("Update Balance"):
                        new_paid = paid_amt + add_payment
                        new_status = f"Paid ({clear_mode})" if new_paid >= total_amt else f"Partial ({clear_mode})"
                        supabase.table("workshop_billing").update({"amount_paid": new_paid, "payment_status": new_status}).eq("id", p['id']).execute()
                        st.rerun()
            with c_print:
                date_str, _ = get_ist(p['created_at'])
                try: saved_items = json.loads(p.get('invoice_details', '[]'))
                except: saved_items = []
                html_inv = generate_invoice_html(
                    inv_no=inv_num, inv_date=date_str, veh=p['vehicle_number'], 
                    parts=float(p.get('parts_cost', 0)), labor=float(p.get('final_labor', 0)), 
                    gst=float(p.get('gst_amount', 0)), total=total_amt, paid=paid_amt, 
                    is_estimate=False, items_list=saved_items, shop_gst=p.get('shop_gst', ''), db_status=stat,
                    customer_gst=p.get('customer_gst', ''), customer_address=p.get('customer_address', ''), customer_name=p.get('customer_name', '')
                )
                st.download_button("🖨️ Print Document", data=html_inv, file_name=f"{inv_num}.html", mime="text/html", key=f"dl_{p['id']}")
            st.divider()

# --- TAB 3: EXPENSES ---
with tab3:
    st.write("### 💸 Daily Expense Logger")
    with st.form("exp_form", clear_on_submit=True):
        e_cat = st.selectbox("Expense Category", ["Spare Parts Purchase", "Staff Salary/Advance", "Electricity/Utilities", "Tea & Snacks", "Maintenance", "Other"])
        e_amt = st.number_input("Amount (₹)", min_value=0.0, step=100.0)
        e_desc = st.text_input("Description / Notes")
        if st.form_submit_button("LOG EXPENSE"):
            if e_amt > 0:
                supabase.table("workshop_expenses").insert({"date": datetime.now(IST).strftime('%d-%m-%Y'), "category": e_cat, "amount": e_amt, "description": e_desc}).execute()
                st.success("Expense Logged!"); st.rerun()
                        
    if exp_data:
        st.dataframe(pd.DataFrame(exp_data[:10]), use_container_width=True, hide_index=True)

# --- TAB 4: CRM & HISTORY ---
with tab4:
    st.write("### 🔍 Vehicle Service History")
    search_veh = st.text_input("Enter Vehicle Number to Search").upper()
    if search_veh and cars_data:
        history = [c for c in cars_data if search_veh in c['vehicle_number']]
        if history:
            for h in history:
                date_ist, time_ist = get_ist(h['created_at'])
                m1 = h.get('mechanic_name', 'N/A')
                m2 = h.get('mechanic_2_name', '')
                mech_display = f"{m1} & {m2}" if m2 else m1
                st.info(f"📅 {date_ist} | ⏱️ {time_ist} | KM: {h.get('entry_km', 'N/A')} | SA: {h.get('service_advisor', 'N/A')} | Service: {h['service_type']} | Mech: {mech_display}")
    
    st.divider()
    c_fb, c_ins = st.columns(2)
    with c_fb:
        st.write("### ⭐ 2-Day Feedback Desk")
        if cars_data:
            target_date = str((datetime.now(IST) - timedelta(days=2)).date())
            feedback_list = [c for c in cars_data if c['status'] == "Delivered" and c.get('delivered_date') == target_date and c.get('feedback_sent') != 'Yes']
            if feedback_list:
                for f in feedback_list:
                    st.warning(f"🚘 **{f['vehicle_number']}** ({f['customer_name']})")
                    fb_msg = f"Dear {f['customer_name']}, it's been 2 days since your vehicle {f['vehicle_number']} was serviced at Shri Shyamji Auto Service Center. We hope it's running smoothly! Are you satisfied with the service? Drive safe! 🚗"
                    c_btn1, c_btn2 = st.columns(2)
                    c_btn1.link_button("📱 Send WhatsApp", f"https://wa.me/{f.get('phone_number', '')}?text={urllib.parse.quote(fb_msg)}")
                    if c_btn2.button(f"✅ Mark Sent", key=f"fb_{f['id']}"):
                        supabase.table("workshop_records").update({"feedback_sent": "Yes"}).eq("id", f['id']).execute()
                        st.rerun()
            else: st.success("No pending feedback requests for today!")
                
    with c_ins:
        st.write("### 🛡️ Insurance Renewals (Alerts)")
        if cars_data:
            ins_list = [c for c in cars_data if c.get('insurance_expiry') and c['insurance_expiry'].strip() != ""]
            today_date_obj = datetime.now(IST).date()
            
            for i in ins_list:
                exp_str = i['insurance_expiry'].strip()
                days_left, needs_renewal = None, False
                try:
                    exp_obj = datetime.strptime(exp_str, '%d-%m-%Y').date()
                    days_left = (exp_obj - today_date_obj).days
                except: pass 
                
                c_data, c_btn = st.columns([3, 1])
                if days_left is not None:
                    if days_left < 0:
                        c_data.error(f"🚨 **EXPIRED:** {i['vehicle_number']} ({i['customer_name']}) - {exp_str}")
                        needs_renewal = True
                    elif days_left <= 30:
                        c_data.warning(f"⏳ **DUE IN {days_left} DAYS:** {i['vehicle_number']} ({i['customer_name']}) - {exp_str}")
                        needs_renewal = True
                    else: c_data.success(f"🚘 **{i['vehicle_number']}** ({i['customer_name']}) - {exp_str}")
                else: c_data.info(f"🚘 **{i['vehicle_number']}** ({i['customer_name']}) - {exp_str}")
                    
                ins_msg = f"Hello {i['customer_name']}, a quick reminder from Shri Shyamji Auto Service Center! Your vehicle ({i['vehicle_number']}) insurance expires on {exp_str}. Let us know if we can help you renew it."
                c_btn.link_button("📱 Message", f"https://wa.me/{i.get('phone_number', '')}?text={urllib.parse.quote(ins_msg)}", key=f"ins_wa_{i['id']}")
                
                if needs_renewal:
                    with st.expander(f"🔄 Renew Policy for {i['vehicle_number']}"):
                        with st.form(key=f"renew_form_{i['id']}"):
                            next_year_str = (exp_obj + timedelta(days=365)).strftime('%d-%m-%Y') if days_left is not None else ""
                            new_exp_date = st.text_input("Enter New Expiry Date (DD-MM-YYYY)", value=next_year_str)
                            if st.form_submit_button("✅ Update Expiry Date"):
                                supabase.table("workshop_records").update({"insurance_expiry": new_exp_date}).eq("id", i['id']).execute()
                                st.rerun()

# --- TAB 5: ANALYTICS & P&L ---
with tab5:
    st.write("### 📊 Financial Profit & Loss (P&L)")
    total_collected = sum(float(b.get('amount_paid', 0)) for b in bills_data if not b.get('is_estimate') and b.get('payment_status') != 'Cancelled')
    total_expenses = sum(float(e.get('amount', 0)) for e in exp_data)
    net_profit = total_collected - total_expenses
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Revenue Collected", f"₹{total_collected:,.2f}")
    m2.metric("Total Workshop Expenses", f"₹{total_expenses:,.2f}", delta_color="inverse")
    m3.metric("Actual Net Profit", f"₹{net_profit:,.2f}", delta=f"Margin: {(net_profit/total_collected*100 if total_collected > 0 else 0):.1f}%")
    
    st.divider()
    st.write("#### 📥 Secure Excel Downloads")
    c_dl1, c_dl2, c_dl3, c_dl4 = st.columns(4)
    today_date_ist = datetime.now(IST).strftime('%d-%m-%Y')
    if cars_data: c_dl1.download_button("🚗 Workshop Flow", data=to_excel(pd.DataFrame(cars_data)), file_name=f"Shyamji_Workshop_{today_date_ist}.xlsx", use_container_width=True)
    if bills_data: c_dl2.download_button("💰 Billing Ledger", data=to_excel(pd.DataFrame(bills_data)), file_name=f"Shyamji_Billing_{today_date_ist}.xlsx", use_container_width=True)
    if exp_data: c_dl3.download_button("💸 Expense Log", data=to_excel(pd.DataFrame(exp_data)), file_name=f"Shyamji_Expenses_{today_date_ist}.xlsx", use_container_width=True)
    if att_data: c_dl4.download_button("📅 Attendance", data=to_excel(pd.DataFrame(att_data)), file_name=f"Shyamji_Attendance_{today_date_ist}.xlsx", use_container_width=True)

# --- TAB 6: STAFF HQ & HR ---
with tab6:
    col_att, col_perf = st.columns(2)
    with col_att:
        st.write("### 📅 Staff Attendance")
        st.write(f"**Today's Date:** {datetime.now(IST).strftime('%d-%m-%Y')}")
        with st.form("attendance_form"):
            staff_member = st.selectbox("Select Staff Member", mechanics)
            attendance_status = st.radio("Status", ["Present", "Absent", "Half-Day"], horizontal=True)
            if st.form_submit_button("LOG ATTENDANCE"):
                supabase.table("staff_attendance").insert({"mechanic_name": staff_member, "date": datetime.now(IST).strftime('%d-%m-%Y'), "status": attendance_status}).execute()
                st.success(f"Logged {attendance_status} for {staff_member}"); st.rerun()
        with st.expander("Recent Attendance Logs"):
            if att_data: st.dataframe(pd.DataFrame(att_data[:10]), use_container_width=True, hide_index=True)

    with col_perf:
        st.write("### 👨‍🔧 Mechanic Output")
        st.write("Total cars serviced per mechanic:")
        if cars_data:
            mech_counts = {}
            for c in cars_data:
                m1 = c.get('mechanic_name', 'Unassigned')
                m2 = c.get('mechanic_2_name', '')
                mech_counts[m1] = mech_counts.get(m1, 0) + 1
                if m2 and m2 != "None": mech_counts[m2] = mech_counts.get(m2, 0) + 1
            for m, count in mech_counts.items(): st.success(f"**{m}:** {count} jobs")
                
    st.divider()
    st.write("### 📞 Official Staff Directory")
    c_dir1, c_dir2 = st.columns(2)
    with c_dir1:
        st.info("**👔 Service Advisors**")
        for sa, phone in sa_contacts.items():
            if sa != "Admin": st.write(f"▪️ **{sa}**: {phone.replace('91', '+91 ', 1)}")
    with c_dir2:
        st.info("**👨‍🔧 Mechanics Team**")
        for m in mechanics: st.write(f"▪️ **{m}**")

# --- TAB 7: TIMELINE ---
with tab7:
    st.write("### 🕒 Daily Flow Timeline")
    if cars_data:
        for car in cars_data:
            date_ist, time_ist = get_ist(car['created_at'])
            if date_ist == datetime.now(IST).strftime('%d-%m-%Y'):
                c1, c2, c3 = st.columns([1, 3, 1])
                c1.write(f"⏱️ **{time_ist}**")
                c2.write(f"**{car['vehicle_number']}** | {car['customer_name']}")
                c3.info(car['status'])
                st.divider()

# --- TAB 8: CENTRAL INVENTORY (SMART SYNC) ---
with tab8:
    st.write("### 📦 Central Master Inventory")
    with st.expander("➕ Inward Stock Form (Smart Merge)", expanded=False):
        with st.form("new_part_form_shyamji", clear_on_submit=True):
            c_name, c_num, c_qty, c_price = st.columns([2, 1, 1, 1])
            p_name = c_name.text_input("Part Name / Description *")
            p_num = c_num.text_input("Part Number (Exact Match)")
            p_qty = c_qty.number_input("Quantity to Add", min_value=0, value=0)
            p_price = c_price.number_input("Selling Price (₹)", min_value=0.0, format="%.2f")
            
            if st.form_submit_button("ADD TO MASTER CLOUD"):
                if p_name and p_num:
                    p_num_clean = p_num.strip().upper()
                    check_db = supabase.table("shared_inventory").select("*").eq("part_number", p_num_clean).execute()
                    if check_db.data:
                        existing = check_db.data[0]
                        supabase.table("shared_inventory").update({"stock_qty": int(existing.get('stock_qty', 0)) + int(p_qty), "selling_price": p_price, "part_name": p_name}).eq("id", existing['id']).execute()
                        st.success(f"✅ Stock Merged!")
                    else:
                        supabase.table("shared_inventory").insert({"part_name": p_name, "part_number": p_num_clean, "stock_qty": p_qty, "selling_price": p_price}).execute()
                        st.success(f"✨ New Part Added!")
                    time.sleep(1); st.rerun()
                else: st.error("Part Name and Part Number are required.")

    st.divider()
    if inv_data:
        search_part = st.text_input("🔍 Search Live Catalog (Part No. or Name):")
        df_inv = pd.DataFrame(inv_data)
        if search_part: df_inv = df_inv[df_inv['part_name'].str.contains(search_part, case=False, na=False) | df_inv['part_number'].str.contains(search_part, case=False, na=False)]
            
        st.write("*(Hint: Click directly inside the table below to edit Quantity or Price, then hit Save Quick Edits!)*")
        edited_inv = st.data_editor(
            df_inv[['id', 'part_number', 'part_name', 'stock_qty', 'selling_price']], use_container_width=True, hide_index=True,
            column_config={
                "id": None, 
                "part_number": st.column_config.TextColumn("Part No.", disabled=True), 
                "part_name": st.column_config.TextColumn("Part Name", disabled=True),
                "stock_qty": st.column_config.NumberColumn("Current Stock", min_value=0),
                "selling_price": st.column_config.NumberColumn("Unit Price (₹)", format="%.2f")
            }
        )
        if st.button("💾 SAVE QUICK GRID EDITS", type="primary"):
            for index, row in edited_inv.iterrows():
                if row['stock_qty'] != df_inv.iloc[index]['stock_qty'] or row['selling_price'] != df_inv.iloc[index]['selling_price']:
                    supabase.table("shared_inventory").update({"stock_qty": row['stock_qty'], "selling_price": row['selling_price']}).eq("id", row['id']).execute()
            st.success("✅ Master Stock Updated!"); time.sleep(1); st.rerun()
