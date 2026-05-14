import streamlit as st
from supabase import create_client
from datetime import datetime, timezone, timedelta
import pandas as pd
import json
import time
import io

# ==========================================
# 1. ULTRA-PREMIUM UI & CSS INJECTION
# ==========================================
st.set_page_config(page_title="Galaxy ERP", layout="wide", page_icon="🌌", initial_sidebar_state="expanded")

# This CSS hides default Streamlit branding and creates a custom software feel
st.markdown("""
    <style>
    /* Hide Streamlit Default Header, Footer, and Menu */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Change main background to soft gray for contrast */
    .stApp { background-color: #f4f6f9; }
    
    /* Style Tabs to look like software tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 5px 5px 0 0; padding: 10px 20px; box-shadow: 0px -2px 5px rgba(0,0,0,0.05); }
    .stTabs [aria-selected="true"] { background-color: #1a237e; color: white !important; border-bottom: none; font-weight: bold; }
    
    /* Style Buttons */
    .stButton>button { width: 100%; border-radius: 6px; background-color: #1a237e; color: white; font-weight: 600; border: none; padding: 10px; transition: all 0.3s ease; }
    .stButton>button:hover { background-color: #0d145c; box-shadow: 0 4px 8px rgba(0,0,0,0.2); border: none; color: white;}
    
    /* Metric Cards */
    [data-testid="stMetric"] { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #1a237e; }
    
    /* Form Backgrounds */
    [data-testid="stForm"] { background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); border: 1px solid #e0e0e0; }
    </style>
    """, unsafe_allow_html=True)

# Custom HTML Header Banner (Deep Space Blue Theme)
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
        i = supabase.table("shared_inventory").select("*").order("part_name").execute().data
        return b, i
    except: return [], []

bills_data, inv_data = fetch_galaxy_data()

# ==========================================
# PROFESSIONAL INVOICE GENERATOR
# ==========================================
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
tab1, tab2, tab3 = st.tabs(["🧾 PROFESSIONAL BILLING", "📦 CENTRAL INVENTORY", "📂 DATABASE & EXCEL"])

# --- TAB 1: BILLING & AUTO-DEDUCT ---
with tab1:
    st.write("### 📝 Invoice Generation Engine")
    
    with st.container():
        c1, c2, c3 = st.columns(3)
        b_veh = c1.text_input("Vehicle Number").upper()
        c_name = c2.text_input("Customer / Billing Name")
        c_gst = c3.text_input("Customer GSTIN (Optional)")
        
        c4, c5, c6 = st.columns(3)
        shop_gst_input = c4.text_input("Workshop GSTIN", value="05XXXXX")
        c_addr = c5.text_input("Customer Address (Optional)")
        doc_type = c6.selectbox("Document Type", ["Tax Invoice", "Estimate", "Pre-Invoice"])

    st.write("#### 🛒 Itemized Parts & Labor")
    
    df_items = pd.DataFrame([
        {"Type": "Part", "Description": "Premium Engine Oil", "Part_Number": "MGP-XXX", "HSN": "2710", "Qty": 1.0, "Rate": 2500.0},
        {"Type": "Labor", "Description": "General Service", "Part_Number": "-", "HSN": "9987", "Qty": 1.0, "Rate": 1200.0}
    ])
    
    edited_items = st.data_editor(
        df_items, num_rows="dynamic", use_container_width=True,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Part", "Labor"], required=True),
            "Description": st.column_config.TextColumn("Description / Part Name", required=True),
            "Part_Number": st.column_config.TextColumn("Part No. (Exact Match for Auto-Deduct)"),
            "HSN": st.column_config.TextColumn("HSN Code (For GST Print)"),
            "Qty": st.column_config.NumberColumn("Qty", min_value=0.1, format="%.1f"),
            "Rate": st.column_config.NumberColumn("Rate (₹)", min_value=0.0, format="%.2f")
        }
    )
    
    p_total = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items.iterrows() if r['Type'] == "Part")
    l_total = sum(float(r['Qty']) * float(r['Rate']) for _, r in edited_items.iterrows() if r['Type'] == "Labor")
    
    st.divider()
    col_d1, col_d2 = st.columns(2)
    l_disc = col_d1.number_input("Labor Discount (%)", min_value=0, max_value=100, value=0)
    apply_gst = col_d2.checkbox("Apply 18% GST", value=True)
    
    final_labor = l_total - (l_total * (l_disc / 100))
    gst_val = (p_total + final_labor) * 0.18 if apply_gst else 0.0
    grand_total = round(p_total + final_labor + gst_val)
    
    st.markdown(f"""
        <div style="background-color: #e8eaf6; padding: 15px; border-radius: 8px; border: 1px solid #c5cae9; text-align: center; margin-bottom: 20px;">
            <h3 style="margin:0; color: #1a237e;">GRAND TOTAL: ₹{grand_total:,.2f}</h3>
            <p style="margin:5px 0 0 0; font-size: 14px; color: #3949ab;">Parts: ₹{p_total:,.2f} | Labor (After Disc): ₹{final_labor:,.2f} | GST: ₹{gst_val:,.2f}</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.form("save_invoice_form"):
        c_pay1, c_pay2 = st.columns(2)
        paid_amt = c_pay1.number_input("Amount Received Now (₹)", min_value=0.0, max_value=float(grand_total))
        due_date_input = c_pay2.date_input("Payment Due Date (If Pending Balance)")
        
        if st.form_submit_button("🚀 GENERATE DOCUMENT & SYNC INVENTORY"):
            if b_veh:
                inv_no = f"GAL-{datetime.now().strftime('%y%m%d%H%M')}"
                items_list = edited_items.to_dict('records')
                is_final_bill = (doc_type == "Tax Invoice")
                
                can_proceed, error_messages, deduction_tasks = True, [], []

                for item in items_list:
                    if item.get('Type') == "Part":
                        p_num_input = str(item.get('Part_Number', '')).strip().upper()
                        qty_needed = float(item.get('Qty', 0))
                        
                        if is_final_bill and not p_num_input:
                            can_proceed = False; error_messages.append(f"❌ Missing Part Number for '{item.get('Description')}'.")
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
                                    can_proceed = False; error_messages.append(f"❌ Invalid Part No: {p_num_input} not found in Master Inventory.")

                if not can_proceed:
                    for err in error_messages: st.error(err)
                    st.warning("⚠️ Fix Part Numbers/Stock or change Document Type to 'Estimate' to bypass stock checks.")
                else:
                    if is_final_bill:
                        for task in deduction_tasks:
                            supabase.table("shared_inventory").update({"stock_qty": task['new_qty']}).eq("id", task['id']).execute()
                
                    payment_status = "Paid" if paid_amt >= grand_total else ("Pending" if paid_amt == 0 else "Partial")
                    
                    supabase.table("galaxy_billing").insert({
                        "invoice_number": inv_no, "vehicle_number": b_veh, "customer_name": c_name,
                        "customer_gst": c_gst, "customer_address": c_addr, "total_amount": grand_total,
                        "amount_paid": paid_amt, "parts_cost": p_total, "final_labor": final_labor,
                        "gst_amount": gst_val, "payment_status": payment_status, "invoice_details": json.dumps(items_list),
                        "is_estimate": (not is_final_bill), "shop_gst": shop_gst_input, "due_date": str(due_date_input)
                    }).execute()
                    
                    st.success(f"✅ {doc_type} {inv_no} saved successfully!"); time.sleep(1.5); st.rerun()
            else:
                st.error("Please enter a Vehicle Number.")

# --- TAB 2: SMART INVENTORY ---
with tab2:
    st.write("### 📦 Master Inventory Management")
    with st.expander("📥 Register Inward Stock (Smart Merge)", expanded=False):
        with st.form("inv_add_form", clear_on_submit=True):
            c_name, c_num, c_qty, c_price = st.columns([2, 1, 1, 1])
            i_name = c_name.text_input("Part Name / Description *")
            i_num = c_num.text_input("Part Number (Exact Match)").upper().strip()
            i_qty = c_qty.number_input("Qty to Add", min_value=0)
            i_price = c_price.number_input("Rate (₹)", min_value=0.0, format="%.2f")
            
            if st.form_submit_button("UPDATE MASTER STOCK"):
                if i_name and i_num:
                    check = supabase.table("shared_inventory").select("*").eq("part_number", i_num).execute()
                    if check.data:
                        old = check.data[0]
                        new_qty = int(old['stock_qty']) + int(i_qty)
                        supabase.table("shared_inventory").update({"stock_qty": new_qty, "selling_price": i_price, "part_name": i_name}).eq("id", old['id']).execute()
                        st.success(f"✅ Stock Merged! {i_name} now has {new_qty} units.")
                    else:
                        supabase.table("shared_inventory").insert({"part_name": i_name, "part_number": i_num, "stock_qty": i_qty, "selling_price": i_price}).execute()
                        st.success(f"✨ New Part Added: {i_name}")
                    time.sleep(1); st.rerun()
                else:
                    st.error("Both Part Name and Part Number are required.")

    st.divider()
    search_q = st.text_input("🔍 Quick Search Catalog (Part No. or Name)")
    if inv_data:
        df_inv = pd.DataFrame(inv_data)
        if search_q:
            df_inv = df_inv[df_inv['part_name'].str.contains(search_q, case=False) | df_inv['part_number'].str.contains(search_q, case=False)]
        
        st.write("*(Hint: Click directly inside the table below to edit Quantity or Price, then hit Save Quick Edits!)*")
        edited_inv = st.data_editor(
            df_inv[['id', 'part_number', 'part_name', 'stock_qty', 'selling_price']],
            use_container_width=True, hide_index=True,
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
                orig_row = df_inv.iloc[index]
                if row['stock_qty'] != orig_row['stock_qty'] or row['selling_price'] != orig_row['selling_price']:
                    supabase.table("shared_inventory").update({"stock_qty": row['stock_qty'], "selling_price": row['selling_price']}).eq("id", row['id']).execute()
            st.success("✅ Changes Saved!"); time.sleep(1); st.rerun()

# --- TAB 3: MANAGE DATABASE ---
with tab3:
    st.write("### 📂 Document Repository & Database")
    
    if bills_data:
        st.write("#### 📥 Secure Excel Backup")
        df_bills = pd.DataFrame(bills_data)
        st.download_button("📊 Export All Galaxy Bills to Excel (.xlsx)", data=to_excel(df_bills), file_name=f"Galaxy_Full_Backup_{datetime.now(IST).strftime('%d-%m-%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.divider()
        
        search_inv = st.text_input("🔍 Enter exact Document Number to Print/Cancel (e.g., GAL-2405...)").strip().upper()
        if search_inv:
            target = next((b for b in bills_data if b['invoice_number'] == search_inv), None)
            if target:
                st.success(f"Found: {target['vehicle_number']} | Total: ₹{target['total_amount']} | Status: {target['payment_status']}")
                
                c1, c2 = st.columns(2)
                with c1:
                    date_str, _ = get_ist(target['created_at'])
                    html = generate_invoice_html(
                        inv_no=target['invoice_number'], inv_date=date_str, veh=target['vehicle_number'],
                        parts=float(target.get('parts_cost', 0)), labor=float(target.get('final_labor', 0)),
                        gst=float(target.get('gst_amount', 0)), total=float(target['total_amount']),
                        paid=float(target.get('amount_paid', 0)), doc_type="Estimate" if target.get('is_estimate') else "Tax Invoice",
                        items_list=json.loads(target.get('invoice_details', '[]')), shop_gst=target.get('shop_gst', ''),
                        db_status=target.get('payment_status', ''), customer_name=target.get('customer_name', ''),
                        customer_gst=target.get('customer_gst', ''), customer_address=target.get('customer_address', '')
                    )
                    st.download_button("🖨️ Print / Download Bill", data=html, file_name=f"{target['invoice_number']}.html", mime="text/html", use_container_width=True)
                with c2:
                    if st.button("❌ CANCEL THIS BILL", type="primary", use_container_width=True):
                        supabase.table("galaxy_billing").update({"payment_status": "Cancelled"}).eq("id", target['id']).execute()
                        st.rerun()
            else:
                st.warning("Document not found.")

        st.write("#### 🧾 Recent Documents")
        cols = [c for c in ['invoice_number', 'vehicle_number', 'customer_name', 'total_amount', 'amount_paid', 'payment_status'] if c in df_bills.columns]
        st.dataframe(df_bills[cols], use_container_width=True, hide_index=True)
    else:
        st.info("No documents found in the database yet.")
