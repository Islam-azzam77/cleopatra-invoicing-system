import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
from fpdf import FPDF
import arabic_reshaper
from bidi.algorithm import get_display

try:
    from num2words import num2words
    NUM2WORDS_AVAILABLE = True
except ImportError:
    NUM2WORDS_AVAILABLE = False

# ---------------------------------------------------------
# إعدادات الصفحة العامة والتنسيق
# ---------------------------------------------------------
st.set_page_config(page_title="نظام إدارة الأسعار والفواتير", page_icon="🧱", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
html, body, [class*="css"] {
    font-family: 'Tajawal', sans-serif;
    direction: rtl;
    text-align: right;
}
.stButton > button {
    background-color: #C89B3C;
    color: white;
    border-radius: 8px;
    padding: 8px 22px;
    font-weight: bold;
    width: 100%;
}
div[data-testid="stVerticalBlock"] > div {
    margin-bottom: 0.3rem;
}
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "cleopatra_pricelist.xlsx"
SHEET_NAME = "قائمة الاسعار المصححة"
SETTINGS_FILE = "settings.json"
CUSTOMERS_FILE = "customers.json"

# نسب الخصم الافتراضية اللي بتاخدها من المورد حسب نوع الصنف
DEFAULT_PURCHASE_DISCOUNTS = {"سيراميك": 28.0, "بورسلين": 31.0, "ديكور": 0.0}

# الزيادات الثابتة الافتراضية (قابلة للتعديل لاحقًا من الإعدادات)
DEFAULT_FIXED_INCREASES = {"سيراميك": 30.0, "بورسلين": 70.0, "ديكور": 0.0}


# ---------------------------------------------------------
# إعدادات عامة (اسم المعرض التجاري + اسم المستخدم + يوزر وباسورد الدخول + الزيادات الثابتة)
# ---------------------------------------------------------
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if "زيادات_ثابتة" not in data:
                data["زيادات_ثابتة"] = DEFAULT_FIXED_INCREASES.copy()
            return data
    return {
        "اسم_المعرض": "اسمك التجاري",
        "اسم_المستخدم": "admin",
        "login_username": "admin",
        "login_password": "1234",
        "زيادات_ثابتة": DEFAULT_FIXED_INCREASES.copy()
    }


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


settings = load_settings()

# ---------------------------------------------------------
# شاشة تسجيل الدخول
# ---------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align:center;'>🔒 تسجيل الدخول</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.container(border=True):
            login_user_input = st.text_input("اسم المستخدم")
            login_pass_input = st.text_input("كلمة المرور", type="password")
            if st.button("دخول", use_container_width=True):
                if (login_user_input == settings.get("login_username", "admin") and
                        login_pass_input == settings.get("login_password", "1234")):
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")
    st.stop()


# ---------------------------------------------------------
# الشريط الجانبي (إعدادات عامة)
# ---------------------------------------------------------
st.sidebar.header("⚙️ الإعدادات العامة")
gallery_name_input = st.sidebar.text_input(
    "اسمك التجاري (يظهر أعلى عرض السعر / الفاتورة)",
    value=settings.get("اسم_المعرض", "اسمك التجاري")
)
username_input = st.sidebar.text_input(
    "اسم المستخدم (يظهر في فوتر الطباعة)",
    value=settings.get("اسم_المستخدم", "admin")
)

st.sidebar.markdown("---")
st.sidebar.markdown("🧱 **الزيادات الثابتة على سعر الليستة**")
st.sidebar.caption("عدّل هذه القيم فقط لو الشركة المصنّعة غيّرت مبلغ الزيادة")

current_increases = settings.get("زيادات_ثابتة", DEFAULT_FIXED_INCREASES.copy())

increase_ceramic_input = st.sidebar.number_input(
    "زيادة السيراميك (جنيه)",
    min_value=0.0,
    value=float(current_increases.get("سيراميك", 30.0)),
    step=1.0
)
increase_porcelain_input = st.sidebar.number_input(
    "زيادة البورسلين (جنيه)",
    min_value=0.0,
    value=float(current_increases.get("بورسلين", 70.0)),
    step=1.0
)
increase_decor_input = st.sidebar.number_input(
    "زيادة الديكور (جنيه)",
    min_value=0.0,
    value=float(current_increases.get("ديكور", 0.0)),
    step=1.0
)

st.sidebar.markdown("---")
st.sidebar.markdown("🔑 **تغيير بيانات تسجيل الدخول**")
new_login_user = st.sidebar.text_input("يوزر تسجيل الدخول الجديد", value=settings.get("login_username", "admin"))
new_login_pass = st.sidebar.text_input("باسورد تسجيل الدخول الجديد", value=settings.get("login_password", "1234"))

if st.sidebar.button("💾 حفظ كل الإعدادات"):
    settings["اسم_المعرض"] = gallery_name_input
    settings["اسم_المستخدم"] = username_input
    settings["login_username"] = new_login_user
    settings["login_password"] = new_login_pass
    settings["زيادات_ثابتة"] = {
        "سيراميك": increase_ceramic_input,
        "بورسلين": increase_porcelain_input,
        "ديكور": increase_decor_input
    }
    save_settings(settings)
    st.sidebar.success("تم حفظ الإعدادات بنجاح ✅")

if st.sidebar.button("🚪 تسجيل خروج"):
    st.session_state.logged_in = False
    st.rerun()

# نأخذ الاسمين مباشرة من الخانة الحية في الشريط الجانبي
# بحيث يظهرا فورًا في الفاتورة حتى لو لسه ما ضغطتش زرار الحفظ
gallery_name = gallery_name_input
admin_username = username_input
fixed_increases = settings.get("زيادات_ثابتة", DEFAULT_FIXED_INCREASES.copy())


# ---------------------------------------------------------
# قاعدة بيانات العملاء
# ---------------------------------------------------------
def load_customers():
    if os.path.exists(CUSTOMERS_FILE):
        with open(CUSTOMERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_customers(customers):
    with open(CUSTOMERS_FILE, "w", encoding="utf-8") as f:
        json.dump(customers, f, ensure_ascii=False, indent=2)


customers_db = load_customers()

if "cust_address" not in st.session_state:
    st.session_state.cust_address = ""
if "cust_phone" not in st.session_state:
    st.session_state.cust_phone = ""
if "cust_prev_balance" not in st.session_state:
    st.session_state.cust_prev_balance = 0.0


# ---------------------------------------------------------
# قراءة ملف الأسعار (ليستة كليوباترا الأساسية)
# ---------------------------------------------------------
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    df.columns = ["المقاس", "اسم_الصنف", "فرز اول", "فرز ثاني", "النوع", "رقم الصفحة"]
    return df


try:
    df = load_data()
except Exception as e:
    st.error(f"⚠️ لم أتمكن من قراءة ملف الأسعار: {e}")
    st.stop()

if "invoice_items" not in st.session_state:
    st.session_state.invoice_items = []


def ar(text):
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


def amount_in_words(amount):
    pounds = int(amount)
    piastres = round((amount - pounds) * 100)
    if NUM2WORDS_AVAILABLE:
        try:
            text = f"فقط {num2words(pounds, lang='ar')} جنيه"
            if piastres > 0:
                text += f" و {num2words(piastres, lang='ar')} قرش"
            text += " لا غير"
            return text
        except Exception:
            pass
    return f"فقط {amount:,.2f} جنيه لا غير"


# ---------------------------------------------------------
# توليد PDF للعميل (يعرض سعر البيع فقط - بدون أي إشارة للتكلفة أو الربح)
# ---------------------------------------------------------
def generate_invoice_pdf(invoice_items, client_name, client_address, client_phone,
                          invoice_date, invoice_number, gallery_name, admin_username,
                          subtotal, invoice_discount_value, shipping_fee, net_after_discount,
                          previous_balance, total_debt, paid_amount, remaining_amount):

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("Amiri", "", "Amiri-Regular.ttf")
    pdf.add_font("Amiri", "B", "Amiri-Bold.ttf")

    if os.path.exists("logo.png"):
        pdf.image("logo.png", 10, 8, 30)

    pdf.set_font("Amiri", "B", 18)
    pdf.cell(0, 12, ar(gallery_name), align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Amiri", "B", 14)
    pdf.cell(0, 10, ar("فاتورة مبيعات"), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    pdf.set_font("Amiri", "", 11)
    pdf.cell(0, 7, ar(f"اسم العميل: {client_name}"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, ar(f"العنوان: {client_address or 'غير مسجل'}"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, ar(f"الهاتف: {client_phone or 'غير مسجل'}"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, ar(f"رقم الفاتورة: {invoice_number}      التاريخ: {invoice_date}"),
             align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    headers = ["#", "الصنف", "المقاس", "النوع", "فرز", "كراتين", "سعة الكرتونة", "عدد الأمتار", "سعر المتر", "الإجمالي"]
    widths = [6, 25, 14, 14, 10, 12, 18, 18, 18, 25]

    pdf.set_font("Amiri", "B", 8)
    pdf.set_fill_color(200, 155, 60)
    for w, h in zip(widths, headers):
        pdf.cell(w, 10, ar(h), border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_font("Amiri", "", 8)
    for idx, row in enumerate(invoice_items, start=1):
        pdf.cell(widths[0], 9, str(idx), border=1, align="C")
        pdf.cell(widths[1], 9, ar(row["الصنف"]), border=1, align="R")
        pdf.cell(widths[2], 9, ar(row["المقاس"]), border=1, align="C")
        pdf.cell(widths[3], 9, ar(row["النوع"]), border=1, align="C")
        pdf.cell(widths[4], 9, ar(row["الفرز"]), border=1, align="C")
        pdf.cell(widths[5], 9, str(row["عدد_الكراتين"]), border=1, align="C")
        pdf.cell(widths[6], 9, f"{row['سعة_الكرتونة']:.2f}", border=1, align="C")
        pdf.cell(widths[7], 9, f"{row['عدد_الأمتار']:.2f}", border=1, align="C")
        pdf.cell(widths[8], 9, f"{row['سعر_البيع_للعميل']:.2f}", border=1, align="C")
        pdf.cell(widths[9], 9, f"{row['إجمالي_البيع']:.2f}", border=1, align="C")
        pdf.ln()

    total_meters_sum = sum(r["عدد_الأمتار"] for r in invoice_items)

    pdf.ln(4)
    pdf.set_font("Amiri", "B", 11)
    pdf.cell(0, 8, ar(f"إجمالي عدد الأمتار: {total_meters_sum:.2f} م²"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, ar(f"إجمالي الفاتورة: {subtotal:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, ar(f"خصم الفاتورة: {invoice_discount_value:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, ar(f"النولون (مصاريف الشحن): {shipping_fee:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, ar(f"صافي هذه الفاتورة: {net_after_discount:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, ar(f"الرصيد السابق: {previous_balance:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Amiri", "B", 13)
    pdf.cell(0, 9, ar(f"إجمالي المطلوب: {total_debt:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Amiri", "B", 11)
    pdf.cell(0, 8, ar(f"المدفوع الآن / العربون: {paid_amount:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Amiri", "B", 13)
    pdf.set_text_color(180, 0, 0)
    pdf.cell(0, 9, ar(f"المتبقي المطلوب: {remaining_amount:,.2f} جنيه"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)

    pdf.set_font("Amiri", "", 10)
    pdf.ln(2)
    pdf.multi_cell(0, 7, ar(amount_in_words(remaining_amount)), align="R")

    pdf.ln(8)
    pdf.set_font("Amiri", "", 8)
    print_time = datetime.now().strftime("%Y/%m/%d - %H:%M:%S")
    pdf.cell(0, 6, ar(f"تاريخ الطباعة: {print_time}      تم بواسطة: {admin_username}"),
             align="C", new_x="LMARGIN", new_y="NEXT")

    return bytes(pdf.output())


# ---------------------------------------------------------
# التبويبات الرئيسية
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 عرض قائمة الأسعار", "🧾 إنشاء فاتورة / عرض سعر", "💰 تكاليف الشراء وصافي الربح"])

with tab1:
    st.subheader("📊 قائمة الأسعار الكاملة (ليستة كليوباترا الأساسية)")
    st.caption("🔍 استخدم القوائم تحت كل عمود للفلترة، بالظبط زي فلتر الإكسل. القوائم متتابعة: كل ما تختار من عمود، القوائم اللي بعده بتتحدث تلقائيًا.")

    if st.button("🧹 مسح كل الفلاتر"):
        for col_name in df.columns:
            st.session_state.pop(f"filter_{col_name}", None)
        st.rerun()

    filtered_df = df.copy()
    filter_cols = st.columns(len(df.columns))

    for i, col_name in enumerate(df.columns):
        with filter_cols[i]:
            options = sorted(filtered_df[col_name].dropna().astype(str).unique().tolist())
            selected_values = st.multiselect(
                col_name,
                options,
                key=f"filter_{col_name}",
                placeholder="اختر..."
            )
            if selected_values:
                filtered_df = filtered_df[filtered_df[col_name].astype(str).isin(selected_values)]

    st.markdown(f"**عدد النتائج المطابقة: {len(filtered_df)} صنف**")
    st.dataframe(filtered_df, use_container_width=True)

with tab2:

    # ---------- بيانات العميل ----------
    with st.container(border=True):
        st.markdown("### 📋 بيانات العميل")
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            client_name = st.text_input("اسم العميل")
        with col2:
            invoice_date = st.date_input("التاريخ")
        with col3:
            invoice_number = st.text_input("رقم الفاتورة")

        if st.button("🔍 تحميل بيانات العميل (لو مسجل قبل كده)"):
            found = customers_db.get(client_name.strip()) if client_name else None
            if found:
                st.session_state.cust_address = found.get("العنوان", "")
                st.session_state.cust_phone = found.get("الهاتف", "")
                st.session_state.cust_prev_balance = float(found.get("الرصيد_السابق", 0.0))
                st.success("تم تحميل بيانات العميل ✅")
            else:
                st.warning("عميل جديد - من فضلك أدخل بياناته يدويًا")

        col1, col2, col3 = st.columns(3)
        with col1:
            client_address = st.text_input("العنوان", key="cust_address")
        with col2:
            client_phone = st.text_input("الهاتف", key="cust_phone")
        with col3:
            previous_balance = st.number_input("الرصيد السابق للعميل (جنيه)", step=0.01, key="cust_prev_balance")

        if st.button("💾 حفظ / تحديث بيانات العميل"):
            if client_name.strip():
                customers_db[client_name.strip()] = {
                    "العنوان": client_address,
                    "الهاتف": client_phone,
                    "الرصيد_السابق": previous_balance
                }
                save_customers(customers_db)
                st.success("تم حفظ بيانات العميل ✅")
            else:
                st.error("من فضلك أدخل اسم العميل أولاً")

    # ---------- بيانات الصنف ----------
    with st.container(border=True):
        st.markdown("### 🧱 بيانات الصنف")

        col1, col2 = st.columns([1, 2])
        with col1:
            sizes_for_invoice = ["الكل"] + sorted(df["المقاس"].dropna().unique().tolist())
            selected_size = st.selectbox("📏 اختر المقاس أولاً", sizes_for_invoice, key="invoice_size_filter")

        if selected_size != "الكل":
            filtered_items_df = df[df["المقاس"] == selected_size]
        else:
            filtered_items_df = df

        filtered_items_list = filtered_items_df["اسم_الصنف"].dropna().unique().tolist()

        with col2:
            if len(filtered_items_list) == 0:
                st.warning("⚠️ لا يوجد أصناف بهذا المقاس")
                st.stop()
            item_name = st.selectbox("اسم الصنف", filtered_items_list, key="invoice_item_select")

        item_row = filtered_items_df[filtered_items_df["اسم_الصنف"] == item_name].iloc[0]
        st.write(f"📏 المقاس الفعلي للصنف: **{item_row['المقاس']}**")

        grade = st.radio("الفرز", ["فرز اول", "فرز ثاني"], horizontal=True)

    # ---------- سعر الشراء (التكلفة) ----------
    with st.container(border=True):
        st.markdown("### 💰 سعر الشراء (تكلفتي)")

        price_before = float(item_row[grade])
        st.write(f"السعر الأساسي بليستة كليوباترا: **{price_before:.2f} جنيه**")

        default_type_text = str(item_row.get("النوع", "")).strip()
        if "بورسلين" in default_type_text:
            default_index = 1
        elif "ديكور" in default_type_text:
            default_index = 2
        else:
            default_index = 0

        item_type = st.radio(
            "نوع الصنف",
            [
                f"سيراميك (+{fixed_increases.get('سيراميك', 30.0):.0f})",
                f"بورسلين (+{fixed_increases.get('بورسلين', 70.0):.0f})",
                f"ديكور (+{fixed_increases.get('ديكور', 0.0):.0f})"
            ],
            index=default_index,
            horizontal=True
        )

        if "سيراميك" in item_type:
            fixed_increase = fixed_increases.get("سيراميك", 30.0)
            discount_key = "سيراميك"
        elif "بورسلين" in item_type:
            fixed_increase = fixed_increases.get("بورسلين", 70.0)
            discount_key = "بورسلين"
        else:
            fixed_increase = fixed_increases.get("ديكور", 0.0)
            discount_key = "ديكور"

        col1, col2 = st.columns(2)
        with col1:
            purchase_discount = st.number_input(
                "نسبة خصمي من المورد (%)",
                min_value=0.0, max_value=100.0,
                value=DEFAULT_PURCHASE_DISCOUNTS.get(discount_key, 0.0),
                step=0.5,
                key=f"purchase_discount_{discount_key}"
            )

        # الترتيب الصحيح: الخصم يُطبَّق على السعر الأساسي أولًا، ثم تُضاف الزيادة الثابتة بعد الخصم
        price_after_discount = price_before * (1 - purchase_discount / 100)
        cost_price = price_after_discount + fixed_increase

        with col2:
            st.info(f"📋 السعر بعد الخصم (قبل الزيادة): **{price_after_discount:.2f} جنيه**")

        st.success(f"✅ سعر تكلفتي الفعلي (الشراء من المورد): **{cost_price:.2f} جنيه / م²**")

    # ---------- سعر البيع للعميل (بحرية كاملة) ----------
    with st.container(border=True):
        st.markdown("### 🏷️ سعر البيع للعميل")
        selling_price = st.number_input(
            "سعر البيع للعميل (جنيه / م²) - براحتك بالكامل",
            min_value=0.0,
            value=float(cost_price),
            step=0.01,
            key=f"selling_price_{item_name}_{grade}"
        )
        profit_per_meter = selling_price - cost_price
        col1, col2 = st.columns(2)
        with col1:
            st.metric("سعر البيع للعميل", f"{selling_price:.2f} جنيه/م²")
        with col2:
            st.metric("ربحي في المتر", f"{profit_per_meter:,.2f} جنيه")

    # ---------- الكمية والكراتين ----------
    with st.container(border=True):
        st.markdown("### 📦 الكمية والكراتين")
        col1, col2 = st.columns(2)
        with col1:
            carton_capacity = st.number_input("سعة الكرتونة (متر)", min_value=0.0, value=1.43, step=0.01)
        with col2:
            carton_count = st.number_input("عدد الكراتين", min_value=0, step=1)

        total_meters = carton_capacity * carton_count
        selling_total = selling_price * total_meters
        cost_total = cost_price * total_meters
        profit_total = selling_total - cost_total

        col1, col2 = st.columns(2)
        with col1:
            st.info(f"📐 عدد الأمتار: **{total_meters:.2f} م²**")
        with col2:
            st.warning(f"💵 إجمالي البيع لهذا الصنف: {selling_total:.2f} جنيه")

    # ---------- إضافة الصنف للفاتورة ----------
    st.markdown("")
    if st.button("➕ إضافة للفاتورة", use_container_width=True):
        if carton_count <= 0:
            st.error("⚠️ من فضلك أدخل عدد كراتين أكبر من صفر قبل الإضافة للفاتورة")
        else:
            st.session_state.invoice_items.append({
                "الصنف": item_name,
                "المقاس": item_row["المقاس"],
                "الفرز": grade,
                "النوع": item_type,
                "السعر_الأساسي": price_before,
                "السعر_بعد_الخصم_قبل_الزيادة": price_after_discount,
                "نسبة_خصم_الشراء": purchase_discount,
                "سعر_التكلفة": cost_price,
                "سعر_البيع_للعميل": selling_price,
                "سعة_الكرتونة": carton_capacity,
                "عدد_الكراتين": carton_count,
                "عدد_الأمتار": total_meters,
                "إجمالي_البيع": selling_total,
                "إجمالي_التكلفة": cost_total,
                "صافي_الربح": profit_total
            })
            st.success(f"تمت إضافة {item_name} للفاتورة ✅")

    # ---------- عرض الفاتورة الحالية مع إمكانية التعديل والحذف لكل صنف ----------
    st.markdown("---")
    st.subheader("📄 الفاتورة الحالية (ما يظهر للعميل)")

    if len(st.session_state.invoice_items) == 0:
        st.info("لا يوجد أصناف مضافة بعد.")
    else:
        st.caption("✏️ تقدر تعدّل عدد الكراتين أو نسبة خصم الشراء أو سعر البيع لأي صنف مباشرة من هنا، أو تحذفه بدون التأثير على باقي الأصناف.")

        items_to_delete = []

        for idx, itm in enumerate(st.session_state.invoice_items):
            with st.container(border=True):
                col1, col2, col3, col4, col5, col6 = st.columns([2, 1, 1, 1, 1, 0.6])

                with col1:
                    st.markdown(f"**{idx + 1}. {itm['الصنف']}**")
                    st.caption(f"{itm['المقاس']} | {itm['الفرز']} | {itm['النوع']}")

                with col2:
                    new_carton_count = st.number_input(
                        "عدد الكراتين", min_value=0, value=int(itm["عدد_الكراتين"]),
                        step=1, key=f"edit_carton_{idx}"
                    )

                with col3:
                    new_purchase_discount = st.number_input(
                        "خصم الشراء %", min_value=0.0, max_value=100.0,
                        value=float(itm["نسبة_خصم_الشراء"]), step=0.5,
                        key=f"edit_discount_{idx}"
                    )

                with col4:
                    new_selling_price = st.number_input(
                        "سعر البيع", min_value=0.0,
                        value=float(itm["سعر_البيع_للعميل"]), step=0.01,
                        key=f"edit_selling_{idx}"
                    )

                # إعادة حساب كل القيم المرتبطة بناءً على التعديل
                new_price_after_discount = itm["السعر_الأساسي"] * (1 - new_purchase_discount / 100)

                if "سيراميك" in itm["النوع"]:
                    type_key = "سيراميك"
                elif "بورسلين" in itm["النوع"]:
                    type_key = "بورسلين"
                else:
                    type_key = "ديكور"

                new_cost_price = new_price_after_discount + fixed_increases.get(type_key, 0.0)
                new_total_meters = itm["سعة_الكرتونة"] * new_carton_count
                new_selling_total = new_selling_price * new_total_meters
                new_cost_total = new_cost_price * new_total_meters
                new_profit_total = new_selling_total - new_cost_total

                # تحديث الصنف نفسه في القائمة (بدون حذف أو إعادة إضافة)
                itm["عدد_الكراتين"] = new_carton_count
                itm["نسبة_خصم_الشراء"] = new_purchase_discount
                itm["السعر_بعد_الخصم_قبل_الزيادة"] = new_price_after_discount
                itm["سعر_التكلفة"] = new_cost_price
                itm["سعر_البيع_للعميل"] = new_selling_price
                itm["عدد_الأمتار"] = new_total_meters
                itm["إجمالي_البيع"] = new_selling_total
                itm["إجمالي_التكلفة"] = new_cost_total
                itm["صافي_الربح"] = new_profit_total

                with col5:
                    st.metric("الإجمالي", f"{new_selling_total:,.2f}")

                with col6:
                    st.write("")
                    if st.button("🗑️", key=f"delete_{idx}", help="حذف هذا الصنف فقط"):
                        items_to_delete.append(idx)

        if items_to_delete:
            for idx in sorted(items_to_delete, reverse=True):
                st.session_state.invoice_items.pop(idx)
            st.rerun()

        invoice_df = pd.DataFrame(st.session_state.invoice_items)

        subtotal = invoice_df["إجمالي_البيع"].sum()
        total_meters_sum = invoice_df["عدد_الأمتار"].sum()
        total_cost_sum = invoice_df["إجمالي_التكلفة"].sum()
        gross_profit = subtotal - total_cost_sum

        st.markdown(f"### 📐 إجمالي عدد الأمتار: {total_meters_sum:.2f} م²")
        st.markdown(f"### 💰 إجمالي الفاتورة قبل خصمها: {subtotal:,.2f} جنيه")

        st.info("💡 لمراجعة تفاصيل التكلفة وصافي الربح، انتقل إلى تبويب '💰 تكاليف الشراء وصافي الربح' أعلى الصفحة.")

        # ---------- خصم الفاتورة والنولون وحساب المديونية ----------
        with st.container(border=True):
            st.markdown("### 💳 خصم الفاتورة والنولون وحساب المديونية")
            col1, col2 = st.columns(2)
            with col1:
                discount_type = st.radio("نوع خصم الفاتورة", ["نسبة %", "قيمة ثابتة"], horizontal=True)
            with col2:
                discount_value_input = st.number_input("قيمة الخصم", min_value=0.0, value=0.0, step=0.01)

            if discount_type == "نسبة %":
                invoice_discount_value = subtotal * (discount_value_input / 100)
            else:
                invoice_discount_value = discount_value_input

            shipping_fee = st.number_input("🚚 النولون (مصاريف الشحن) جنيه", min_value=0.0, value=0.0, step=0.01)

            net_after_discount = (subtotal - invoice_discount_value) + shipping_fee
            total_debt = net_after_discount + previous_balance

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("خصم الفاتورة", f"{invoice_discount_value:,.2f} جنيه")
            with col2:
                st.metric("النولون", f"{shipping_fee:,.2f} جنيه")
            with col3:
                st.metric("صافي هذه الفاتورة", f"{net_after_discount:,.2f} جنيه")
            with col4:
                st.metric("إجمالي المطلوب من العميل", f"{total_debt:,.2f} جنيه")

            st.markdown("---")
            st.markdown("#### 💵 المدفوع الآن / العربون والمتبقي")
            col1, col2 = st.columns(2)
            with col1:
                paid_amount = st.number_input("المدفوع الآن / العربون (جنيه)", min_value=0.0, value=0.0, step=0.01)
            remaining_amount = total_debt - paid_amount
            with col2:
                st.metric("المتبقي المطلوب من العميل", f"{remaining_amount:,.2f} جنيه")

            st.info(amount_in_words(remaining_amount))

            net_profit_after_discount = gross_profit - invoice_discount_value
            st.caption(f"📌 صافي الربح بعد خصم الفاتورة (بدون احتساب تكلفة نولون فعلية): {net_profit_after_discount:,.2f} جنيه")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🗑️ مسح كل الأصناف", use_container_width=True):
                st.session_state.invoice_items = []
                st.rerun()
        with col2:
            if st.button("🔄 تحديث رصيد العميل بعد هذه الفاتورة", use_container_width=True):
                if client_name.strip():
                    customers_db[client_name.strip()] = {
                        "العنوان": client_address,
                        "الهاتف": client_phone,
                        "الرصيد_السابق": remaining_amount
                    }
                    save_customers(customers_db)
                    st.success("تم تحديث رصيد العميل ✅")
                else:
                    st.error("من فضلك أدخل اسم العميل أولاً")
        with col3:
            pdf_bytes = generate_invoice_pdf(
                st.session_state.invoice_items, client_name, client_address, client_phone,
                invoice_date, invoice_number, gallery_name, admin_username,
                subtotal, invoice_discount_value, shipping_fee, net_after_discount,
                previous_balance, total_debt, paid_amount, remaining_amount
            )
            st.download_button(
                label="⬇️ تحميل الفاتورة PDF (نسخة العميل)",
                data=pdf_bytes,
                file_name=f"فاتورة_{invoice_number or 'جديدة'}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

with tab3:
    st.subheader("💰 تكاليف الشراء وصافي الربح (خاص بك فقط - لا يظهر للعميل)")
    st.caption("هذه الصفحة تعرض موقفك المالي الحقيقي مقابل المورد، بعيدًا تمامًا عن شاشة الفاتورة الخاصة بالعميل.")

    if len(st.session_state.invoice_items) == 0:
        st.info("لا يوجد أصناف مضافة بعد. أضف أصنافًا من تبويب '🧾 إنشاء فاتورة / عرض سعر' أولًا.")
    else:
        cost_df = pd.DataFrame(st.session_state.invoice_items)

        internal_columns = ["الصنف", "المقاس", "الفرز", "النوع",
                             "السعر_الأساسي", "نسبة_خصم_الشراء", "السعر_بعد_الخصم_قبل_الزيادة",
                             "سعر_التكلفة", "إجمالي_التكلفة",
                             "سعر_البيع_للعميل", "إجمالي_البيع", "صافي_الربح"]
        st.dataframe(cost_df[internal_columns], use_container_width=True)

        total_cost_sum = cost_df["إجمالي_التكلفة"].sum()
        subtotal_sales = cost_df["إجمالي_البيع"].sum()
        gross_profit = subtotal_sales - total_cost_sum
        total_meters_sum = cost_df["عدد_الأمتار"].sum()

        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📐 إجمالي الأمتار", f"{total_meters_sum:.2f} م²")
        with col2:
            st.metric("💳 إجمالي المستحق للمورد", f"{total_cost_sum:,.2f} جنيه")
        with col3:
            st.metric("🏷️ إجمالي البيع للعميل", f"{subtotal_sales:,.2f} جنيه")
        with col4:
            st.metric("✅ صافي الربح الإجمالي", f"{gross_profit:,.2f} جنيه")

        st.markdown("---")
        st.markdown("#### 🧮 حساب صافي الربح بعد خصم الفاتورة والنولون (اختياري)")
        col1, col2 = st.columns(2)
        with col1:
            manual_discount = st.number_input(
                "خصم الفاتورة المطبق على العميل (جنيه) - لو رصدته في تبويب الفاتورة اكتبه هنا",
                min_value=0.0, value=0.0, step=0.01,
                key="cost_tab_manual_discount"
            )
        with col2:
            manual_shipping_cost = st.number_input(
                "تكلفة النولون الفعلية عليك (جنيه) - لو مختلفة عن اللي حصّلته من العميل",
                min_value=0.0, value=0.0, step=0.01,
                key="cost_tab_manual_shipping"
            )

        net_profit_final = gross_profit - manual_discount - manual_shipping_cost
        st.success(f"✅ صافي ربحك النهائي المتوقع من هذه الفاتورة: **{net_profit_final:,.2f} جنيه**")
