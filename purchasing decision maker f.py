import streamlit as st
import pandas as pd
import math
from datetime import datetime

# ===============================
# 1. 基础配置与数据加载
# ===============================
st.set_page_config(page_title="SADE 采购决策支持系统", layout="wide")

@st.cache_data
def load_all_data():
    try:
        # 1. 加载合同表 (Excel)
        df_contracts = pd.read_excel("contracts_b.xlsx")
        for col in ["DE", "PN", "Price", "MOQ 12ml"]:
            df_contracts[col] = pd.to_numeric(df_contracts[col], errors='coerce')
        
        # 2. 加载运费表 (修改这里：从 read_csv 改为 read_excel)
        # 请确保文件名正确，例如 "Transport PE.xlsx"
        df_transport = pd.read_excel("Transport PE.xlsx") 
        
        # 3. 数据清洗 (Excel 读取后列名和字符串前后常有空格)
        df_transport.columns = df_transport.columns.str.strip()
        df_transport['Dpt'] = df_transport['Dpt'].astype(str).str.strip()
        df_transport['DEPARTEMENTS'] = df_transport['DEPARTEMENTS'].astype(str).str.strip()
        df_transport['Supplier'] = df_transport['Supplier'].astype(str).str.strip()
        
        # 4. 生成省份列表 (用于下拉菜单显示)
        dept_info = df_transport[['Dpt', 'DEPARTEMENTS']].drop_duplicates().sort_values('Dpt')
        dept_list = [f"{row['Dpt']} - {row['DEPARTEMENTS']}" for _, row in dept_info.iterrows()]
        
        return df_contracts, df_transport, dept_list
    except Exception as e:
        st.error(f"加载文件失败，请检查文件是否存在且格式正确: {e}")
        return None, None, []

contracts, transport_db, dept_options_list = load_all_data()

# ===============================
# 2. 你的核心业务规则 (Rules)
# ===============================
def rule_distributor_purchase(quantity, package, DE):
    return (package == "couronne" or DE < 125 or (DE < 200 and quantity < 1200))

def rule_contract_purchase(quantity, package, DE):
    return ((package == "barre" and 125 <= DE <= 200 and 1200 <= quantity)
            or (package == "barre" and 225 <= DE <= 315 and quantity < 2000))

def rule_factory_purchase(quantity, package, DE):
    return ((package == "barre" and 225 <= DE <= 315 and 2000 <= quantity) 
            or package.lower() == "touret" or (package == "barre" and 315 < DE))

def rule_distributor_purchase_dipipe(quantity, DE):
    return (DE < 80)

def rule_contract_purchase_dipipe(quantity, DE):
    # 铸铁管(Fonte)的合同规则
    conditions = [
        (DE >= 300 and quantity <= 264), (DE >= 250 and quantity <= 396),
        (DE >= 200 and quantity <= 440), (DE >= 150 and quantity <= 594),
        (DE >= 125 and quantity <= 770), (DE >= 100 and quantity <= 891),
        (DE >= 80 and quantity <= 968)
    ]
    return any(conditions)

def rule_factory_purchase_dipipe(quantity, DE):
    return not rule_contract_purchase_dipipe(quantity, DE) and DE >= 80

# ===============================
# 3. 价格计算逻辑 (MOQ + Transport)
# ===============================
def calculate_all_totals(material, de, pn, quantity, package, dept_code, today):
    pkg_str = str(package).lower() if package else ""
    mask = (
        (contracts["Material"] == material) &
        (contracts["Valid_Until"] >= today) &
        (contracts["DE"] == float(de)) &
        (contracts["PN"] == float(pn)) &
        (contracts["Package"].astype(str).str.lower() == pkg_str)
    )
    valid_matches = contracts[mask].copy()
    
    # 关键点：如果 MOQ 12ml 为空，说明不符合合同价执行条件
    valid_matches = valid_matches[valid_matches["MOQ 12ml"].notna() & (valid_matches["MOQ 12ml"] > 0)]
    if valid_matches.empty:
        return None

    # 计算车数
    valid_matches["Nb_Camions"] = valid_matches["MOQ 12ml"].apply(lambda x: math.ceil(quantity / x))

    # 获取对应省份和供应商的运费
    def get_fee(supplier):
        fee_m = (transport_db["Supplier"].str.contains(supplier, case=False, na=False)) & (transport_db["Dpt"] == str(dept_code))
        res = transport_db[fee_m]["Transport"]
        return res.iloc[0] if not res.empty else 0

    valid_matches["Transport_Unit"] = valid_matches["Supplier"].apply(get_fee)
    
    # 金额计算
    valid_matches["Material_Total"] = valid_matches["Price"] * quantity
    valid_matches["Total_Transport"] = valid_matches["Nb_Camions"] * valid_matches["Transport_Unit"]
    valid_matches["Grand_Total"] = valid_matches["Material_Total"] + valid_matches["Total_Transport"]

    display_df = valid_matches[["Supplier", "Price", "Nb_Camions", "Transport_Unit", "Total_Transport", "Grand_Total"]].copy()
    display_df.columns = ["Fournisseur", "Unit (€/ml)", "Camions", "Frais/Cam", "Total Trans", "TOTAL HT"]
    
    for col in ["Unit (€/ml)", "Frais/Cam", "Total Trans", "TOTAL HT"]:
        display_df[col] = display_df[col].map("{:,.2f} €".format)
    return display_df.sort_values("TOTAL HT")

# ===============================
# 3.5 邮件草稿生成函数
# ===============================
def generate_email_template(target, material, quantity, de, pn, package, dept):
    subject = f"Demande de prix – {material} DN{de} PN{pn}"
    body = (
        f"Bonjour,\n\n"
        f"Dans le cadre de nos besoins, nous sollicitons votre offre pour :\n\n"
        f"  - Matériau      : {material}\n"
        f"  - Diamètre (DE) : {de} mm\n"
        f"  - Pression (PN) : {pn} bar\n"
        f"  - Conditionnement: {package}\n"
        f"  - Quantité      : {quantity} ml\n\n"
        f"  - Département de livraison : {dept}\n\n" 
        f"Merci de nous transmettre votre meilleur prix dans les meilleurs délais.\n\n"


    )
    return subject, body
# 4. Streamlit UI
# ===============================
st.title("🛡️ SADE Purchasing Decision Support")

if contracts is not None:
    with st.form("purchase_form"):
        col1, col2 = st.columns(2)
        with col1:
            material_choice = st.selectbox("Matériau:", options=[""] + sorted(contracts["Material"].dropna().unique().tolist()))
            package_choice = st.selectbox("Conditionnement:", options=["", "barre", "couronne", "touret"])
            qty_input = st.number_input("Quantité (ml):", min_value=0, step=1)
        with col2:
            de_choice = st.selectbox("DE (Diamètre):", options=[""] + sorted([int(x) for x in contracts["DE"].dropna().unique()]))
            pn_choice = st.selectbox("PN (Pression):", options=[""] + sorted([float(x) for x in contracts["PN"].dropna().unique()]))
            dept_full = st.selectbox("Département de livraison:", options=[""] + dept_options_list)
        
        submit_btn = st.form_submit_button("Run Decision", type="primary")

    if submit_btn:
        if not (material_choice and package_choice and de_choice and pn_choice and dept_full):
            st.warning("⚠️ Veuillez remplir tous les champs.")
        else:
            dept_code = dept_full.split(" - ")[0]
            today = datetime.today()
            
            # --- 初始变量 ---
            decision_msg = ""
            show_prices = False
            target_supplier = "Fournisseur"
            price_table = None

            # --- 1. 执行你原有的判定逻辑 ---
            if "fonte" in material_choice.lower():
                if rule_factory_purchase_dipipe(qty_input, de_choice):
                    decision_msg = "✅ Decision: Consultation Electrosteel sous contrat"
                    show_prices = True
                elif rule_contract_purchase_dipipe(qty_input, de_choice):
                    decision_msg = "✅ Decision: Application tarif contractuel Electrosteel"
                    show_prices = True
                else:
                    decision_msg = "🛒 Decision: Consultation Négoce"
            else:
                if package_choice.lower() == "touret":
                    decision_msg = "✅ Décision: Consultation Elydan (Délai 4-6 sem)"
                    show_prices = True
                elif rule_factory_purchase(qty_input, package_choice, de_choice):
                    decision_msg = "✅ Decision: Consultation Fabricant (Elydan, Centraltubi)"
                    show_prices = True
                elif rule_contract_purchase(qty_input, package_choice, de_choice):
                    decision_msg = "✅ Decision: Application tarif contractuelle"
                    show_prices = True
                else:
                    decision_msg = "🛒 Decision: Consultation Négoce"

            # --- 2. 显示结果 ---
            st.divider()
            st.subheader(decision_msg)

            # --- 3. 如果命中 show_prices，调用新计算函数 ---
            if show_prices:
                price_table = calculate_all_totals(
                    material_choice, 
                    de_choice, 
                    pn_choice, 
                    qty_input, 
                    package_choice, 
                    dept_code, 
                    today
                )
                
                if price_table is not None:
                    st.write("### 💰 Comparatif des prix (Transport inclus)")
                    st.table(price_table)
                    # 如果需要提示用户推荐方案
                    st.success("💡 Le calcul inclut le nombre de camions et les frais de transport par fournisseur.")
                else:
                    # 如果匹配不到价格（比如 MOQ 没填），显示提示
                    st.info("ℹ️ Les tarifs contractuels ne sont pas disponibles pour cette configuration (MOQ non renseignée).")
            # ===============================

            # --- 4. 邮件草稿逻辑 ---
            # --- 4. 邮件草稿逻辑 ---
if not show_prices or price_table is None:
    st.info("📧 **Brouillon d'Email de consultation**")
    
    if "Electrosteel" in decision_msg:
        target = "Electrosteel"
    elif "Elydan" in decision_msg:
        target = "Elydan / Centraltubi"
    else:
        target = "Négoce"

    subject, body = generate_email_template(target, material_choice, qty_input, de_choice, pn_choice, package_choice, dept_choice)
    
    # 显示草稿预览
    st.text_area("Brouillon :", value=body, height=150)

    # ✅ 新增：Outlook 按钮
    import urllib.parse
    mailto_subject = urllib.parse.quote(subject)
    mailto_body = urllib.parse.quote(body)
    mailto_link = f"mailto:?subject={mailto_subject}&body={mailto_body}"
    
    st.link_button(
        label="📨 Ouvrir dans Outlook",
        url=mailto_link,
        type="primary"
    )
