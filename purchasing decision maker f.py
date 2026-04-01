import streamlit as st
import pandas as pd
import math
from datetime import datetime
import urllib.parse

# ===============================
# 1. 基础配置与数据加载
# ===============================
st.set_page_config(page_title="SADE 采购决策支持系统", layout="wide")

@st.cache_data
def load_all_data():
    try:
        df_contracts = pd.read_excel("contracts_b.xlsx")
        for col in ["DE", "PN", "Price", "MOQ 12ml"]:
            df_contracts[col] = pd.to_numeric(df_contracts[col], errors='coerce')

        # ✅ NEW: Load negoce contracts
        df_negoce = pd.read_excel("contracts_negoce.xlsx")
        for col in ["DE", "PN", "Price"]:
            if col in df_negoce.columns:
                df_negoce[col] = pd.to_numeric(df_negoce[col], errors='coerce')

        df_transport = pd.read_excel("Transport PE.xlsx")
        df_transport.columns = df_transport.columns.str.strip()
        df_transport['Dpt'] = df_transport['Dpt'].astype(str).str.strip()
        df_transport['DEPARTEMENTS'] = df_transport['DEPARTEMENTS'].astype(str).str.strip()
        df_transport['Supplier'] = df_transport['Supplier'].astype(str).str.strip()

        dept_info = df_transport[['Dpt', 'DEPARTEMENTS']].drop_duplicates().sort_values('Dpt')
        dept_list = [f"{row['Dpt']} - {row['DEPARTEMENTS']}" for _, row in dept_info.iterrows()]

        return df_contracts, df_negoce, df_transport, dept_list
    except Exception as e:
        st.error(f"Erreur de chargement des fichiers : {e}")
        return None, None, None, []

contracts, negoce_contracts, transport_db, dept_options_list = load_all_data()


# ===============================
# 2. Business Rules (unchanged)
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
# 3. Price Calculation (unchanged)
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

    if valid_matches.empty:
        return None

    has_moq = valid_matches["MOQ 12ml"].notna() & (valid_matches["MOQ 12ml"] > 0)
    matches_with_moq = valid_matches[has_moq].copy()
    matches_without_moq = valid_matches[~has_moq].copy()

    def get_fee(supplier):
        fee_m = (transport_db["Supplier"].str.contains(supplier, case=False, na=False)) & (transport_db["Dpt"] == str(dept_code))
        res = transport_db[fee_m]["Transport"]
        return res.iloc[0] if not res.empty else 0

    results = []

    if not matches_with_moq.empty:
        matches_with_moq["Nb_Camions"] = matches_with_moq["MOQ 12ml"].apply(lambda x: math.ceil(quantity / x))
        matches_with_moq["Transport_Unit"] = matches_with_moq["Supplier"].apply(get_fee)
        matches_with_moq["Material_Total"] = matches_with_moq["Price"] * quantity
        matches_with_moq["Total_Transport"] = matches_with_moq["Nb_Camions"] * matches_with_moq["Transport_Unit"]
        matches_with_moq["Grand_Total"] = matches_with_moq["Material_Total"] + matches_with_moq["Total_Transport"]
        matches_with_moq["Camions"] = matches_with_moq["Nb_Camions"]
        matches_with_moq["Frais/Cam"] = matches_with_moq["Transport_Unit"].map("{:,.2f} €".format)
        matches_with_moq["Total Trans"] = matches_with_moq["Total_Transport"].map("{:,.2f} €".format)
        matches_with_moq["TOTAL HT"] = matches_with_moq["Grand_Total"].map("{:,.2f} €".format)
        results.append(matches_with_moq[["Supplier", "Price", "Camions", "Frais/Cam", "Total Trans", "TOTAL HT"]])

    if not matches_without_moq.empty:
        matches_without_moq["Price"] = pd.to_numeric(matches_without_moq["Price"], errors="coerce")
        matches_without_moq["Material_Total"] = matches_without_moq["Price"] * quantity
        matches_without_moq["Camions"] = "-"
        matches_without_moq["Frais/Cam"] = "-"
        matches_without_moq["Total Trans"] = "-"
        matches_without_moq["TOTAL HT"] = matches_without_moq["Material_Total"].map("{:,.2f} €".format)
        results.append(matches_without_moq[["Supplier", "Price", "Camions", "Frais/Cam", "Total Trans", "TOTAL HT"]])

    if not results:
        return None

    display_df = pd.concat(results, ignore_index=True)
    display_df.columns = ["Fournisseur", "Unit (€/ml)", "Camions", "Frais/Cam", "Total Trans", "TOTAL HT"]
    display_df["Unit (€/ml)"] = display_df["Unit (€/ml)"].map("{:,.2f} €".format)
    return display_df.sort_values("TOTAL HT").reset_index(drop=True)


# ===============================
# ✅ NEW: Negoce Price Lookup
# ===============================
def get_negoce_prices(material, de, pn, quantity, today):
    if negoce_contracts is None or negoce_contracts.empty:
        return None

    mask = (
        (negoce_contracts["Material"] == material) &
        (negoce_contracts["DE"] == float(de)) &
        (negoce_contracts["PN"] == float(pn))
    )

    if "Valid_Until" in negoce_contracts.columns:
        mask &= (negoce_contracts["Valid_Until"] >= today)

    matches = negoce_contracts[mask].copy()

    if matches.empty:
        return None

    matches["Prix Total HT"] = (matches["Price"] * quantity).map("{:,.2f} €".format)
    matches["Prix Unit (€/ml)"] = matches["Price"].map("{:,.4f} €".format)

    result = pd.DataFrame()
    if "Supplier" in matches.columns:
        result["Négoce"] = matches["Supplier"].values
    if "ml par unit" in matches.columns:
        result["ML par unité"] = matches["ml par unit"].values   # ✅ nouvelle colonne
    result["Prix Unit (€/ml)"] = matches["Prix Unit (€/ml)"].values
    result["Prix Total HT"] = matches["Prix Total HT"].values

    return result.reset_index(drop=True)

    # Show relevant columns — adjust if your file has different column names
    display_cols = []
    for col in ["Supplier", "Prix Unit (€/ml)", "Prix Total HT"]:
        if col in matches.columns or col in ["Prix Unit (€/ml)", "Prix Total HT"]:
            display_cols.append(col)

    # Build clean display table
    result = pd.DataFrame()
    if "Supplier" in matches.columns:
        result["Négoce"] = matches["Supplier"].values
    result["Prix Unit (€/ml)"] = matches["Prix Unit (€/ml)"].values
    result["Prix Total HT"] = matches["Prix Total HT"].values

    return result.reset_index(drop=True)


# ===============================
# 3.5 Email Draft
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


# ===============================
# 4. Streamlit UI
# ===============================
st.title("🛡️ SADE Purchasing Decision Support")

if contracts is not None:
    show_prices = False
    price_table = None
    decision_msg = ""
    is_negoce = False  # ✅ NEW flag

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

            decision_msg = ""
            show_prices = False
            is_negoce = False  # ✅ reset
            price_table = None

            # --- Decision logic ---
            if "fonte" in material_choice.lower():
                if rule_factory_purchase_dipipe(qty_input, de_choice):
                    decision_msg = "✅ Decision: Consultation Electrosteel sous contrat"
                    show_prices = True
                elif rule_contract_purchase_dipipe(qty_input, de_choice):
                    decision_msg = "✅ Decision: Application tarif contractuel Electrosteel"
                    show_prices = True
                else:
                    decision_msg = "🛒 Decision: Consultation Négoce"
                    is_negoce = True  # ✅
            else:
                if package_choice.lower() == "touret":
                    decision_msg = "✅ Décision: Consultation Elydan"
                    show_prices = True
                elif rule_factory_purchase(qty_input, package_choice, de_choice):
                    decision_msg = "✅ Decision: Consultation Fabricant (Elydan, Centraltubi)"
                    show_prices = True
                elif rule_contract_purchase(qty_input, package_choice, de_choice):
                    decision_msg = "✅ Decision: Application tarif contractuelle"
                    show_prices = True
                else:
                    decision_msg = "🛒 Decision: Consultation Négoce"
                    is_negoce = True  # ✅

            # --- Display decision ---
            st.divider()
            st.subheader(decision_msg)

            # --- Contract prices ---
            if show_prices:
                price_table = calculate_all_totals(
                    material_choice, de_choice, pn_choice,
                    qty_input, package_choice, dept_code, today
                )

            if price_table is not None:
                if "-" in price_table["Frais/Cam"].values:
                    st.write("### 💰 Comparatif des prix (Transport non inclus)")
                    st.dataframe(price_table, hide_index=True, use_container_width=True)
                    st.warning("💡 Le calcul n'inclut pas de frais de transport par fournisseur.")
                else:
                    st.write("### 💰 Comparatif des prix (Transport inclus)")
                    st.dataframe(price_table, hide_index=True, use_container_width=True)
                    st.success("💡 Le calcul inclut le nombre de camions et les frais de transport.")

            # ✅ NEW: Negoce price table
            if is_negoce:
                negoce_table = get_negoce_prices(
                    material_choice, de_choice, pn_choice, qty_input, today
                )
                if negoce_table is not None:
                    st.write("### 🏪 Prix de référence Négoce")
                    st.dataframe(negoce_table, hide_index=True, use_container_width=True)
                    st.info("💡 Ces prix sont issus du fichier de référence négoce. Vérifiez la disponibilité auprès du fournisseur.")
                else:
                    st.warning("⚠️ Aucun prix négoce trouvé pour cette référence. Merci de contacter le négoce directement.")

            # --- Email draft ---
            if not show_prices or price_table is None or package_choice.lower() == "touret":
                st.info("📧 **Brouillon d'Email de consultation**")

                if "Electrosteel" in decision_msg:
                    target = "Electrosteel"
                elif "Elydan" in decision_msg:
                    target = "Elydan / Centraltubi"
                else:
                    target = "Négoce"

                subject, body = generate_email_template(
                    target, material_choice, qty_input,
                    de_choice, pn_choice, package_choice, dept_full
                )
                st.text_area("Brouillon :", value=body, height=150)

                mailto_subject = urllib.parse.quote(subject)
                mailto_body = urllib.parse.quote(body)
                mailto_link = f"mailto:?subject={mailto_subject}&body={mailto_body}"

                st.link_button("📨 Ouvrir dans Outlook", url=mailto_link, type="primary")
