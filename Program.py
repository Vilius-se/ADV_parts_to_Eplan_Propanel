import streamlit as st
import pandas as pd
import io

# ===============================================================
# 0. Pradinė terminalų bazė (vartotojo redaguojama)
# ===============================================================

def load_terminal_base():
    st.subheader("0️⃣ Terminalų duomenų bazė")
    default_data = pd.DataFrame({
        "Terminalas": ["2002-1301", "2002-1304", "2002-3201", "2002-3207", "2006-8031", "2006-8034", "2016-1201"],
        "Plotis (mm)": [5.2, 5.2, 5.2, 5.2, 9.0, 9.0, 12.0],
        "Pajungimų skaičius": [2, 2, 3, 3, 7, 7, 2]
    })
    edited = st.data_editor(default_data, num_rows="dynamic", key="terminal_base")
    return edited


# ===============================================================
# 1. Excel įkėlimas
# ===============================================================

def stage1_load_excel():
    st.subheader("1️⃣ Įkelkite Excel failą su terminalų duomenimis")
    uploaded_file = st.file_uploader("Pasirinkite Excel failą", type=["xlsx", "xls"], key="upload")
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.success("Failas įkeltas sėkmingai ✅")
        st.dataframe(df.head())
        return df
    else:
        st.info("Įkelkite Excel failą, kad tęstumėte.")
        return None


# ===============================================================
# 2. Pašalintinų terminalų pasirinkimas
# ===============================================================

def stage2_exclude_selection(df):
    st.subheader("2️⃣ Pasirinkite terminalus, kuriuos reikia išskirti")

    default_excluded = ["-X0100","-X0101","-X0102","-X111","-X908","-X923","-X927","-X928","-XTB10"]
    unique_terms = sorted(df.iloc[:, 0].dropna().unique())

    selected = []
    with st.form("exclude_form"):
        st.write("Pažymėkite terminalus, kuriuos norite pašalinti:")
        for t in unique_terms:
            checked = t in default_excluded
            if st.checkbox(t, value=checked, key=t):
                selected.append(t)
        approved = st.form_submit_button("✅ Approve")

    if approved:
        st.success(f"Patvirtinta. Pašalinti terminalai: {', '.join(selected) if selected else 'Nėra'}")
        return selected
    return None


# ===============================================================
# 3. Duomenų apdorojimas ir rezultato lentelė
# ===============================================================

def stage3_process_results(df, excluded, terminal_table):
    st.subheader("3️⃣ Rezultatai")

    if excluded is None:
        st.warning("Pirma patvirtinkite pašalintinus terminalus.")
        return

    # Filtruojam
    df_filtered = df[~df.iloc[:, 0].isin(excluded)].copy()
    df_filtered.columns = ["Terminalo pavadinimas", "Tipas", "Jungimo taškas", "Matomumas", "Grupė"]

    # Pridedam plotį pagal tipą
    df_filtered = df_filtered.merge(
        terminal_table[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    # Grupavimas: sujungiame jungimo taškus
    agg_cols = ["Terminalo pavadinimas", "Tipas", "Matomumas", "Grupė", "Plotis (mm)", "Pajungimų skaičius"]
    df_grouped = df_filtered.groupby(agg_cols)["Jungimo taškas"].apply(list).reset_index()
    df_grouped["Jungimo taškas"] = df_grouped["Jungimo taškas"].apply(lambda x: ", ".join(map(str, sorted(x))))

    # Rikiavimas pagal grupę ir jungimo taškus
    def min_conn(x):
        try:
            return min(map(int, str(x).replace(" ", "").split(",")))
        except:
            return 9999
    df_grouped["min_conn"] = df_grouped["Jungimo taškas"].apply(min_conn)
    df_grouped = df_grouped.sort_values(by=["Grupė", "min_conn"]).drop(columns="min_conn")

    st.dataframe(df_grouped)

    # Parsisiuntimo mygtukas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_grouped.to_excel(writer, index=False, sheet_name="Rezultatas")
    st.download_button("📥 Atsisiųsti rezultatą (Excel)", data=output.getvalue(),
                       file_name="terminalai_rezultatas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ===============================================================
# MAIN PIPELINE
# ===============================================================

def main():
    st.title("🔌 Terminalų apdorojimo pipeline")

    # Stage 0: bazinė lentelė
    terminal_table = load_terminal_base()

    # Stage 1
    df = stage1_load_excel()
    if df is not None:
        # Stage 2
        excluded = stage2_exclude_selection(df)
        # Stage 3
        stage3_process_results(df, excluded, terminal_table)


if __name__ == "__main__":
    main()
