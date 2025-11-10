import streamlit as st
import pandas as pd
import io, math, re

st.set_page_config(page_title="EPLAN Terminal Generator", layout="wide")


# ===============================================================
# 0️⃣ Terminalų bazė (redaguojama)
# ===============================================================
def load_terminal_base():
    st.subheader("0️⃣ Terminalų bazė")
    default_data = pd.DataFrame({
        "Terminalas": ["2002-1301", "2002-1304", "2002-3201", "2002-3207",
                       "2006-8031", "2006-8034", "2016-1201"],
        "Plotis (mm)": [5.2, 5.2, 5.2, 5.2, 9.0, 9.0, 12.0],
        "Pajungimų skaičius": [2, 2, 3, 3, 7, 7, 2]
    })
    return st.data_editor(default_data, num_rows="dynamic", key="term_base")


# ===============================================================
# 1️⃣ Excel įkėlimas
# ===============================================================
def stage1_load_excel():
    st.subheader("1️⃣ Įkelkite Excel failą")
    file = st.file_uploader("Pasirinkite Excel failą", type=["xlsx", "xls"])
    if not file:
        st.info("Įkelkite failą, kad tęstumėte.")
        return None
    df = pd.read_excel(file)
    st.success("✅ Failas įkeltas")
    st.dataframe(df.head())
    return df


# ===============================================================
# 2️⃣ Terminalų išskyrimas
# ===============================================================
def stage2_exclude_selection(df):
    st.subheader("2️⃣ Terminalų išskyrimas (Exclude)")
    default_excluded = ["-X0100", "-X0101", "-X0102", "-X111", "-X908",
                        "-X923", "-X927", "-X928", "-XTB10"]

    unique_terms = sorted(df.iloc[:, 0].dropna().unique())
    if "excluded" not in st.session_state:
        st.session_state.excluded = default_excluded

    with st.form("exclude_form"):
        selected = []
        for t in unique_terms:
            checked = t in st.session_state.excluded
            if st.checkbox(t, value=checked, key=f"exc_{t}"):
                selected.append(t)
        approved = st.form_submit_button("✅ Approve")

    if approved:
        st.session_state.excluded = selected
        st.success(f"Patvirtinta: {len(selected)} terminalų išskirta.")
    return st.session_state.excluded


# ===============================================================
# 3️⃣ Rezultatai + VB.NET skripto generavimas
# ===============================================================
def stage3_process_results(df, excluded, term_base):
    import math, re, io, pandas as pd

    st.subheader("3️⃣ Rezultatai ir EPLAN 2025 VB.NET skripto generavimas")

    if not excluded:
        st.warning("⚠️ Pirma patvirtinkite pašalintinus terminalus.")
        return

    # --- DUOMENŲ PARUOŠIMAS ---
    df_filtered = df[~df.iloc[:, 0].isin(excluded)].copy()
    rename_map = {
        df_filtered.columns[0]: "Terminalo pavadinimas",
        df_filtered.columns[1]: "Tipas",
        df_filtered.columns[2]: "Jungimo taškas",
        df_filtered.columns[3]: "Matomumas",
        df_filtered.columns[4]: "Grupė"
    }
    df_filtered = df_filtered.rename(columns=rename_map)

    df_filtered = df_filtered.merge(
        term_base[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    grouped = (
        df_filtered.groupby(["Terminalo pavadinimas", "Tipas", "Matomumas", "Grupė"])
        .agg({"Jungimo taškas": lambda x: sorted(set(str(v).strip() for v in x if pd.notna(v) and str(v).strip() not in ["", "nan", "None"]))})
        .reset_index()
    )

    grouped = grouped.sort_values(by=["Grupė", "Terminalo pavadinimas"])
    st.dataframe(grouped, use_container_width=True)

    # --- VB.NET SKRIPTO GENERAVIMAS (be InputBox, be Excel, be CreateObject) ---
    if st.button("💻 Generuoti EPLAN 2025 VB.NET skriptą (.vb)"):
    vb_code = """' ================================================================
' EPLAN 2025 – Terminalų automatinis įkėlimas (sugeneruota iš Streamlit)
' ================================================================
Imports Eplan.EplApi.Scripting
Imports Eplan.EplApi.ApplicationFramework
Imports System.Windows.Forms

Public Class Import_Terminals_2025

    <Start>
    Public Sub Main()
        Try
            Dim actMgr As New ActionManager()
            Dim eplanAction As Eplan.EplApi.ApplicationFramework.Action = actMgr.GetAction("XEsCreateDevice")

"""
    # --- automatinis terminalų sąrašo įrašymas ---
    for _, r in grouped.iterrows():
        name = str(r["Terminalo pavadinimas"]).replace('"', "'")
        ttype = str(r["Tipas"]).replace('"', "'")
        group = str(r["Grupė"]).replace('"', "'")
        vb_code += f'            eplanAction.Execute("Name:{name},Type:{ttype},FunctionDefinition:Terminal,MountingLocation:{group}")\n'

    vb_code += """
            MessageBox.Show("✅ Terminalai sėkmingai įkelti į projektą!", "EPLAN Script", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("❌ Klaida: " & ex.Message, "EPLAN Script", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
"""

    st.download_button(
        label="📦 Atsisiųsti EPLAN 2025 VB.NET skriptą",
        data=vb_code.encode("utf-8"),
        file_name="Import_Terminals_2025.vb",
        mime="text/plain"
    )



# ===============================================================
# MAIN PIPELINE
# ===============================================================
def main():
    st.title("⚙️ EPLAN Terminalų Generatorius")

    term_base = load_terminal_base()
    df = stage1_load_excel()
    if df is not None:
        excluded = stage2_exclude_selection(df)
        stage3_process_results(df, excluded, term_base)


if __name__ == "__main__":
    main()
