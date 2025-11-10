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
    import math, re, io

    st.subheader("3️⃣ Rezultatai ir EPLAN VBScript generavimas")

    if not excluded:
        st.warning("⚠️ Pirma paspauskite 'Approve'.")
        return

    # ===============================================================
    # 🔹 DUOMENŲ APDOROJIMAS
    # ===============================================================
    df_filtered = df[~df.iloc[:, 0].isin(excluded)].copy()
    rename_map = {
        df_filtered.columns[0]: "Terminalo pavadinimas",
        df_filtered.columns[1]: "Tipas",
        df_filtered.columns[2]: "Jungimo taškas",
        df_filtered.columns[3]: "Matomumas",
        df_filtered.columns[4]: "Grupė"
    }
    df_filtered = df_filtered.rename(columns=rename_map)
    df_filtered["Jungimo taškas"] = df_filtered["Jungimo taškas"].astype(str).str.strip()

    df_filtered = df_filtered.merge(
        term_base[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    grouped = (
        df_filtered.groupby(["Terminalo pavadinimas", "Tipas", "Matomumas",
                             "Grupė", "Plotis (mm)", "Pajungimų skaičius"])
        .agg({"Jungimo taškas": lambda x: sorted(set([v for v in x if v not in ["nan", "None", ""]]))})
        .reset_index()
    )

    def natural_key(v):
        return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(v))]

    def fill_missing_conns(conns, per_terminal):
        if not conns:
            return ""
        conns_sorted = sorted(conns, key=natural_key)
        total_conns = len(conns_sorted)
        total_slots = math.ceil(total_conns / per_terminal) * per_terminal
        out = [conns_sorted[i] if i < len(conns_sorted) else "" for i in range(total_slots)]
        return ", ".join(out)

    grouped["Jungimų seka"] = grouped.apply(
        lambda r: fill_missing_conns(r["Jungimo taškas"], int(r["Pajungimų skaičius"]))
        if pd.notna(r["Pajungimų skaičius"]) and r["Pajungimų skaičius"] > 0 else "",
        axis=1
    )
    grouped["Jungimų kiekis"] = grouped["Jungimo taškas"].apply(len)
    grouped["Terminalų kiekis"] = grouped.apply(
        lambda r: max(1, math.ceil(r["Jungimų kiekis"] / r["Pajungimų skaičius"]))
        if pd.notna(r["Pajungimų skaičius"]) and r["Pajungimų skaičius"] > 0 else 1,
        axis=1
    )

    grouped = grouped.sort_values(by=["Grupė", "Terminalo pavadinimas"])
    display_cols = [
        "Terminalo pavadinimas", "Tipas", "Jungimų seka", "Jungimų kiekis",
        "Pajungimų skaičius", "Terminalų kiekis", "Matomumas", "Grupė", "Plotis (mm)"
    ]
    st.dataframe(grouped[display_cols], use_container_width=True)

    total_terminals = grouped["Terminalų kiekis"].sum()
    st.markdown(f"### 🧮 Iš viso terminalų: **{int(total_terminals)}**")

    # ===============================================================
    # 🧩 VBScript (.vbs) generavimas
    # ===============================================================
    if st.button("🧩 Generuoti EPLAN skriptą (.vbs)"):
        vbs_code = """' ================================================================
' EPLAN Pro Panel – Terminalų automatinis įkėlimas
' Sugeneruota iš Python Streamlit programos
' ================================================================

Option Explicit

Sub Main
    Dim oProject, xlApp, xlBook, xlSheet, row
    Dim termName, termType, connList, connCount, groupCode

    Set oProject = Projects.GetCurrentProject()
    If oProject Is Nothing Then
        MsgBox "❌ Atidarykite projektą prieš paleisdami skriptą!", vbCritical
        Exit Sub
    End If

    Dim xlFile
    xlFile = InputBox("Įveskite Excel failo kelią:", "Terminalų įkėlimas", "C:\\Temp\\terminalai_rezultatas.xlsx")
    If xlFile = "" Then
        MsgBox "Veiksmas nutrauktas – failas nepasirinktas.", vbExclamation
        Exit Sub
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(xlFile)
    Set xlSheet = xlBook.Sheets(1)

    row = 2
    Do While xlSheet.Cells(row, 1).Value <> ""
        termName = Trim(xlSheet.Cells(row, 1).Value)
        termType = Trim(xlSheet.Cells(row, 2).Value)
        connList = Trim(xlSheet.Cells(row, 3).Value)
        connCount = xlSheet.Cells(row, 5).Value
        groupCode = Trim(xlSheet.Cells(row, 8).Value)

        Call AddTerminal(oProject, termName, termType, connList, connCount, groupCode)
        row = row + 1
    Loop

    xlBook.Close False
    xlApp.Quit
    MsgBox "✅ Terminalai sėkmingai importuoti!", vbInformation
End Sub


Sub AddTerminal(oProject, name, tType, conns, connCount, groupNo)
    Dim oFunc
    Set oFunc = New Eplan.EplApi.DataModel.Function(oProject)
    oFunc.Name = name
    oFunc.Properties("20010") = tType
    oFunc.Properties("20013") = connCount
    oFunc.Properties("20220") = groupNo
    oFunc.Generate
End Sub
"""

        vbs_bytes = vbs_code.encode("utf-8")
        st.download_button(
            label="💾 Atsisiųsti VBScript (.vbs)",
            data=vbs_bytes,
            file_name="Import_Terminals_From_List.vbs",
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
