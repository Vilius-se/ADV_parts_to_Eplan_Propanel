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
    import pandas as pd

    st.subheader("3️⃣ Rezultatai ir EPLAN 2025 VB.NET skripto generavimas")

    # ---------------------------------------------------------------
    # 1️⃣ Patikrinimas
    # ---------------------------------------------------------------
    if not excluded:
        st.warning("⚠️ Pirma paspauskite 'Approve'.")
        return

    # ---------------------------------------------------------------
    # 2️⃣ Duomenų paruošimas
    # ---------------------------------------------------------------
    df_filtered = df[~df.iloc[:, 0].isin(excluded)].copy()
    rename_map = {
        df_filtered.columns[0]: "Terminalo pavadinimas",
        df_filtered.columns[1]: "Tipas",
        df_filtered.columns[2]: "Jungimo taškas",
        df_filtered.columns[3]: "Matomumas",
        df_filtered.columns[4]: "Grupė"
    }
    df_filtered = df_filtered.rename(columns=rename_map)
    df_filtered["Jungimo taškas"] = df_filtered["Jungimo taškas"].astype(str)

    # Prijungiam terminalų bazę (plotis, pajungimų sk.)
    df_filtered = df_filtered.merge(
        term_base[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    # ---------------------------------------------------------------
    # 3️⃣ Grupavimas ir jungčių sutraukimas
    # ---------------------------------------------------------------
    grouped = (
        df_filtered.groupby(
            ["Terminalo pavadinimas", "Tipas", "Matomumas",
             "Grupė", "Plotis (mm)", "Pajungimų skaičius"]
        )
        .agg({
            "Jungimo taškas": lambda x: sorted(
                set(
                    str(v).strip()
                    for v in x
                    if pd.notna(v) and str(v).strip() not in ["", "nan", "None"]
                )
            )
        })
        .reset_index()
    )

    # Helper funkcijos
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
        "Terminalo pavadinimas", "Tipas", "Jungimų seka",
        "Jungimų kiekis", "Pajungimų skaičius", "Terminalų kiekis",
        "Matomumas", "Grupė", "Plotis (mm)"
    ]
    st.dataframe(grouped[display_cols], use_container_width=True)

    total_terminals = grouped["Terminalų kiekis"].sum()
    st.markdown(f"### 🧮 Iš viso terminalų: **{int(total_terminals)}**")

    # ---------------------------------------------------------------
    # 4️⃣ VB.NET skripto (EPLAN 2025) generavimas
    # ---------------------------------------------------------------
    if st.button("💻 Generuoti EPLAN 2025 VB.NET skriptą (.vb)"):
        vb_code = """' ================================================================
' EPLAN 2025 – Terminalų įkėlimas iš sąrašo (modernus API)
' Sugeneruota Streamlit programos
' ================================================================
Imports System.IO
Imports System.Windows.Forms
Imports Eplan.EplApi.Scripting
Imports Eplan.EplApi.ApplicationFramework

Public Class Import_Terminals_2025

    <Start>
    Public Sub Main()
        Try
            Dim excelPath As String = InputBox("Įveskite Excel failo kelią:", "Importuoti terminalus", "C:\\Temp\\terminalai_rezultatas.xlsx")
            If excelPath = "" Then
                MessageBox.Show("Veiksmas nutrauktas – failas nepasirinktas.")
                Exit Sub
            End If

            Dim xlApp As Object = CreateObject("Excel.Application")
            xlApp.Visible = False
            Dim xlBook As Object = xlApp.Workbooks.Open(excelPath)
            Dim xlSheet As Object = xlBook.Sheets(1)

            Dim row As Integer = 2
            Do While xlSheet.Cells(row, 1).Value <> ""
                Dim name As String = CStr(xlSheet.Cells(row, 1).Value)
                Dim tType As String = CStr(xlSheet.Cells(row, 2).Value)
                Dim groupCode As String = CStr(xlSheet.Cells(row, 5).Value)

                Call CreateTerminal(name, tType, groupCode)
                row += 1
            Loop

            xlBook.Close(False)
            xlApp.Quit()
            MessageBox.Show("✅ Terminalai importuoti sėkmingai!", "EPLAN Script")

        Catch ex As Exception
            MessageBox.Show("❌ Klaida: " & ex.Message)
        End Try
    End Sub

    Private Sub CreateTerminal(name As String, tType As String, groupCode As String)
        Dim actSvc As New ActionService()
        Dim actionName As String = "XEsCreateDevice"
        Dim actionParams As String = "Name:" & name & ",Type:" & tType & ",FunctionDefinition:Terminal"
        actSvc.Execute(actionName, actionParams)
    End Sub

End Class
"""

        vb_bytes = vb_code.encode("utf-8")
        st.download_button(
            label="📦 Atsisiųsti EPLAN 2025 VB.NET skriptą",
            data=vb_bytes,
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
