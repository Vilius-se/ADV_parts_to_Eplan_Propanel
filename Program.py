import streamlit as st
import pandas as pd
import io

# ===============================================================
# 0️⃣ Terminalų bazinė lentelė
# ===============================================================
def load_terminal_base():
    st.subheader("0️⃣ Terminalų duomenų bazė (redaguojama)")
    default_data = pd.DataFrame({
        "Terminalas": ["2002-1301", "2002-1304", "2002-3201", "2002-3207",
                       "2006-8031", "2006-8034", "2016-1201"],
        "Plotis (mm)": [5.2, 5.2, 5.2, 5.2, 9.0, 9.0, 12.0],
        "Pajungimų skaičius": [2, 2, 3, 3, 7, 7, 2]
    })
    edited = st.data_editor(default_data, num_rows="dynamic", key="terminal_base")
    return edited


# ===============================================================
# 1️⃣ Excel įkėlimas
# ===============================================================
def stage1_load_excel():
    st.subheader("1️⃣ Įkelkite Excel failą su terminalų duomenimis")
    uploaded_file = st.file_uploader("Pasirinkite Excel failą", type=["xlsx", "xls"], key="upload")
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Failas įkeltas sėkmingai")
        st.dataframe(df.head())
        return df
    else:
        st.info("Įkelkite Excel failą, kad tęstumėte.")
        return None


# ===============================================================
# 2️⃣ Terminalų pašalinimo pasirinkimas
# ===============================================================
def stage2_exclude_selection(df):
    st.subheader("2️⃣ Pašalintinų terminalų pasirinkimas")

    default_excluded = ["-X0100", "-X0101", "-X0102", "-X111",
                        "-X908", "-X923", "-X927", "-X928", "-XTB10"]
    unique_terms = sorted(df.iloc[:, 0].dropna().unique())

    # --- Naudojam session_state, kad duomenys neišsitrintų ---
    if "excluded" not in st.session_state:
        st.session_state.excluded = None

    selected = []
    with st.form("exclude_form"):
        st.write("Pažymėkite terminalus, kuriuos norite pašalinti:")
        for t in unique_terms:
            checked = t in default_excluded
            if st.checkbox(t, value=checked, key=f"exc_{t}"):
                selected.append(t)
        approved = st.form_submit_button("✅ Approve")

    if approved:
        st.session_state.excluded = selected
        st.success(f"Patvirtinta. Pašalinti terminalai: {', '.join(selected) if selected else 'nėra'}")

    return st.session_state.excluded


def stage3_process_results(df, excluded, terminal_table):
    import math, re, io

    st.subheader("3️⃣ Rezultatai ir EPLAN skripto generavimas")

    if excluded is None:
        st.warning("⚠️ Pirma patvirtinkite pašalintinus terminalus.")
        return

    # === DUOMENŲ PARUOŠIMAS ===
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
        terminal_table[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    grouped = (
        df_filtered.groupby(["Terminalo pavadinimas", "Tipas", "Matomumas", "Grupė",
                             "Plotis (mm)", "Pajungimų skaičius"])
        .agg({"Jungimo taškas": lambda x: sorted(set([v for v in x if v not in ["nan", "None", ""]]))})
        .reset_index()
    )

    def natural_key(value):
        return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(value))]

    def fill_missing_conns(conns, per_terminal):
        if not conns:
            return ""
        conns_sorted = sorted(conns, key=natural_key)
        total_conns = len(conns_sorted)
        total_slots = math.ceil(total_conns / per_terminal) * per_terminal
        filled = []
        for i in range(total_slots):
            filled.append(conns_sorted[i] if i < len(conns_sorted) else "")
        return ", ".join(filled)

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

    # === EPLAN SCRIPT GENERAVIMAS ===
    if st.button("🧩 Generuoti EPLAN skriptą (.vbs)"):
        script = []
        script.append("' ================================================")
        script.append("'  AUTOMATINIS TERMINALŲ ĮKĖLIMO SKRIPTAS EPLAN")
        script.append("'  Sukurta automatiškai iš Streamlit programos")
        script.append("' ================================================")
        script.append("Option Explicit")
        script.append("Sub Main()")
        script.append("  Dim oProject, xlSheet, termName, termType, connList, connCount, groupCode")
        script.append("  Set oProject = Projects.GetCurrentProject()")
        script.append("  If oProject Is Nothing Then")
        script.append("    MsgBox \"Atidarykite projektą prieš paleisdami skriptą!\", vbCritical")
        script.append("    Exit Sub")
        script.append("  End If")
        script.append("  MsgBox \"Terminalų įkėlimas prasideda...\", vbInformation")

        # kiekviena eilutė iš lentelės
        for _, r in grouped.iterrows():
            name = r["Terminalo pavadinimas"]
            typ = r["Tipas"]
            seq = r["Jungimų seka"]
            conn_count = r["Pajungimų skaičius"]
            group = r["Grupė"]
            script.append(f"  Call AddTerminal(oProject, \"{name}\", \"{typ}\", \"{seq}\", {conn_count}, \"{group}\")")

        # papildomos funkcijos
        script.append("  MsgBox \"✅ Visi terminalai įkelti į EPLAN projektą!\", vbInformation")
        script.append("End Sub")
        script.append("")
        script.append("Sub AddTerminal(oProject, name, tType, conns, connCount, groupNo)")
        script.append("  Dim oFunc, arr, i")
        script.append("  Set oFunc = New Eplan.EplApi.DataModel.Function(oProject)")
        script.append("  oFunc.Name = name")
        script.append("  oFunc.Properties(\"20010\") = tType")
        script.append("  oFunc.Properties(\"20013\") = connCount")
        script.append("  oFunc.Properties(\"20220\") = groupNo")
        script.append("  arr = Split(conns, \",\")")
        script.append("  For i = LBound(arr) To UBound(arr)")
        script.append("    If Trim(arr(i)) <> \"\" Then")
        script.append("      oFunc.Properties(\"20014\") = Trim(arr(i))")
        script.append("    End If")
        script.append("  Next")
        script.append("  oFunc.Generate()")
        script.append("End Sub")

        vbs_content = "\n".join(script)
        vbs_bytes = vbs_content.encode("utf-8")

        st.download_button(
            label="💾 Atsisiųsti EPLAN skriptą",
            data=vbs_bytes,
            file_name="Import_Terminals_From_List.vbs",
            mime="text/plain"
        )

    # Bendras terminalų kiekis
    total_terminals = grouped["Terminalų kiekis"].sum()
    st.markdown(f"### 🧮 Viso terminalų: **{int(total_terminals)}**")



# ===============================================================
# 🔁 MAIN PIPELINE
# ===============================================================
def main():
    st.title("🔌 Terminalų apdorojimo pipeline")

    terminal_table = load_terminal_base()
    df = stage1_load_excel()

    if df is not None:
        excluded = stage2_exclude_selection(df)
        stage3_process_results(df, excluded, terminal_table)


if __name__ == "__main__":
    main()
