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

    selected = []
    with st.form("exclude_form"):
        st.write("Pažymėkite terminalus, kuriuos norite pašalinti:")
        for t in unique_terms:
            checked = t in default_excluded
            if st.checkbox(t, value=checked, key=f"exc_{t}"):
                selected.append(t)
        approved = st.form_submit_button("✅ Approve")

    if approved:
        st.success(f"Patvirtinta. Pašalinti terminalai: {', '.join(selected) if selected else 'nėra'}")
        return selected
    return None


def stage3_process_results(df, excluded, terminal_table):
    import math

    st.subheader("3️⃣ Rezultatai")

    if excluded is None:
        st.warning("⚠️ Pirma patvirtinkite pašalintinus terminalus.")
        return

    # Filtruojam pašalintus terminalus
    df_filtered = df[~df.iloc[:, 0].isin(excluded)].copy()

    # Aiškūs pavadinimai
    rename_map = {
        df_filtered.columns[0]: "Terminalo pavadinimas",
        df_filtered.columns[1]: "Tipas",
        df_filtered.columns[2]: "Jungimo taškas",
        df_filtered.columns[3]: "Matomumas",
        df_filtered.columns[4]: "Grupė"
    }
    df_filtered = df_filtered.rename(columns=rename_map)

    # Paverčiam jungimo taškus į skaičius
    df_filtered["Jungimo taškas"] = pd.to_numeric(df_filtered["Jungimo taškas"], errors="coerce")

    # Pridedam informaciją iš terminalų bazės
    df_filtered = df_filtered.merge(
        terminal_table[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    # Grupavimas pagal terminalą
    grouped = (
        df_filtered.groupby(["Terminalo pavadinimas", "Tipas", "Matomumas", "Grupė",
                             "Plotis (mm)", "Pajungimų skaičius"])
        .agg({"Jungimo taškas": lambda x: sorted(set([v for v in x if pd.notna(v)]))})
        .reset_index()
    )

    # Funkcija jungimo sąrašo užpildymui
    def fill_missing_conns(conns, per_terminal):
        """Papildo sąrašą tuščiomis reikšmėmis iki artimiausio pilno modulo."""
        if not conns:
            return ""
        max_conn = int(max(conns))
        total_positions = math.ceil(max_conn / per_terminal) * per_terminal
        filled = []
        all_positions = list(range(1, total_positions + 1))
        for pos in all_positions:
            filled.append(str(int(pos)) if pos in conns else "")
        return ", ".join(filled)

    # Pridėkime jungimo sekas ir apskaičiuokime terminalų kiekį
    grouped["Jungimų seka"] = grouped.apply(
        lambda r: fill_missing_conns(r["Jungimo taškas"], int(r["Pajungimų skaičius"])) 
        if pd.notna(r["Pajungimų skaičius"]) else "",
        axis=1
    )

    # Didžiausias jungimas
    grouped["Didžiausias jungimas"] = grouped["Jungimo taškas"].apply(lambda x: max(x) if x else 0)

    # Kiek terminalų
    grouped["Terminalų kiekis"] = grouped.apply(
        lambda r: max(1, math.ceil(r["Didžiausias jungimas"] / r["Pajungimų skaičius"]))
        if pd.notna(r["Didžiausias jungimas"]) and pd.notna(r["Pajungimų skaičius"]) else 1,
        axis=1
    )

    # Rikiavimas
    grouped = grouped.sort_values(by=["Grupė", "Terminalo pavadinimas"])

    # Lentelės atvaizdavimas
    display_cols = [
        "Terminalo pavadinimas", "Tipas", "Jungimų seka",
        "Pajungimų skaičius", "Terminalų kiekis", "Matomumas", "Grupė", "Plotis (mm)"
    ]
    st.dataframe(grouped[display_cols], use_container_width=True)

    # Eksportas į Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, index=False, sheet_name="Rezultatas")

    st.download_button(
        "📥 Atsisiųsti rezultatą (Excel)",
        data=output.getvalue(),
        file_name="terminalai_rezultatas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Suminis terminalų kiekis
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
