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
    st.subheader("3️⃣ Rezultatai")

    if excluded is None:
        st.warning("⚠️ Pirma patvirtinkite pašalintinus terminalus.")
        return

    # Filtruojam pašalintus terminalus
    df_filtered = df[~df.iloc[:, 0].isin(excluded)].copy()

    # Paruošiam stulpelius
    rename_map = {
        df_filtered.columns[0]: "Terminalo pavadinimas",
        df_filtered.columns[1]: "Tipas",
        df_filtered.columns[2]: "Jungimo taškas",
        df_filtered.columns[3]: "Matomumas" if len(df_filtered.columns) > 3 else "Matomumas",
        df_filtered.columns[4]: "Grupė" if len(df_filtered.columns) > 4 else "Grupė"
    }
    df_filtered = df_filtered.rename(columns=rename_map)

    keep_cols = ["Terminalo pavadinimas", "Tipas", "Jungimo taškas", "Matomumas", "Grupė"]
    df_filtered = df_filtered[[c for c in keep_cols if c in df_filtered.columns]]

    # Pridedam pločio info
    df_filtered = df_filtered.merge(
        terminal_table[["Terminalas", "Plotis (mm)", "Pajungimų skaičius"]],
        how="left", left_on="Tipas", right_on="Terminalas"
    ).drop(columns=["Terminalas"])

    # Grupavimas
    df_filtered["Jungimo taškas"] = df_filtered["Jungimo taškas"].astype(str)
    agg_cols = ["Terminalo pavadinimas", "Tipas", "Matomumas", "Grupė",
                "Plotis (mm)", "Pajungimų skaičius"]
    df_grouped = df_filtered.groupby(agg_cols)["Jungimo taškas"].apply(list).reset_index()

    # Jungimo taškų tekstinis formatas
    def safe_join(x):
        if isinstance(x, list):
            try:
                return ", ".join(map(str, sorted(set(x))))
            except Exception:
                return ", ".join(map(str, x))
        elif pd.isna(x):
            return ""
        else:
            return str(x)

    df_grouped["Jungimo taškas"] = df_grouped["Jungimo taškas"].apply(safe_join)

    # Terminalų kiekio apskaičiavimas
    def count_conns(x):
        return len([v for v in str(x).replace(" ", "").split(",") if v])

    df_grouped["Jungimų kiekis"] = df_grouped["Jungimo taškas"].apply(count_conns)

    # Apskaičiuojam reikalingų terminalų kiekį (ceil)
    import math
    df_grouped["Terminalų kiekis"] = df_grouped.apply(
        lambda r: math.ceil(r["Jungimų kiekis"] / r["Pajungimų skaičius"])
        if pd.notna(r["Pajungimų skaičius"]) and r["Pajungimų skaičius"] > 0 else 0,
        axis=1
    )

    # Rikiavimas
    def min_conn(x):
        try:
            nums = [int(i) for i in str(x).replace(" ", "").split(",") if i.isdigit()]
            return min(nums) if nums else 9999
        except:
            return 9999

    df_grouped["min_conn"] = df_grouped["Jungimo taškas"].apply(min_conn)
    df_grouped = df_grouped.sort_values(by=["Grupė", "min_conn"]).drop(columns="min_conn")

    # Galutinė lentelė su pridėtu kiekiu
    display_cols = [
        "Terminalo pavadinimas", "Tipas", "Jungimo taškas", "Jungimų kiekis",
        "Pajungimų skaičius", "Terminalų kiekis", "Matomumas", "Grupė", "Plotis (mm)"
    ]
    df_final = df_grouped[display_cols]

    st.dataframe(df_final, use_container_width=True)

    # Eksportas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Rezultatas")

    st.download_button(
        "📥 Atsisiųsti rezultatą (Excel)",
        data=output.getvalue(),
        file_name="terminalai_rezultatas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



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
