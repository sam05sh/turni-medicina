import streamlit as st
import pandas as pd
from itertools import cycle
from io import BytesIO

st.set_page_config(page_title="Gestione Turni Sbobinature", layout="centered")

st.title("📚 App per Gestione Turni Sbobinature")
st.markdown(
    "Carica i file Excel per generare automaticamente il file con i turni."
)

# UI for deadline days
days_to_add = st.number_input(
    "Giorni da aggiungere per la Scadenza Consigliata:",
    min_value=0,
    max_value=30,
    value=3,
    step=1,
)

# File uploaders
schedule_file = st.file_uploader(
    "1. Carica il file Orario (Excel)", type=["xlsx", "xls"]
)
roster_file = st.file_uploader(
    "2. Carica il file Studenti (Excel)", type=["xlsx", "xls"]
)

if schedule_file and roster_file:
    if st.button("Genera Turni"):
        try:
            # 1. Read the Schedule (Headers are on row 6, index 5)
            df_schedule = pd.read_excel(schedule_file, header=5)
            # Drop rows where the 'Giorno' or 'Insegnamento' is completely missing
            df_schedule.dropna(subset=["Giorno", "Insegnamento"], inplace=True)

            # 2. Read the Students roster and sort A-Z
            df_roster = pd.read_excel(roster_file)
            col_name = df_roster.columns[0]  # Get the first column dynamically
            # Drop empty names, convert to string, and sort A-Z
            students = df_roster[col_name].dropna().astype(str).tolist()
            students = sorted([s.strip() for s in students])

            # Create an infinite round-robin loop over the students
            student_pool = cycle(students)

            # 3. Process N. Lezione (Cumulative count per subject)
            df_schedule["N. Lezione"] = (
                df_schedule.groupby("Insegnamento").cumcount() + 1
            )

            # 4. Process Dates and Deadlines
            # Ensure Giorno is parsed as a datetime object
            df_schedule["Giorno"] = pd.to_datetime(
                df_schedule["Giorno"], errors="coerce"
            )
            df_schedule["Scadenza consigliata"] = df_schedule[
                "Giorno"
            ] + pd.Timedelta(days=days_to_add)

            # 5. Assign Roles
            sbobinatore_1 = []
            sbobinatore_2 = []
            controllore = []
            super_controllore = []

            for _ in range(len(df_schedule)):
                sbobinatore_1.append(next(student_pool))
                sbobinatore_2.append(next(student_pool))
                controllore.append(next(student_pool))
                super_controllore.append(next(student_pool))

            # 6. Build Output DataFrame
            df_output = pd.DataFrame(
                {
                    "Data": df_schedule["Giorno"].dt.strftime("%d/%m/%Y"),
                    "Materia": df_schedule["Insegnamento"],
                    "N. Lezione": df_schedule["N. Lezione"],
                    "Orario": df_schedule["Ora"],
                    "Sbobinatore 1": sbobinatore_1,
                    "Sbobinatore 2": sbobinatore_2,
                    "Controllore": controllore,
                    "Super controllore": super_controllore,
                    "Scadenza consigliata": df_schedule[
                        "Scadenza consigliata"
                    ].dt.strftime("%d/%m/%Y"),
                    "Consegnato (vero o falso)": "Falso",
                }
            )

            # 7. Convert Output to Excel in memory for download
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_output.to_excel(writer, index=False, sheet_name="Turni")
            processed_data = output.getvalue()

            st.success("✅ File generato con successo!")

            # Download button
            st.download_button(
                label="📥 Scarica il file Excel dei Turni",
                data=processed_data,
                file_name="Turni_Sbobinature.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
            )

        except Exception as e:
            st.error(f"Si è verificato un errore durante l'elaborazione: {e}")