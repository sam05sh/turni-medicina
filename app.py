import streamlit as st
import pandas as pd
import re
from itertools import cycle
from io import BytesIO

st.set_page_config(page_title="Gestione Turni Sbobinature", layout="centered")

st.title("Gestione Turni Sbobinature")
st.markdown(
    "GUIDA VELOCE:", 
    "ciao"
)

# UI for deadline days
days_to_add = st.number_input(
    "Scadenza consegna sbobina in giorni post lezione:",
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

# Helper function to extract the starting time from strings like "09:00 - 11:00"
def extract_start_time(time_val):
    if pd.isna(time_val):
        return "00:00:00"
    time_str = str(time_val).strip()
    # Looks for a time pattern like 09:00, 9:00, 09.00
    match = re.search(r"\d{1,2}[:.]\d{2}", time_str)
    if match:
        # Standardize to HH:MM format
        time_clean = match.group().replace(".", ":")
        # Add a leading zero if it's single digit (e.g. 9:00 -> 09:00)
        if len(time_clean.split(":")[0]) == 1:
            time_clean = "0" + time_clean
        return time_clean + ":00"
    return "00:00:00"

if schedule_file and roster_file:
    if st.button("Genera Turni"):
        try:
            # --- STEP 0: Read Data ---
            df_schedule = pd.read_excel(schedule_file, header=5)
            df_schedule.dropna(subset=["Giorno", "Insegnamento"], inplace=True)

            # --- STEP 1: STRICT CHRONOLOGICAL SORTING (DATE + TIME) ---
            # Parse Date
            df_schedule["Giorno"] = pd.to_datetime(
                df_schedule["Giorno"], errors="coerce"
            )
            
            # Extract starting time
            df_schedule["start_time_str"] = df_schedule["Ora"].apply(
                extract_start_time
            )
            
            # Combine Date and Time into a single temporary sorting column
            df_schedule["datetime_sort"] = pd.to_datetime(
                df_schedule["Giorno"].dt.strftime("%Y-%m-%d")
                + " "
                + df_schedule["start_time_str"],
                errors="coerce",
            )

            # Sort everything chronologically and reset index
            df_schedule.sort_values(by="datetime_sort", inplace=True)
            df_schedule.reset_index(drop=True, inplace=True)


            # --- STEP 2: Student Preparation ---
            df_roster = pd.read_excel(roster_file)
            col_name = df_roster.columns[0]
            students = df_roster[col_name].dropna().astype(str).tolist()
            students = sorted([s.strip() for s in students])
            
            # Create the infinite loop for round-robin assignment
            student_pool = cycle(students)

            # --- STEP 3: Lesson Numbering (Now perfectly chronological) ---
            df_schedule["N. Lezione"] = (
                df_schedule.groupby("Insegnamento").cumcount() + 1
            )

            # --- STEP 4: Deadlines ---
            df_schedule["Scadenza consigliata"] = df_schedule[
                "Giorno"
            ] + pd.Timedelta(days=days_to_add)

            # --- STEP 5: Assign Roles ---
            sbobinatore_1 = []
            sbobinatore_2 = []
            controllore = []
            super_controllore = []

            for _ in range(len(df_schedule)):
                sbobinatore_1.append(next(student_pool))
                sbobinatore_2.append(next(student_pool))
                controllore.append(next(student_pool))
                super_controllore.append(next(student_pool))

            # --- STEP 6: Build Output ---
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
                    "Consegnato": "Falso",
                }
            )

            # --- STEP 7: Generate Excel File ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_output.to_excel(writer, index=False, sheet_name="Turni")
            processed_data = output.getvalue()

            st.success("✅ File elaborato, ordinato cronologicamente e generato con successo!")

            # Download button
            st.download_button(
                label="📥 Scarica il file Excel dei Turni",
                data=processed_data,
                file_name="Turni_Sbobinature_Ordinati.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
            )

        except Exception as e:
            st.error(f"Si è verificato un errore durante l'elaborazione: {e}")