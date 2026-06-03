import streamlit as st
import pandas as pd
import datetime
import io

# Set page configuration to wide layout similar to Streamlit defaults
st.set_page_config(
    page_title="Class Management Dashboard",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# 1. State Initialization with Mock Data
# -------------------------------------------------------------
if 'classes' not in st.session_state:
    st.session_state['classes'] = {
        'Class Section A': pd.DataFrame([
            {'ID Number': '101', 'First Name': 'Sophia', 'Last Name': 'Alvarez', 'Full Name': 'Sophia Alvarez'},
            {'ID Number': '102', 'First Name': 'Benjamin', 'Last Name': 'Chen', 'Full Name': 'Benjamin Chen'},
            {'ID Number': '103', 'First Name': 'Emma', 'Last Name': 'Dmitriev', 'Full Name': 'Emma Dmitriev'},
            {'ID Number': '104', 'First Name': 'Liam', 'Last Name': 'Johnson', 'Full Name': 'Liam Johnson'},
            {'ID Number': '105', 'First Name': 'Olivia', 'Last Name': 'Smith', 'Full Name': 'Olivia Smith'},
        ])
    }

if 'attendance' not in st.session_state:
    st.session_state['attendance'] = {
        'Class Section A': pd.DataFrame([
            {'ID Number': '101', 'Full Name': 'Sophia Alvarez', '2026-05-24': 'Attended'},
            {'ID Number': '102', 'Full Name': 'Benjamin Chen', '2026-05-24': 'Absent'},
            {'ID Number': '103', 'Full Name': 'Emma Dmitriev', '2026-05-24': 'Not Taken'},
            {'ID Number': '104', 'Full Name': 'Liam Johnson', '2026-05-24': 'Attended'},
            {'ID Number': '105', 'Full Name': 'Olivia Smith', '2026-05-24': 'Attended'},
        ])
    }

if 'grades' not in st.session_state:
    st.session_state['grades'] = {
        'Class Section A': pd.DataFrame([
            {'ID Number': '101', 'Full Name': 'Sophia Alvarez', 'Quiz 1 (w:10%)': 95.0, 'Midterm (w:40%)': 88.0},
            {'ID Number': '102', 'Full Name': 'Benjamin Chen', 'Quiz 1 (w:10%)': 82.0, 'Midterm (w:40%)': 79.0},
            {'ID Number': '103', 'Full Name': 'Emma Dmitriev', 'Quiz 1 (w:10%)': 90.0, 'Midterm (w:40%)': 92.0},
            {'ID Number': '104', 'Full Name': 'Liam Johnson', 'Quiz 1 (w:10%)': 70.0, 'Midterm (w:40%)': 85.0},
            {'ID Number': '105', 'Full Name': 'Olivia Smith', 'Quiz 1 (w:10%)': 88.0, 'Midterm (w:40%)': 91.0},
        ])
    }

if 'assessments' not in st.session_state:
    st.session_state['assessments'] = {
        'Class Section A': [
            {'name': 'Quiz 1', 'weight': 10.0, 'col_name': 'Quiz 1 (w:10%)'},
            {'name': 'Midterm', 'weight': 40.0, 'col_name': 'Midterm (w:40%)'}
        ]
    }

if 'pending_import_data' not in st.session_state:
    st.session_state['pending_import_data'] = None

# Helper to align dataframes on roster additions/deletions
def align_dataframes(class_name):
    students_df = st.session_state['classes'][class_name]
    
    # Align Attendance
    if class_name not in st.session_state['attendance']:
        st.session_state['attendance'][class_name] = pd.DataFrame(columns=['ID Number', 'Full Name'])
    
    att_df = st.session_state['attendance'][class_name]
    date_cols = [c for c in att_df.columns if c not in ['ID Number', 'Full Name']]
    new_att_data = []
    
    for _, s_row in students_df.iterrows():
        st_id = s_row['ID Number']
        name = s_row['Full Name']
        existing_row = att_df[att_df['ID Number'] == st_id]
        row_dict = {'ID Number': st_id, 'Full Name': name}
        for col in date_cols:
            if not existing_row.empty and col in existing_row.columns:
                row_dict[col] = existing_row.iloc[0][col]
            else:
                row_dict[col] = 'Not Taken'
        new_att_data.append(row_dict)
    
    if new_att_data:
        st.session_state['attendance'][class_name] = pd.DataFrame(new_att_data)
    else:
        st.session_state['attendance'][class_name] = pd.DataFrame(columns=['ID Number', 'Full Name'] + date_cols)

    # Align Grades
    if class_name not in st.session_state['grades']:
        st.session_state['grades'][class_name] = pd.DataFrame(columns=['ID Number', 'Full Name'])
        
    gr_df = st.session_state['grades'][class_name]
    assessment_cols = [c for c in gr_df.columns if c not in ['ID Number', 'Full Name', 'Final Grade']]
    new_gr_data = []
    for _, s_row in students_df.iterrows():
        st_id = s_row['ID Number']
        name = s_row['Full Name']
        existing_row = gr_df[gr_df['ID Number'] == st_id]
        row_dict = {'ID Number': st_id, 'Full Name': name}
        for col in assessment_cols:
            if not existing_row.empty and col in existing_row.columns:
                row_dict[col] = existing_row.iloc[0][col]
            else:
                row_dict[col] = 0.0
        new_gr_data.append(row_dict)
        
    if new_gr_data:
        st.session_state['grades'][class_name] = pd.DataFrame(new_gr_data)
    else:
        st.session_state['grades'][class_name] = pd.DataFrame(columns=['ID Number', 'Full Name'] + assessment_cols)


# Helper: recalculate final grade
def compute_final_grade(class_name):
    gr_df = st.session_state['grades'][class_name]
    assessments_list = st.session_state['assessments'].get(class_name, [])
    
    if not assessments_list:
        gr_df['Final Grade'] = 0.0
        return
        
    for idx, row in gr_df.iterrows():
        total_score = 0.0
        for assessment in assessments_list:
            col = assessment['col_name']
            val = pd.to_numeric(row[col], errors='coerce')
            if pd.isna(val):
                val = 0.0
            total_score += val
        
        gr_df.at[idx, 'Final Grade'] = round(total_score, 2)
        
    st.session_state['grades'][class_name] = gr_df


# Header and intro
st.title("🎓 Class Management Dashboard")
st.write("Maintain student lists, log daily attendance, and calculate final grades with custom export engines.")

# -------------------------------------------------------------
# 2. Sidebar Layout
# -------------------------------------------------------------
st.sidebar.header("🕹️ Class Configuration")

# Class Selection / Creation
existing_classes = list(st.session_state['classes'].keys())
new_class_option = "+ Create New Class"
class_selectbox_list = existing_classes + [new_class_option]

# Choose class section
selected_class = st.sidebar.selectbox("Choose Active Class Section", class_selectbox_list, index=0)

if selected_class == new_class_option:
    with st.sidebar.form("create_class_form", clear_on_submit=True):
        st.subheader("➕ New Class Section")
        new_class_name = st.text_input("Class Name (e.g., Mathematics 101)").strip()
        submit_new_class = st.form_submit_button("Create Class")
        
        if submit_new_class:
            if not new_class_name:
                st.sidebar.error("Class Name cannot be empty.")
            elif new_class_name in st.session_state['classes']:
                st.sidebar.warning(f"Class '{new_class_name}' already exists.")
            else:
                st.session_state['classes'][new_class_name] = pd.DataFrame(columns=['ID Number', 'First Name', 'Last Name', 'Full Name'])
                st.session_state['attendance'][new_class_name] = pd.DataFrame(columns=['ID Number', 'Full Name'])
                st.session_state['grades'][new_class_name] = pd.DataFrame(columns=['ID Number', 'Full Name'])
                st.session_state['assessments'][new_class_name] = []
                st.success(f"Created {new_class_name}!")
                st.rerun()
    st.stop()

# If classes is empty and they didn't create, show instructions
if not selected_class or selected_class not in st.session_state['classes']:
    st.info("Please create or select a class in the sidebar to begin.")
    st.stop()

# -------------------------------------------------------------
# Sidebar: Delete Class Section Module
# -------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🗑️ Delete Current Section")
confirm_deletion_key = f"confirm_del_{selected_class}"
if confirm_deletion_key not in st.session_state:
    st.session_state[confirm_deletion_key] = False

if len(st.session_state['classes']) <= 1:
    st.sidebar.caption("Cannot delete active section. At least one class database must exist.")
else:
    if not st.session_state[confirm_deletion_key]:
        if st.sidebar.button(f"Delete Section '{selected_class}'", type="secondary"):
            st.session_state[confirm_deletion_key] = True
            st.rerun()
    else:
        st.sidebar.warning(f"Delete '{selected_class}'? This will completely remove all students, grades, and logs!")
        col_del1, col_del2 = st.sidebar.columns(2)
        with col_del1:
            if st.button("Yes, Clear Data", type="primary"):
                # Remove active section databases
                del st.session_state['classes'][selected_class]
                if selected_class in st.session_state['attendance']:
                    del st.session_state['attendance'][selected_class]
                if selected_class in st.session_state['grades']:
                    del st.session_state['grades'][selected_class]
                if selected_class in st.session_state['assessments']:
                    del st.session_state['assessments'][selected_class]
                
                # Settle active selected section
                remaining_classes = list(st.session_state['classes'].keys())
                st.sidebar.success(f"Deleted {selected_class}.")
                del st.session_state[confirm_deletion_key]
                st.rerun()
        with col_del2:
            if st.button("Cancel"):
                st.session_state[confirm_deletion_key] = False
                st.rerun()

st.sidebar.markdown("---")

# -------------------------------------------------------------
# Sidebar: Roster Spreadsheet Import
# -------------------------------------------------------------
st.sidebar.subheader("📥 Import Student Roster")
uploaded_file = st.sidebar.file_uploader(
    "Import CSV/Excel with 'First Name', 'Last Name', and 'ID Number'",
    type=['csv', 'xlsx']
)

if uploaded_file is not None:
    try:
        # Load tables
        if uploaded_file.name.endswith('.csv'):
            imported_df = pd.read_csv(uploaded_file)
        else:
            imported_df = pd.read_excel(uploaded_file)
            
        # Standardize column naming
        imported_df.columns = [str(c).strip().title() for c in imported_df.columns]
        
        # Mapping matching candidates
        col_mappings = {}
        for candidate in ['Id Number', 'Id', 'Student Id', 'Idnumber']:
            col_match = [c for c in imported_df.columns if c.lower() == candidate.lower()]
            if col_match:
                col_mappings['ID Number'] = col_match[0]
                break
                
        for candidate in ['First Name', 'Firstname', 'First']:
            col_match = [c for c in imported_df.columns if c.lower() == candidate.lower()]
            if col_match:
                col_mappings['First Name'] = col_match[0]
                break
                
        for candidate in ['Last Name', 'Lastname', 'Last']:
            col_match = [c for c in imported_df.columns if c.lower() == candidate.lower()]
            if col_match:
                col_mappings['Last Name'] = col_match[0]
                break

        # Validation checks
        if 'First Name' not in col_mappings or 'Last Name' not in col_mappings or 'ID Number' not in col_mappings:
            st.sidebar.error("Uploaded file must contain 'First Name', 'Last Name', and 'ID Number' columns.")
        else:
            # Build student dictionaries
            temp_students = []
            ids_raw = imported_df[col_mappings['ID Number']].astype(str).str.strip().tolist()
            firsts = imported_df[col_mappings['First Name']].astype(str).str.strip().tolist()
            lasts = imported_df[col_mappings['Last Name']].astype(str).str.strip().tolist()
            
            for i in range(len(ids_raw)):
                st_id = ids_raw[i]
                fn = firsts[i]
                ln = lasts[i]
                # exclude nan values
                if st_id and st_id != 'nan' and fn and fn != 'nan' and ln and ln != 'nan':
                    temp_students.append({
                        'ID Number': st_id,
                        'First Name': fn,
                        'Last Name': ln,
                        'Full Name': f"{fn} {ln}"
                    })
            
            # Exclude duplicate IDs from import
            unique_students = []
            seen_ids = set()
            for s in temp_students:
                if s['ID Number'] not in seen_ids:
                    unique_students.append(s)
                    seen_ids.add(s['ID Number'])
                    
            # Filter student IDs containing 6 or less digits
            short_ids = [s for s in unique_students if len(s['ID Number']) <= 6]
            normal_ids = [s for s in unique_students if len(s['ID Number']) > 6]
            
            if short_ids:
                # Intercept file loader and set pending approval payload state
                st.session_state['pending_import_data'] = {
                    'class': selected_class,
                    'normal': normal_ids,
                    'short': short_ids
                }
                st.sidebar.info("⚠️ Verification check is required on Main Console.")
                st.rerun()
            else:
                # Direct save
                cleaned_roster = pd.DataFrame(normal_ids).sort_values(by='First Name').reset_index(drop=True)
                st.session_state['classes'][selected_class] = cleaned_roster
                align_dataframes(selected_class)
                compute_final_grade(selected_class)
                st.sidebar.success(f"Imported {len(cleaned_roster)} students successfully!")
                st.rerun()
                
    except Exception as e:
        st.sidebar.error(f"Error parsing file: {e}")

st.sidebar.markdown("---")

# -------------------------------------------------------------
# Sidebar: Manual Roster Insert/Removal Forms
# -------------------------------------------------------------
st.sidebar.subheader("👤 Add New Student")
with st.sidebar.form("add_student_form", clear_on_submit=True):
    add_id = st.text_input("ID Number").strip()
    add_first = st.text_input("First Name").strip()
    add_last = st.text_input("Last Name").strip()
    submit_add = st.form_submit_button("Add Student to Roster")
    
    if submit_add:
        if not add_id or not add_first or not add_last:
            st.sidebar.error("All fields are required.")
        else:
            current_roster = st.session_state['classes'][selected_class]
            if add_id in current_roster['ID Number'].values:
                st.sidebar.error(f"ID Number '{add_id}' already registered.")
            else:
                new_student = {
                    'ID Number': add_id,
                    'First Name': add_first,
                    'Last Name': add_last,
                    'Full Name': f"{add_first} {add_last}"
                }
                updated_roster = pd.concat([current_roster, pd.DataFrame([new_student])], ignore_index=True)
                updated_roster = updated_roster.sort_values(by='First Name').reset_index(drop=True)
                
                st.session_state['classes'][selected_class] = updated_roster
                align_dataframes(selected_class)
                compute_final_grade(selected_class)
                st.sidebar.success(f"Added {add_first} to roster!")
                st.rerun()

st.sidebar.markdown("---")

# Remove Student form
st.sidebar.subheader("🗑️ Delete Student")
current_roster = st.session_state['classes'][selected_class]

if not current_roster.empty:
    student_removal_options = [f"{row['ID Number']} - {row['Full Name']}" for _, row in current_roster.iterrows()]
    selected_removal_str = st.sidebar.selectbox("Select Student to Remove", student_removal_options)
    removal_mode = st.sidebar.radio("Scope of Deletion", ["Remove from current class", "Remove from all classes"])
    
    if st.sidebar.button("Execute Deletion", type="primary"):
        target_id = selected_removal_str.split(" - ")[0]
        st_name = selected_removal_str.split(" - ")[1]
        
        if removal_mode == "Remove from current class":
            st.session_state['classes'][selected_class] = current_roster[current_roster['ID Number'] != target_id]
            align_dataframes(selected_class)
            compute_final_grade(selected_class)
            st.sidebar.success(f"Removed {st_name} from active roster.")
        else:
            for c_name in list(st.session_state['classes'].keys()):
                st.session_state['classes'][c_name] = st.session_state['classes'][c_name][st.session_state['classes'][c_name]['ID Number'] != target_id]
                align_dataframes(c_name)
                compute_final_grade(c_name)
            st.sidebar.success(f"Removed {st_name} from all class files.")
        st.rerun()
else:
    st.sidebar.caption("No registered students found to delete.")

# -------------------------------------------------------------
# 3. Main Console Display & Interactive ID Approval
# -------------------------------------------------------------
if st.session_state['pending_import_data'] is not None:
    pending = st.session_state['pending_import_data']
    st.markdown("---")
    st.warning("⚠️ **Roster Verification Required: Short ID Numbers Detected**")
    st.write(
        f"We located students in your imported spreadsheet who have ID numbers containing **6 or less digits**. "
        f"Choose which of these students you would like to keep in **{pending['class']}**. "
        "Deselected students will be discarded from importation:"
    )
    
    # Render interactive decider grid
    short_df = pd.DataFrame(pending['short'])
    short_df.insert(0, 'Keep Student', True) # Default keep selection
    
    edited_short_df = st.data_editor(
        short_df,
        column_config={
            "Keep Student": st.column_config.CheckboxColumn("Keep Student?", default=True),
            "ID Number": st.column_config.TextColumn("ID Number", disabled=True),
            "First Name": st.column_config.TextColumn("First Name", disabled=True),
            "Last Name": st.column_config.TextColumn("Last Name", disabled=True),
            "Full Name": st.column_config.TextColumn("Full Name", disabled=True),
        },
        hide_index=True,
        use_container_width=True,
        key="short_id_roster_decider"
    )
    
    col_app1, col_app2, col_app3 = st.columns([1, 1, 3])
    with col_app1:
        if st.button("Keep Approved & Import", type="primary"):
            approved_short_list = []
            for idx, r in edited_short_df.iterrows():
                if r['Keep Student']:
                    approved_short_list.append({
                        'ID Number': str(r['ID Number']),
                        'First Name': str(r['First Name']),
                        'Last Name': str(r['Last Name']),
                        'Full Name': str(r['Full Name'])
                    })
            
            # Combine normal and approved IDs
            final_roster_list = pending['normal'] + approved_short_list
            final_df = pd.DataFrame(final_roster_list)
            if not final_df.empty:
                final_df = final_df.sort_values(by='First Name').reset_index(drop=True)
            else:
                final_df = pd.DataFrame(columns=['ID Number', 'First Name', 'Last Name', 'Full Name'])
                
            st.session_state['classes'][pending['class']] = final_df
            align_dataframes(pending['class'])
            compute_final_grade(pending['class'])
            st.session_state['pending_import_data'] = None
            st.success("Successfully imported student roster database sections!")
            st.rerun()
            
    with col_app2:
        if st.button("Discard All Short IDs"):
            final_df = pd.DataFrame(pending['normal'])
            if not final_df.empty:
                final_df = final_df.sort_values(by='First Name').reset_index(drop=True)
            else:
                final_df = pd.DataFrame(columns=['ID Number', 'First Name', 'Last Name', 'Full Name'])
                
            st.session_state['classes'][pending['class']] = final_df
            align_dataframes(pending['class'])
            compute_final_grade(pending['class'])
            st.session_state['pending_import_data'] = None
            st.success("Roster successfully loaded (all ≤ 6 digit IDs was skipped)!")
            st.rerun()
            
    with col_app3:
        if st.button("Abort Importation"):
            st.session_state['pending_import_data'] = None
            st.info("Spreadsheet import cancelled.")
            st.rerun()
    st.markdown("---")

# Main Interface Tabs Setup
tab_std, tab_att, tab_grd = st.tabs(["👥 Students List", "📅 Attendance Registry", "📊 Academic Grades"])
current_students = st.session_state['classes'][selected_class]

# --- TAB 1: STUDENTS ---
with tab_std:
    st.subheader(f"Class Roster Database: {selected_class}")
    if current_students.empty:
        st.info("The roster list is currently empty. Use the sidebar roster tools to import or manually log student information.")
    else:
        st.markdown(f"**Total Registered Students:** `{len(current_students)}`")
        
        display_roster = current_students[['ID Number', 'Full Name']].copy()
        st.dataframe(
            display_roster,
            column_config={
                "ID Number": st.column_config.TextColumn("ID Number", width="medium"),
                "Full Name": st.column_config.TextColumn("Full Name", width="large")
            },
            hide_index=True,
            use_container_width=True
        )

# --- TAB 2: ATTENDANCE ---
with tab_att:
    st.subheader("Daily Attendance Registry")
    
    # Take Attendance Trigger
    lead_col1, lead_col2 = st.columns([2, 5])
    with lead_col1:
        att_date = st.date_input("Attendance Log Date", datetime.date.today())
    
    with lead_col2:
        st.write("") # spacer
        st.write("")
        if st.button("Take Attendance (Add Column)", type="secondary"):
            date_col_str = att_date.strftime("%Y-%m-%d")
            att_df = st.session_state['attendance'][selected_class]
            
            if date_col_str in att_df.columns:
                st.warning(f"Attendance record for {date_col_str} already exists. You can modify states below.")
            else:
                att_df[date_col_str] = 'Not Taken'
                st.session_state['attendance'][selected_class] = att_df
                st.success(f"Added column: {date_col_str}")
                st.rerun()
                
    st.markdown("---")
    
    att_df = st.session_state['attendance'][selected_class]
    date_cols = [c for c in att_df.columns if c not in ['ID Number', 'Full Name']]
    
    if current_students.empty:
        st.info("No student records found. Setup a roster first.")
    elif not date_cols:
        st.info("Select a date above and click 'Take Attendance' to initiate your first attendance column.")
    else:
        st.markdown("**Instructions:** Configure attendance records below. You can change select status values or utilize the edit grid directly.")
        
        status_options = ['Not Taken', 'Attended', 'Absent']
        
        edited_att_df = st.data_editor(
            att_df,
            column_config={
                "ID Number": st.column_config.TextColumn("ID Number", disabled=True),
                "Full Name": st.column_config.TextColumn("Full Name", disabled=True),
                **{col: st.column_config.SelectboxColumn(col, options=status_options) for col in date_cols}
            },
            hide_index=True,
            use_container_width=True,
            key="attendance_ledger_editor"
        )
        
        if not edited_att_df.equals(att_df):
            st.session_state['attendance'][selected_class] = edited_att_df
            st.success("Attendance ledger updated successfully.")
            st.rerun()

# --- TAB 3: GRADES ---
with tab_grd:
    st.subheader("Academic Ledger & Grading Sheet")
    
    # Form to create a new assessment
    st.markdown("### 📝 Define Assessment Metrics")
    col_a, col_b, col_c = st.columns([3, 2, 2])
    with col_a:
        assess_name = st.text_input("Exam / Assessment Name", placeholder="e.g., Final Exam").strip()
    with col_b:
        assess_weight = st.number_input("Weight (Value coefficient / %)", min_value=0.0, max_value=100.0, value=20.0, step=5.0)
    with col_c:
        st.write("") # spacer
        st.write("")
        create_assessment_btn = st.button("Register Assessment", use_container_width=True)
        
    if create_assessment_btn:
        if not assess_name:
            st.error("Please enter a valid assessment title.")
        else:
            col_id = f"{assess_name} (w:{assess_weight}%)"
            gr_df = st.session_state['grades'][selected_class]
            
            existing_cols = st.session_state['assessments'].get(selected_class, [])
            if any(x['name'].lower() == assess_name.lower() for x in existing_cols):
                st.error(f"An assessment named '{assess_name}' is already defined.")
            else:
                if selected_class not in st.session_state['assessments']:
                    st.session_state['assessments'][selected_class] = []
                st.session_state['assessments'][selected_class].append({
                    'name': assess_name,
                    'weight': assess_weight,
                    'col_name': col_id
                })
                
                gr_df[col_id] = 0.0
                st.session_state['grades'][selected_class] = gr_df
                compute_final_grade(selected_class)
                st.success(f"Created '{assess_name}' with weight {assess_weight}%!")
                st.rerun()
                
    st.markdown("---")
    
    # Active Grading Ledger
    st.markdown("### 📊 Active Grading Ledger")
    gr_df = st.session_state['grades'][selected_class]
    active_assessments = st.session_state['assessments'].get(selected_class, [])
    
    if current_students.empty:
        st.info("Roster is currently empty. Define students in the sidebar first.")
    elif not active_assessments:
        st.info("No assessment metrics registered. Create a new assessment using the forms above to start entering marks.")
    else:
        st.markdown("*Adjust numerical values (0-100) under each registered Assessment. The **Final Grade** represents the sum of all assessments computed automatically.*")
        
        grad_cols = [a['col_name'] for a in active_assessments]
        column_conf = {
            "ID Number": st.column_config.TextColumn("ID Number", disabled=True),
            "Full Name": st.column_config.TextColumn("Full Name", disabled=True),
            "Final Grade": st.column_config.NumberColumn(
                "📊 Final Grade", 
                disabled=True, 
                format="%.2f",
                help="Computed dynamically based on registered items"
            )
        }
        
        for col_name in grad_cols:
            column_conf[col_name] = st.column_config.NumberColumn(
                col_name,
                min_value=0.0,
                max_value=120.0,
                step=0.5,
                format="%.1f"
            )
            
        compute_final_grade(selected_class)
        gr_df = st.session_state['grades'][selected_class]
        
        edited_gr_df = st.data_editor(
            gr_df,
            column_config=column_conf,
            hide_index=True,
            use_container_width=True,
            key="grades_ledger_editor"
        )
        
        if not edited_gr_df.equals(gr_df):
            st.session_state['grades'][selected_class] = edited_gr_df
            compute_final_grade(selected_class)
            st.success("Marks saved and final scores re-evaluated successfully.")
            st.rerun()

# -------------------------------------------------------------
# 4. Global Excel Export Capabilities (In Sidebar Base)
# -------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("📤 Export & Download Ledgers")

# Helper function to generate cleanly formatted sheets with borders & alignment
def create_styled_excel(df, export_type='grades'):
    output_bytes = io.BytesIO()
    writer = pd.ExcelWriter(output_bytes, engine='xlsxwriter')
    
    df.to_excel(writer, sheet_name=f'Exported_{export_type}', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets[f'Exported_{export_type}']
    
    # Premium borders styling
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'middle',
        'fg_color': '#1F4E79',
        'font_color': '#FFFFFF',
        'border': 1,
        'border_color': '#112233'
    })
    
    cell_border_format = workbook.add_format({
        'border': 1,
        'border_color': '#D0D3D4',
        'align': 'left',
        'valign': 'middle'
    })
    
    attended_format = workbook.add_format({
        'bg_color': '#DECF9', # custom soft green
        'fg_color': '#000000',
        'bg_color': '#D0E1D4',
        'font_color': '#2C5E3B',
        'border': 1,
        'border_color': '#ACD0B6',
        'align': 'center'
    })
    
    absent_format = workbook.add_format({
        'bg_color': '#FADBD8',
        'font_color': '#78281F',
        'border': 1,
        'border_color': '#EBB6B3',
        'align': 'center'
    })
    
    not_taken_format = workbook.add_format({
        'bg_color': '#E5E7E9',
        'font_color': '#5D6D7E',
        'border': 1,
        'border_color': '#D3D5D7',
        'align': 'center'
    })
    
    # Cover the layout head row
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell_val = df.iloc[row_idx, col_idx]
            
            if export_type == 'attendance':
                val_str = str(cell_val)
                if val_str == 'Attended':
                    worksheet.write(row_idx + 1, col_idx, val_str, attended_format)
                elif val_str == 'Absent':
                    worksheet.write(row_idx + 1, col_idx, val_str, absent_format)
                elif val_str == 'Not Taken':
                    worksheet.write(row_idx + 1, col_idx, val_str, not_taken_format)
                else:
                    worksheet.write(row_idx + 1, col_idx, val_str, cell_border_format)
            elif export_type == 'grades':
                final_grade_col_idx = len(df.columns) - 1
                if col_idx == final_grade_col_idx:
                    final_grade_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#FCF3CF', # soft yellow
                        'font_color': '#7E5109',
                        'border': 1,
                        'border_color': '#BDC3C7',
                        'align': 'right'
                    })
                    try:
                        worksheet.write_number(row_idx + 1, col_idx, float(cell_val), final_grade_format)
                    except:
                        worksheet.write(row_idx + 1, col_idx, str(cell_val), final_grade_format)
                else:
                    try:
                        worksheet.write_number(row_idx + 1, col_idx, float(cell_val), cell_border_format)
                    except:
                        worksheet.write(row_idx + 1, col_idx, str(cell_val), cell_border_format)
                        
    # Autofit columns
    for col_num in range(len(df.columns)):
        col_name = df.columns[col_num]
        max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 3
        worksheet.set_column(col_num, col_num, min(max(max_len, 10), 30))
        
    writer.close()
    return output_bytes.getvalue()


# Helper: Dynamic Single calculator for merged exports
def get_individual_final_score(class_name, student_id):
    gr_df = st.session_state['grades'].get(class_name, pd.DataFrame())
    assessments_list = st.session_state['assessments'].get(class_name, [])
    if gr_df.empty or not assessments_list:
        return 0.0
    student_row = gr_df[gr_df['ID Number'] == student_id]
    if student_row.empty:
        return 0.0
    
    total_score = 0.0
    for assessment in assessments_list:
        col = assessment['col_name']
        val = pd.to_numeric(student_row[col].iloc[0], errors='coerce')
        if pd.isna(val):
            val = 0.0
        total_score += val
        
    return round(total_score, 2)


# MERGED GENERATION ENGINE: Combines sections with clean separation
def create_merged_excel_report(export_type='grades'):
    output_bytes = io.BytesIO()
    writer = pd.ExcelWriter(output_bytes, engine='xlsxwriter')
    workbook = writer.book
    
    sheet_name = f"Merged_{export_type.capitalize()}"
    worksheet = workbook.add_sheet(sheet_name)
    
    # Styled formats
    class_title_format = workbook.add_format({
        'bold': True,
        'font_size': 13,
        'bg_color': '#D4E6F1',
        'font_color': '#1B4F72',
        'border': 1,
        'border_color': '#112233',
        'align': 'left',
        'valign': 'middle'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 11,
        'text_wrap': True,
        'bg_color': '#1F4E79',
        'font_color': '#FFFFFF',
        'border': 1,
        'border_color': '#1F4E79',
        'align': 'left',
        'valign': 'middle'
    })
    
    cell_border_format = workbook.add_format({
        'border': 1,
        'border_color': '#D0D3D4',
        'align': 'left',
        'valign': 'middle'
    })
    
    attended_format = workbook.add_format({
        'bg_color': '#D0E1D4',
        'font_color': '#2C5E3B',
        'border': 1,
        'border_color': '#BDC3C7',
        'align': 'center'
    })
    
    absent_format = workbook.add_format({
        'bg_color': '#FADBD8',
        'font_color': '#78281F',
        'border': 1,
        'border_color': '#BDC3C7',
        'align': 'center'
    })
    
    not_taken_format = workbook.add_format({
        'bg_color': '#E5E7E9',
        'font_color': '#5D6D7E',
        'border': 1,
        'border_color': '#BDC3C7',
        'align': 'center'
    })
    
    final_grade_format = workbook.add_format({
        'bold': True,
        'bg_color': '#FCF3CF',
        'font_color': '#7E5109',
        'border': 1,
        'border_color': '#BDC3C7',
        'align': 'right'
    })
    
    current_row = 0
    col_widths = {}
    
    for class_name in list(st.session_state['classes'].keys()):
        students_df = st.session_state['classes'][class_name]
        if students_df.empty:
            continue
            
        # 1. Write Merged / Separated Class title row
        worksheet.write(current_row, 0, f"🏫 {class_name}", class_title_format)
        worksheet.set_row(current_row, 26)
        current_row += 1
        
        # 2. Write headers row
        if export_type == 'attendance':
            att_df = st.session_state['attendance'].get(class_name, pd.DataFrame(columns=['ID Number', 'Full Name']))
            date_cols = [c for c in att_df.columns if c not in ['ID Number', 'Full Name']]
            headers = ['ID Number', 'Full Name'] + date_cols
            
            for col_num, h_name in enumerate(headers):
                worksheet.write(current_row, col_num, h_name, header_format)
                col_widths[col_num] = max(col_widths.get(col_num, 0), len(str(h_name)) + 3)
            current_row += 1
            
            # Write student details & status keys
            for idx, r in att_df.iterrows():
                for col_num, h_name in enumerate(headers):
                    val = str(r.get(h_name, ''))
                    if val == 'Attended':
                        worksheet.write(current_row, col_num, val, attended_format)
                    elif val == 'Absent':
                        worksheet.write(current_row, col_num, val, absent_format)
                    elif val == 'Not Taken':
                        worksheet.write(current_row, col_num, val, not_taken_format)
                    else:
                        worksheet.write(current_row, col_num, val, cell_border_format)
                    col_widths[col_num] = max(col_widths.get(col_num, 0), len(str(val)) + 3)
                current_row += 1
                
        else: # Grades
            gr_df = st.session_state['grades'].get(class_name, pd.DataFrame(columns=['ID Number', 'Full Name']))
            assessments_list = st.session_state['assessments'].get(class_name, [])
            assess_cols = [a['col_name'] for a in assessments_list]
            headers = ['ID Number', 'Full Name'] + assess_cols + ['Final Grade']
            
            for col_num, h_name in enumerate(headers):
                worksheet.write(current_row, col_num, h_name, header_format)
                col_widths[col_num] = max(col_widths.get(col_num, 0), len(str(h_name)) + 3)
            current_row += 1
            
            # Write Student score details
            for idx, r in gr_df.iterrows():
                for col_num, h_name in enumerate(headers):
                    val = r.get(h_name, '')
                    
                    if h_name == 'Final Grade':
                        final_v = get_individual_final_score(class_name, str(r['ID Number']))
                        worksheet.write_number(current_row, col_num, float(final_v), final_grade_format)
                        col_widths[col_num] = max(col_widths.get(col_num, 0), len(f"{final_v}") + 3)
                    elif h_name in assess_cols:
                        try:
                            num_val = float(val)
                            worksheet.write_number(current_row, col_num, num_val, cell_border_format)
                        except:
                            worksheet.write(current_row, col_num, str(val), cell_border_format)
                        col_widths[col_num] = max(col_widths.get(col_num, 0), len(str(val)) + 3)
                    else:
                        worksheet.write(current_row, col_num, str(val), cell_border_format)
                        col_widths[col_num] = max(col_widths.get(col_num, 0), len(str(val)) + 3)
                current_row += 1
                
        # 3. Add 2 lines of empty cells
        current_row += 2
        
    # Auto-adjust Dynamic sized width lengths
    for col_idx, width in col_widths.items():
        worksheet.set_column(col_idx, col_idx, min(max(width, 10), 30))
        
    writer.close()
    return output_bytes.getvalue()


# Download Action buttons
if not current_students.empty:
    st.sidebar.markdown("**Current Section Reports:**")
    grades_excel = create_styled_excel(st.session_state['grades'][selected_class], export_type='grades')
    st.sidebar.download_button(
        label="📥 Descargar Grades (Current Excel)",
        data=grades_excel,
        file_name=f"Grades_Report_{selected_class.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    att_excel = create_styled_excel(st.session_state['attendance'][selected_class], export_type='attendance')
    st.sidebar.download_button(
        label="📥 Descargar Attendance (Current Excel)",
        data=att_excel,
        file_name=f"Attendance_Log_{selected_class.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Merged Downloader
if len(st.session_state['classes']) > 0:
    st.sidebar.markdown("---")
    st.sidebar.markdown("**🌐 Merged Global Reports (All Sections):**")
    
    merged_grades_excel = create_merged_excel_report(export_type='grades')
    st.sidebar.download_button(
        label="📥 Download Merged Grades (All Classes)",
        data=merged_grades_excel,
        file_name=f"Merged_Grades_Report_All_Sections_{datetime.date.today().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    merged_att_excel = create_merged_excel_report(export_type='attendance')
    st.sidebar.download_button(
        label="📥 Download Merged Attendance (All Classes)",
        data=merged_att_excel,
        file_name=f"Merged_Attendance_Log_All_Sections_{datetime.date.today().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.sidebar.caption("State persistent in st.session_state equivalents.")
