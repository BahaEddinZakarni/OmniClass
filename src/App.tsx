import React, { useState, useEffect, useRef } from 'react';
import { 
  GraduationCap, 
  Users, 
  CalendarCheck, 
  Trash2, 
  Plus, 
  Download, 
  Upload, 
  AlertCircle, 
  Check, 
  Calendar, 
  UserPlus,
  FileSpreadsheet,
  Code,
  Copy,
  ChevronRight,
  Info,
  X,
  Menu,
  Search,
  LogOut,
  LogIn,
  Cloud,
  CloudOff,
  RefreshCw,
  Lock
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { Student, Assessment, ClassSection } from './types';

// Import Firebase config & handlers
import { auth, db, googleProvider, handleFirestoreError, OperationType } from './firebase';
import { signInWithPopup, signOut, onAuthStateChanged, User, signInWithEmailAndPassword, createUserWithEmailAndPassword } from 'firebase/auth';
import { doc, setDoc, onSnapshot } from 'firebase/firestore';

// Matching app.py representation precisely for visual source codes exporter panel
const PYTHON_STREAMLIT_CODE = `import streamlit as st
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

# Deletion module in sidebar, etc. (See exported code inside the tabs)
`;

// Initial classes set
const INITIAL_CLASSES: ClassSection[] = [
  {
    name: 'Class Section A',
    students: [
      { id: '101', firstName: 'Sophia', lastName: 'Alvarez', fullName: 'Sophia Alvarez' },
      { id: '102', firstName: 'Benjamin', lastName: 'Chen', fullName: 'Benjamin Chen' },
      { id: '103', firstName: 'Emma', lastName: 'Dmitriev', fullName: 'Emma Dmitriev' },
      { id: '104', firstName: 'Liam', lastName: 'Johnson', fullName: 'Liam Johnson' },
      { id: '105', firstName: 'Olivia', lastName: 'Smith', fullName: 'Olivia Smith' },
    ],
    attendance: {
      '101': { '2026-05-24': 'Attended' },
      '102': { '2026-05-24': 'Absent' },
      '103': { '2026-05-24': 'Not Taken' },
      '104': { '2026-05-24': 'Attended' },
      '105': { '2026-05-24': 'Attended' },
    },
    assessments: [
      { name: 'Quiz 1', weight: 10 },
      { name: 'Midterm', weight: 40 }
    ],
    grades: {
      '101': { 'Quiz 1': 95, 'Midterm': 88 },
      '102': { 'Quiz 1': 82, 'Midterm': 79 },
      '103': { 'Quiz 1': 90, 'Midterm': 92 },
      '104': { 'Quiz 1': 70, 'Midterm': 85 },
      '105': { 'Quiz 1': 88, 'Midterm': 91 },
    }
  },
  {
    name: 'Class Section B',
    students: [
      { id: '201', firstName: 'Amara', lastName: 'Patel', fullName: 'Amara Patel' },
      { id: '202', firstName: 'Gabriel', lastName: 'Muller', fullName: 'Gabriel Muller' },
      { id: '203', firstName: 'Zoe', lastName: 'Dupont', fullName: 'Zoe Dupont' }
    ],
    attendance: {},
    assessments: [],
    grades: {}
  }
];

const HIGHLIGHT_COLORS = [
  { id: 'emerald', name: 'Green', bgClass: 'bg-emerald-50', hoverClass: 'hover:bg-emerald-100/60', borderClass: 'border-l-emerald-500', dotClass: 'bg-emerald-500' },
  { id: 'sky', name: 'Blue', bgClass: 'bg-sky-50', hoverClass: 'hover:bg-sky-100/60', borderClass: 'border-l-sky-500', dotClass: 'bg-sky-500' }
];

const getStudentHighlightClasses = (colorId?: string) => {
  if (!colorId) return { bg: '', hover: 'hover:bg-slate-50/50', border: '', text: 'text-slate-800' };
  const found = HIGHLIGHT_COLORS.find(c => c.id === colorId);
  if (!found) return { bg: '', hover: 'hover:bg-slate-50/50', border: '', text: 'text-slate-800' };
  return {
    bg: found.bgClass,
    hover: found.hoverClass,
    border: `border-l-4 ${found.borderClass}`,
    text: 'text-slate-800'
  };
};

export default function App() {
  // Session States modeled on st.session_state
  const [classes, setClasses] = useState<ClassSection[]>(() => {
    const saved = localStorage.getItem('streamlit_class_sections');
    return saved ? JSON.parse(saved) : INITIAL_CLASSES;
  });
  
  const setStudentHighlightColor = (studentId: string, className: string, colorId: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === className) {
        return {
          ...cl,
          students: cl.students.map(st => {
            if (st.id === studentId) {
              return { ...st, color: colorId === 'none' ? undefined : colorId };
            }
            return st;
          })
        };
      }
      return cl;
    }));
  };
  
  const [activeClassName, setActiveClassName] = useState<string>('Class Section A');
  const [activeTab, setActiveTab] = useState<'students' | 'attendance' | 'grades' | 'code'>('students');

  // Sidebar Controls
  const [newClassNameInput, setNewClassNameInput] = useState('');
  const [isCreatingClass, setIsCreatingClass] = useState(false);
  const [confirmDeleteClass, setConfirmDeleteClass] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  // Manual Add Student form
  const [studentIdInput, setStudentIdInput] = useState('');
  const [studentFirstInput, setStudentFirstInput] = useState('');
  const [studentLastInput, setStudentLastInput] = useState('');
  const [studentAddError, setStudentAddError] = useState<string | null>(null);
  const [studentAddSuccess, setStudentAddSuccess] = useState<string | null>(null);

  // Deletion state
  const [studentToDeleteId, setStudentToDeleteId] = useState('');
  const [deletionScope, setDeletionScope] = useState<'current' | 'all'>('current');
  const [deletionSuccess, setDeletionSuccess] = useState<string | null>(null);

  // Attendance state
  const [attendanceDate, setAttendanceDate] = useState<string>('2026-05-25');
  const [takeAttendanceMsg, setTakeAttendanceMsgMessage] = useState<{type: 'success'|'warning', text: string} | null>(null);

  // Grade Assessment Input state
  const [assessNameInput, setAssessNameInput] = useState('');
  const [assessWeightInput, setAssessWeightInput] = useState<number>(20);
  const [assessFeedback, setAssessFeedback] = useState<{type: 'error' | 'success', text: string} | null>(null);

  // Source code copy indicator
  const [copiedCode, setCopiedCode] = useState(false);

  // Global Student Search state
  const [globalSearchQuery, setGlobalSearchQuery] = useState('');

  // File Upload states
  const [fileError, setFileError] = useState<string | null>(null);
  const [fileSuccess, setFileSuccess] = useState<string | null>(null);

  // Staging state for list imports containing short ID values
  const [pendingImport, setPendingImport] = useState<{
    className: string;
    normal: Student[];
    short: Student[];
  } | null>(null);
  const [selectedShortIds, setSelectedShortIds] = useState<string[]>([]);

  // Firebase Authentication & Realtime Cloud Sync engine
  const [user, setUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncErrorMessage, setSyncErrorMessage] = useState<string | null>(null);
  
  const lastFirestoreDataRef = useRef<string>('');
  const isInitialDocFetched = useRef<boolean>(false);

  // 1. Listen for Authentication Changes
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (firebaseUser) => {
      setUser(firebaseUser);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // 2. Real-time Firestore inbound state updates (onSnapshot)
  useEffect(() => {
    if (!user) {
      lastFirestoreDataRef.current = '';
      isInitialDocFetched.current = false;
      return;
    }

    isInitialDocFetched.current = false;
    setIsSyncing(true);
    setSyncErrorMessage(null);

    const docRef = doc(db, 'users', user.uid);
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      setIsSyncing(false);
      if (docSnap.exists()) {
        const data = docSnap.data();
        if (data && Array.isArray(data.classes)) {
          const firestoreClassesStr = JSON.stringify(data.classes);
          lastFirestoreDataRef.current = firestoreClassesStr;
          // Avoid setting state and causing infinite loops if data didn't actually change
          if (firestoreClassesStr !== JSON.stringify(classes)) {
            setClasses(data.classes);
          }
        }
      } else {
        // First-time user: persist current client-side state into Firestore automatically
        setDoc(docRef, { classes }, { merge: true })
          .then(() => {
            lastFirestoreDataRef.current = JSON.stringify(classes);
          })
          .catch((err) => {
            console.error("Failed to sync initial roster to Firestore:", err);
            setSyncErrorMessage("Failed to sync initial roster to Firestore.");
          });
      }
      isInitialDocFetched.current = true;
    }, (error) => {
      setIsSyncing(false);
      console.error("Firestore synchronizer error:", error);
      setSyncErrorMessage("Unauthorized connection or offline database error.");
    });

    return () => unsubscribe();
  }, [user]);

  // 3. Keep local fallback in sync & upload local modifications to cloud doc
  useEffect(() => {
    // Sync to localStorage always as our bulletproof local fallback buffer
    localStorage.setItem('streamlit_class_sections', JSON.stringify(classes));

    if (!user) return;
    if (!isInitialDocFetched.current) return; // Do not upload before receiving current firestore data!

    const currentClassesStr = JSON.stringify(classes);
    // If the state was updated locally and differs from last Firestore data, upload it!
    if (currentClassesStr !== lastFirestoreDataRef.current) {
      setIsSyncing(true);
      const docRef = doc(db, 'users', user.uid);
      setDoc(docRef, { classes }, { merge: true })
        .then(() => {
          lastFirestoreDataRef.current = currentClassesStr;
          setIsSyncing(false);
          setSyncErrorMessage(null);
        })
        .catch((err) => {
          setIsSyncing(false);
          console.error("Firestore sync write failed: ", err);
          try {
            handleFirestoreError(err, OperationType.WRITE, `users/${user.uid}`);
          } catch (handledErr) {
            setSyncErrorMessage("Database save rejected. Check Firebase rules.");
          }
        });
    }
  }, [classes, user]);

  const [usernameInput, setUsernameInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');

  const handlePasswordSignIn = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!usernameInput.trim() || !passwordInput) {
      setSyncErrorMessage("Please enter username and password.");
      return;
    }
    setIsSyncing(true);
    setSyncErrorMessage(null);
    
    // Map custom simple username to a firebase-auth email internally
    const email = usernameInput.includes('@') ? usernameInput : `${usernameInput.trim()}@omniclass-cd395.firebaseapp.com`;
    
    try {
      await signInWithEmailAndPassword(auth, email, passwordInput);
    } catch (error: any) {
      // If the user hasn't been created yet, let's automatically registered to make setup seamless!
      if (error.code === 'auth/user-not-found' || error.code === 'auth/invalid-credential') {
        try {
          // Attempt registering since we want seamless cross-device auth with baked-in defaults
          await createUserWithEmailAndPassword(auth, email, passwordInput);
        } catch (createErr: any) {
          console.error("Auto registration failed: ", createErr);
          setSyncErrorMessage(createErr.message || "Invalid credentials. If user exists, check your password.");
        }
      } else {
        console.error("Password sign-in failed:", error);
        setSyncErrorMessage(error.message || "Invalid credentials or login failed.");
      }
    } finally {
      setIsSyncing(false);
    }
  };

  const handleGoogleSignIn = async () => {
    try {
      setSyncErrorMessage(null);
      await signInWithPopup(auth, googleProvider);
    } catch (error: any) {
      console.error("Sign-in failed:", error);
      setSyncErrorMessage(error.message || "Failed to sign in with Google.");
    }
  };

  const handleSignOut = async () => {
    try {
      setSyncErrorMessage(null);
      await signOut(auth);
    } catch (error: any) {
      console.error("Sign-out failed:", error);
      setSyncErrorMessage("Failed to sign out.");
    }
  };

  // Set default checkboxes when a new pending import is staged
  useEffect(() => {
    if (pendingImport) {
      setSelectedShortIds(pendingImport.short.map(s => s.id));
    } else {
      setSelectedShortIds([]);
    }
  }, [pendingImport]);

  // Find active class
  const activeClass = classes.find(c => c.name === activeClassName) || classes[0] || {
    name: 'None',
    students: [],
    attendance: {},
    assessments: [],
    grades: {}
  };

  // Helper metrics for final grade calculating
  const calculateFinalGrade = (studentId: string, currentClass: ClassSection) => {
    if (currentClass.assessments.length === 0) return 0;

    let totalScore = 0;
    currentClass.assessments.forEach(ass => {
      const score = currentClass.grades[studentId]?.[ass.name] ?? 0;
      totalScore += score;
    });

    return Math.round(totalScore * 100) / 100;
  };

  // Cross-section score changes helper for Search Result updates
  const handleScoreChangeCrossSection = (studentId: string, assessmentName: string, scoreStr: string, className: string) => {
    let scoreNum = parseFloat(scoreStr);
    if (isNaN(scoreNum)) {
      scoreNum = 0;
    }

    setClasses(prev => prev.map(cl => {
      if (cl.name === className) {
        const nextStudentScores = cl.grades[studentId] ? { ...cl.grades[studentId] } : {};
        nextStudentScores[assessmentName] = scoreNum;

        return {
          ...cl,
          grades: {
            ...cl.grades,
            [studentId]: nextStudentScores
          }
        };
      }
      return cl;
    }));
  };

  // Creation of a New Class Section database row
  const handleCreateClass = (e: React.FormEvent) => {
    e.preventDefault();
    const formattedName = newClassNameInput.trim();
    if (!formattedName) return;

    if (classes.some(c => c.name.toLowerCase() === formattedName.toLowerCase())) {
      alert(`Class '${formattedName}' already exists.`);
      return;
    }

    const newClass: ClassSection = {
      name: formattedName,
      students: [],
      attendance: {},
      assessments: [],
      grades: {}
    };

    setClasses(prev => [...prev, newClass]);
    setActiveClassName(formattedName);
    setNewClassNameInput('');
    setIsCreatingClass(false);
  };

  // Deletion of the active Class Section
  const handleDeleteActiveSection = () => {
    if (classes.length <= 1) return;
    
    const remainingClasses = classes.filter(c => c.name !== activeClass.name);
    setClasses(remainingClasses);
    setActiveClassName(remainingClasses[0].name);
    setConfirmDeleteClass(false);
  };

  // Manual Roster Add Form submission
  const handleManualAddStudent = (e: React.FormEvent) => {
    e.preventDefault();
    setStudentAddError(null);
    setStudentAddSuccess(null);

    const formatId = studentIdInput.trim();
    const formatFirst = studentFirstInput.trim();
    const formatLast = studentLastInput.trim();

    if (!formatId || !formatFirst || !formatLast) {
      setStudentAddError("All input fields are required.");
      return;
    }

    // Duplicate Check
    if (activeClass.students.some(s => s.id === formatId)) {
      setStudentAddError(`Duplicate error: ID Number '${formatId}' already exists in this roster.`);
      return;
    }

    const newStudent: Student = {
      id: formatId,
      firstName: formatFirst,
      lastName: formatLast,
      fullName: `${formatFirst} ${formatLast}`
    };

    // Append and sort alphabetically by FIRST name
    const updatedRoster = [...activeClass.students, newStudent].sort((a, b) => 
      a.firstName.localeCompare(b.firstName)
    );

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        // Initialize attendance states, grades
        const nextAtt = { ...cl.attendance };
        nextAtt[formatId] = {};
        
        const nextGrades = { ...cl.grades };
        nextGrades[formatId] = {};

        return {
          ...cl,
          students: updatedRoster,
          attendance: nextAtt,
          grades: nextGrades
        };
      }
      return cl;
    }));

    setStudentAddSuccess(`Added student '${formatFirst} ${formatLast}' to the roster successfully!`);
    setStudentIdInput('');
    setStudentFirstInput('');
    setStudentLastInput('');
  };

  // Student Removal
  const handleRemoveStudentSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setDeletionSuccess(null);
    if (!studentToDeleteId) return;

    const studentName = activeClass.students.find(s => s.id === studentToDeleteId)?.fullName || 'Student';

    if (deletionScope === 'current') {
      // Delete only from the active class
      setClasses(prev => prev.map(cl => {
        if (cl.name === activeClass.name) {
          const nextAtt = { ...cl.attendance };
          delete nextAtt[studentToDeleteId];

          const nextGrades = { ...cl.grades };
          delete nextGrades[studentToDeleteId];

          return {
            ...cl,
            students: cl.students.filter(s => s.id !== studentToDeleteId),
            attendance: nextAtt,
            grades: nextGrades
          };
        }
        return cl;
      }));
      setDeletionSuccess(`Successfully removed ${studentName} from '${activeClass.name}'.`);
    } else {
      // Delete globally from ALL classes
      setClasses(prev => prev.map(cl => {
        const nextAtt = { ...cl.attendance };
        delete nextAtt[studentToDeleteId];

        const nextGrades = { ...cl.grades };
        delete nextGrades[studentToDeleteId];

        return {
          ...cl,
          students: cl.students.filter(s => s.id !== studentToDeleteId),
          attendance: nextAtt,
          grades: nextGrades
        };
      }));
      setDeletionSuccess(`Successfully removed ${studentName} from all classes database wide.`);
    }

    setStudentToDeleteId('');
  };

  // Add Attendance Date Column
  const handleAddAttendanceColumn = () => {
    setTakeAttendanceMsgMessage(null);
    const dateStr = attendanceDate;

    // Check if the current section already has this column
    const dates = getAttendanceDates(activeClass);
    if (dates.includes(dateStr)) {
      setTakeAttendanceMsgMessage({
        type: 'warning',
        text: `The working column '${dateStr}' already exists in table. See interactive edits below.`
      });
      return;
    }

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextAtt = { ...cl.attendance };
        cl.students.forEach(st => {
          if (!nextAtt[st.id]) nextAtt[st.id] = {};
          nextAtt[st.id][dateStr] = 'Not Taken';
        });
        return { ...cl, attendance: nextAtt };
      }
      return cl;
    }));

    setTakeAttendanceMsgMessage({
      type: 'success',
      text: `Added column '${dateStr}' structure successfully.`
    });
  };

  // Edit Single attendance cell state choice index
  const handleToggleAttendanceCell = (studentId: string, date: string, currentVal: string) => {
    let nextVal: 'Not Taken' | 'Attended' | 'Absent' = 'Attended';
    if (currentVal === 'Attended') nextVal = 'Absent';
    else if (currentVal === 'Absent') nextVal = 'Not Taken';

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextStudentAtt = cl.attendance[studentId] ? { ...cl.attendance[studentId] } : {};
        nextStudentAtt[date] = nextVal;
        return {
          ...cl,
          attendance: {
            ...cl.attendance,
            [studentId]: nextStudentAtt
          }
        };
      }
      return cl;
    }));
  };

  // Define new numerical assessment item
  const handleAddAssessment = (e: React.FormEvent) => {
    e.preventDefault();
    setAssessFeedback(null);
    const formName = assessNameInput.trim();
    if (!formName) return;

    if (activeClass.assessments.some(a => a.name.toLowerCase() === formName.toLowerCase())) {
      setAssessFeedback({
        type: 'error',
        text: `The scoring metric '${formName}' is already registered.`
      });
      return;
    }

    const newAssessment: Assessment = {
      name: formName,
      weight: assessWeightInput
    };

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        // Init scores as zero
        const nextGrades = { ...cl.grades };
        cl.students.forEach(st => {
          if (!nextGrades[st.id]) nextGrades[st.id] = {};
          nextGrades[st.id][formName] = 0;
        });

        return {
          ...cl,
          assessments: [...cl.assessments, newAssessment],
          grades: nextGrades
        };
      }
      return cl;
    }));

    setAssessFeedback({
      type: 'success',
      text: `Registered '${formName}' assessment with a max scale of ${assessWeightInput}!`
    });
    setAssessNameInput('');
  };

  // Edit Single score in table matrix
  const handleScoreChange = (studentId: string, assessmentName: string, scoreStr: string) => {
    let scoreNum = parseFloat(scoreStr);
    if (isNaN(scoreNum)) {
      scoreNum = 0;
    }

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextStudentScores = cl.grades[studentId] ? { ...cl.grades[studentId] } : {};
        nextStudentScores[assessmentName] = scoreNum;

        return {
          ...cl,
          grades: {
            ...cl.grades,
            [studentId]: nextStudentScores
          }
        };
      }
      return cl;
    }));
  };

  // Excel roster file parser (Drag & Drop + Input File Select)
  const handleRosterFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFileError(null);
    setFileSuccess(null);
    const file = e.target.files?.[0];
    if (!file) return;

    processRosterFile(file);
  };

  const processRosterFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        let jsonRows: any[] = [];

        if (file.name.endsWith('.csv')) {
          const workbook = XLSX.read(data, { type: 'binary' });
          const firstSheet = workbook.SheetNames[0];
          jsonRows = XLSX.utils.sheet_to_json<any>(workbook.Sheets[firstSheet]);
        } else {
          // Excel sheets
          const workbook = XLSX.read(data, { type: 'binary' });
          const firstSheet = workbook.SheetNames[0];
          jsonRows = XLSX.utils.sheet_to_json<any>(workbook.Sheets[firstSheet]);
        }

        if (!jsonRows || jsonRows.length === 0) {
          setFileError("The uploaded file contains zero records or is corrupted.");
          return;
        }

        // Standardize column keys to locate values
        const normalizedRows = jsonRows.map(row => {
          const keys = Object.keys(row);
          const normalized: any = {};
          keys.forEach(k => {
            normalized[k.trim().toLowerCase().replace(/\s+/g, '')] = row[k];
          });
          return { original: row, norm: normalized };
        });

        // Find matches
        const idKeys = ['idnumber', 'id', 'studentid', 'studentnumber'];
        const firstKeys = ['firstname', 'first', 'givenname'];
        const lastKeys = ['lastname', 'last', 'surname'];

        let resolvedIdKeySrc = '';
        let resolvedFirstKeySrc = '';
        let resolvedLastKeySrc = '';

        for (const row of normalizedRows) {
          const keys = Object.keys(row.original);
          keys.forEach(k => {
            const cleanK = k.trim().toLowerCase().replace(/\s+/g, '');
            if (idKeys.includes(cleanK) && !resolvedIdKeySrc) resolvedIdKeySrc = k;
            if (firstKeys.includes(cleanK) && !resolvedFirstKeySrc) resolvedFirstKeySrc = k;
            if (lastKeys.includes(cleanK) && !resolvedLastKeySrc) resolvedLastKeySrc = k;
          });
          if (resolvedIdKeySrc && resolvedFirstKeySrc && resolvedLastKeySrc) break;
        }

        if (!resolvedIdKeySrc || !resolvedFirstKeySrc || !resolvedLastKeySrc) {
          setFileError(`Columns matching 'ID Number', 'First Name', and 'Last Name' were not located. Found keys: ${Object.keys(jsonRows[0]).join(', ')}`);
          return;
        }

        const parsedStudents: Student[] = jsonRows.map(r => {
          const fName = String(r[resolvedFirstKeySrc] || '').trim();
          const lName = String(r[resolvedLastKeySrc] || '').trim();
          return {
            id: String(r[resolvedIdKeySrc] || '').trim(),
            firstName: fName,
            lastName: lName,
            fullName: `${fName} ${lName}`
          };
        }).filter(st => st.id && st.firstName);

        if (parsedStudents.length === 0) {
          setFileError("Failed to parse valid student records containing names and ID numbers.");
          return;
        }

        // Sort dynamically on First Name
        const sortedStudents = parsedStudents.sort((a, b) => 
          a.firstName.localeCompare(b.firstName)
        );

        // Exclude / Stage students with ID lengths <= 6
        const shortIds = sortedStudents.filter(st => st.id.length <= 6);
        const normalIds = sortedStudents.filter(st => st.id.length > 6);

        if (shortIds.length > 0) {
          setPendingImport({
            className: activeClass.name,
            normal: normalIds,
            short: shortIds
          });
          setFileSuccess("Verification check required for short student IDs.");
        } else {
          // Direct commit
          setClasses(prev => prev.map(cl => {
            if (cl.name === activeClass.name) {
              const nextAtt: any = {};
              const nextGrades: any = {};
              sortedStudents.forEach(st => {
                nextAtt[st.id] = cl.attendance[st.id] || {};
                nextGrades[st.id] = cl.grades[st.id] || {};
              });

              return {
                ...cl,
                students: sortedStudents,
                attendance: nextAtt,
                grades: nextGrades
              };
            }
            return cl;
          }));
          setFileSuccess(`Successfully ingested offline roster! Parsed ${sortedStudents.length} entries sorted alphabetically.`);
        }

      } catch (err: any) {
        setFileError(`Error parsing workbook: ${err?.message || 'Unsupported format'}`);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Commit Staged ID review checkboxes
  const handleConfirmPendingImport = () => {
    if (!pendingImport) return;

    // Filter accepted short students
    const approvedShortStudents = pendingImport.short.filter(st => 
      selectedShortIds.includes(st.id)
    );

    const consolidatedList = [...pendingImport.normal, ...approvedShortStudents].sort((a, b) => 
      a.firstName.localeCompare(b.firstName)
    );

    setClasses(prev => prev.map(cl => {
      if (cl.name === pendingImport.className) {
        const nextAtt: any = {};
        const nextGrades: any = {};
        consolidatedList.forEach(st => {
          nextAtt[st.id] = cl.attendance[st.id] || {};
          nextGrades[st.id] = cl.grades[st.id] || {};
        });

        return {
          ...cl,
          students: consolidatedList,
          attendance: nextAtt,
          grades: nextGrades
        };
      }
      return cl;
    }));

    setPendingImport(null);
    setFileSuccess(`Successfully uploaded filtered roster containing ${consolidatedList.length} records.`);
  };

  const handleDiscardPendingImportShortIds = () => {
    if (!pendingImport) return;

    // Direct sort on normal entries
    const consolidatedList = [...pendingImport.normal].sort((a, b) => 
      a.firstName.localeCompare(b.firstName)
    );

    setClasses(prev => prev.map(cl => {
      if (cl.name === pendingImport.className) {
        const nextAtt: any = {};
        const nextGrades: any = {};
        consolidatedList.forEach(st => {
          nextAtt[st.id] = cl.attendance[st.id] || {};
          nextGrades[st.id] = cl.grades[st.id] || {};
        });

        return {
          ...cl,
          students: consolidatedList,
          attendance: nextAtt,
          grades: nextGrades
        };
      }
      return cl;
    }));

    setPendingImport(null);
    setFileSuccess(`Successfully uploaded roster (completely skipped all ≤ 6 digit ID matches) containing ${consolidatedList.length} students.`);
  };

  // Extract date columns from attendance ledger
  const getAttendanceDates = (cl: ClassSection) => {
    const datesSet = new Set<string>();
    Object.values(cl.attendance).forEach(record => {
      Object.keys(record).forEach(d => datesSet.add(d));
    });
    return Array.from(datesSet).sort();
  };

  // Single Class Exporters
  const handleExportGrades = () => {
    const tableData = activeClass.students.map(s => {
      const row: any = {
        'ID Number': s.id,
        'Full Name': s.fullName
      };
      
      activeClass.assessments.forEach(ass => {
        const colTitle = ass.name;
        row[colTitle] = activeClass.grades[s.id]?.[ass.name] ?? 0;
      });

      row['Final Grade'] = calculateFinalGrade(s.id, activeClass);
      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(tableData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Grades Roster');
    
    XLSX.writeFile(workbook, `Grades_Report_${activeClass.name.replace(/\s+/g, '_')}.xlsx`);
  };

  const handleExportAttendance = () => {
    const dates = getAttendanceDates(activeClass);
    const tableData = activeClass.students.map(s => {
      const row: any = {
        'ID Number': s.id,
        'Full Name': s.fullName
      };

      dates.forEach(d => {
        row[d] = activeClass.attendance[s.id]?.[d] ?? 'Not Taken';
      });

      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(tableData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Attendance Ledger');

    XLSX.writeFile(workbook, `Attendance_Log_${activeClass.name.replace(/\s+/g, '_')}.xlsx`);
  };

  // GLOBAL MULTI-CLASS MERGED DOWNLOADING SYSTEM
  // Downloads structured sheets containing exactly 2 lines of empty space between class sections!
  const handleExportMerged = (type: 'grades' | 'attendance') => {
    const aoa: any[][] = [];

    classes.forEach(cl => {
      if (cl.students.length === 0) return;

      // 1. Add Section Banner Title Row
      aoa.push([`🏫 ${cl.name}`]);

      // 2. Add Subheaders row
      if (type === 'attendance') {
        const dates = getAttendanceDates(cl);
        const headers = ['ID Number', 'Full Name', ...dates];
        aoa.push(headers);

        cl.students.forEach(st => {
          const row: any[] = [st.id, st.fullName];
          dates.forEach(d => {
            row.push(cl.attendance[st.id]?.[d] ?? 'Not Taken');
          });
          aoa.push(row);
        });
      } else {
        // Grades Roster
        const assessCols = cl.assessments.map(a => a.name);
        const headers = ['ID Number', 'Full Name', ...assessCols, 'Final Grade'];
        aoa.push(headers);

        cl.students.forEach(st => {
          const row: any[] = [st.id, st.fullName];
          cl.assessments.forEach(ass => {
            row.push(cl.grades[st.id]?.[ass.name] ?? 0);
          });
          row.push(calculateFinalGrade(st.id, cl));
          aoa.push(row);
        });
      }

      // 3. Write exactly 2 lines of empty cell structures
      aoa.push([]);
      aoa.push([]);
    });

    if (aoa.length === 0) {
      alert("No rosters containing student listings found across sections.");
      return;
    }

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `Merged_${type.charAt(0).toUpperCase() + type.slice(1)}`);

    XLSX.writeFile(workbook, `Merged_${type.charAt(0).toUpperCase() + type.slice(1)}_All_Classes.xlsx`);
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(PYTHON_STREAMLIT_CODE);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2000);
  };

  // Cross-section student search logic
  const getGlobalSearchResults = () => {
    if (!globalSearchQuery.trim()) return [];
    const query = globalSearchQuery.trim().toLowerCase();
    
    const results: Array<{
      student: Student;
      classSection: ClassSection;
    }> = [];

    classes.forEach(cl => {
      cl.students.forEach(st => {
        if (
          st.fullName.toLowerCase().includes(query) ||
          st.id.toLowerCase().includes(query)
        ) {
          results.push({
            student: st,
            classSection: cl
          });
        }
      });
    });

    return results;
  };

  const globalSearchResults = getGlobalSearchResults();

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col text-slate-800 font-sans antialiased" id="streamlit-app-container">
      {/* Streamlit Brand Accent Strip */}
      <div className="h-1 bg-red-500 w-full" id="streamlit-accent-line"></div>

      {/* Header Banner */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4" id="streamlit-header">
        <div className="flex items-center gap-3">
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className="p-2 hover:bg-slate-100 rounded-lg text-slate-600 transition-colors mr-1 cursor-pointer flex items-center justify-center border border-slate-200 shadow-sm"
            id="toggle-sidebar-hamburger"
            title={isSidebarOpen ? "Hide Sidebar Settings" : "Show Sidebar Settings"}
          >
            <Menu className="h-5 w-5" />
          </button>
          <div className="p-2 bg-red-50 py-2.5 rounded text-red-500">
            <GraduationCap className="h-7 w-7" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Streamlit Class Operations Panel</h1>
            <p className="text-xs text-slate-500">Dual Execution: Live Interactive Applet + Pure Python Source Exporter</p>
          </div>
        </div>
        
        <div className="flex items-center gap-3 flex-wrap">
          <button 
            onClick={() => setActiveTab(activeTab === 'code' ? 'students' : 'code')}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition-colors border ${
              activeTab === 'code' 
                ? 'bg-amber-500 text-white border-amber-500' 
                : 'bg-white text-slate-700 hover:bg-slate-100 border-slate-300'
            }`}
            id="toggle-source-code-btn"
          >
            <Code className="h-3.5 w-3.5" />
            {activeTab === 'code' ? 'Back to Class App' : 'View streamlit app.py Code'}
          </button>
          
          <span className="inline-flex bg-slate-100 text-slate-600 font-mono text-[10px] uppercase font-bold items-center px-2 py-1 rounded border border-slate-200 gap-1">
            <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-ping"></span>
            st.session_state is active
          </span>
        </div>
      </header>

      {/* Interface workspace */}
      <div className="flex-1 flex flex-col lg:flex-row" id="streamlit-workspace">
        {/* SIDEBAR WIDGET PANELS */}
        <aside className={`${isSidebarOpen ? 'flex' : 'hidden'} w-full lg:w-80 bg-white border-r border-slate-200 p-5 flex-col shrink-0 gap-6`} id="streamlit-sidebar">
          
          {/* Firebase Sync Widget */}
          <div className="bg-slate-50/80 p-3.5 rounded-xl border border-slate-200 shadow-3xs flex flex-col gap-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                <Cloud className={`h-4 w-4 ${user ? 'text-indigo-600' : 'text-slate-400'}`} />
                Cloud Database Sync
              </span>
              
              {user && (
                <span className={`inline-flex items-center px-1.5 py-0.5 rounded-full text-[9px] font-bold ${
                  isSyncing ? 'bg-amber-50 text-amber-700 animate-pulse' : 'bg-emerald-50 text-emerald-700'
                }`}>
                  <span className={`w-1 h-1 rounded-full mr-1 ${isSyncing ? 'bg-amber-500 animate-ping' : 'bg-emerald-500'}`} />
                  {isSyncing ? "Syncing..." : "Synced"}
                </span>
              )}
            </div>

            {authLoading ? (
              <div className="flex items-center gap-2 text-xs text-slate-400 font-medium">
                <RefreshCw className="h-3 w-3 animate-spin text-slate-500" />
                <span>Checking cloud sync...</span>
              </div>
            ) : user ? (
              <div className="flex flex-col gap-2">
                <div className="text-[11px] text-slate-500 bg-white p-2.5 rounded-lg border border-slate-150 flex flex-col gap-1 shadow-3xs">
                  <div className="font-semibold text-slate-800 truncate flex items-center gap-1.5" title={user.email || ""}>
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                    <span className="truncate">{user.displayName || user.email?.split('@')[0] || "Logged In"}</span>
                  </div>
                  <div className="text-[10px] text-slate-400 truncate font-mono">{user.email}</div>
                </div>

                <button
                  onClick={handleSignOut}
                  className="w-full flex items-center justify-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 border border-rose-200 hover:border-rose-300 text-rose-700 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-3xs"
                >
                  <LogOut className="h-3.5 w-3.5" />
                  Sign Out / Disconnect
                </button>
              </div>
            ) : (
              <form onSubmit={handlePasswordSignIn} className="flex flex-col gap-2.5">
                <p className="text-[11px] text-slate-500 leading-relaxed font-semibold">
                  Sign in to keep your changes synced across all devices and browsers (work and home PCs).
                </p>

                <div className="flex flex-col gap-1 text-[11px]">
                  <span className="text-slate-600 font-bold uppercase tracking-wider text-[9px]">Username</span>
                  <input
                    type="text"
                    value={usernameInput}
                    onChange={(e) => setUsernameInput(e.target.value)}
                    placeholder="e.g. workshop2"
                    className="w-full text-xs px-2.5 py-1.5 rounded border border-slate-200 bg-white shadow-3xs focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono"
                  />
                </div>

                <div className="flex flex-col gap-1 text-[11px]">
                  <span className="text-slate-600 font-bold uppercase tracking-wider text-[9px]">Password</span>
                  <input
                    type="password"
                    value={passwordInput}
                    onChange={(e) => setPasswordInput(e.target.value)}
                    placeholder="Password"
                    className="w-full text-xs px-2.5 py-1.5 rounded border border-slate-200 bg-white shadow-3xs focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono"
                  />
                </div>

                <button
                  type="submit"
                  disabled={isSyncing}
                  className="w-full flex items-center justify-center gap-1.5 px-3 py-2 bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] text-white rounded-lg text-xs font-bold transition-all cursor-pointer shadow-xs border border-indigo-700 mt-1 disabled:opacity-50"
                >
                  {isSyncing ? (
                    <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                  ) : (
                    <LogIn className="h-3.5 w-3.5" />
                  )}
                  Sign In (Username & Password)
                </button>

                <div className="flex items-center gap-2 my-1">
                  <div className="flex-1 h-px bg-slate-200"></div>
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">or</span>
                  <div className="flex-1 h-px bg-slate-200"></div>
                </div>

                <button
                  type="button"
                  onClick={handleGoogleSignIn}
                  className="w-full flex items-center justify-center gap-2 px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-semibold transition-all cursor-pointer shadow-3xs border border-slate-200"
                >
                  <LogIn className="h-3 w-3 text-slate-500" />
                  Sign In with Google
                </button>
              </form>
            )}

            {syncErrorMessage && (
              <div className="bg-red-50 text-red-700 border border-red-100 p-2 rounded-lg text-[10.5px] font-medium leading-normal flex items-start gap-1.5">
                <AlertCircle className="h-3.5 w-3.5 shrink-0 mt-0.5" />
                <span>{syncErrorMessage}</span>
              </div>
            )}
          </div>

          {/* Section Selector */}

          <div className="flex flex-col gap-2" id="sidebar-class-selector">
            <label className="text-xs font-extrabold text-slate-600 uppercase tracking-widest flex items-center justify-between">
              Active Class Section
              <span className="text-[10px] font-normal text-slate-400 capitalize">selected_class</span>
            </label>
            
            {isCreatingClass ? (
              <form onSubmit={handleCreateClass} className="flex flex-col gap-2 bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                <span className="text-xs font-semibold text-slate-700">Enter Class Name:</span>
                <input 
                  type="text" 
                  value={newClassNameInput}
                  onChange={e => setNewClassNameInput(e.target.value)}
                  placeholder="e.g. Physics 301" 
                  className="w-full text-xs px-2.5 py-1.5 rounded border border-slate-300 bg-white shadow-sm focus:outline-none focus:ring-1 focus:ring-red-500"
                  autoFocus
                />
                <div className="flex items-center gap-2 justify-end">
                  <button 
                    type="button" 
                    onClick={() => setIsCreatingClass(false)}
                    className="text-[10px] text-slate-500 px-2 py-1 hover:bg-slate-100 rounded"
                  >
                    Cancel
                  </button>
                  <button 
                    type="submit" 
                    className="bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] px-2.5 py-1 rounded font-semibold"
                  >
                    Create
                  </button>
                </div>
              </form>
            ) : (
              <div className="flex gap-1.5">
                <select
                  value={activeClassName}
                  onChange={e => {
                    setActiveClassName(e.target.value);
                    setConfirmDeleteClass(false);
                  }}
                  className="flex-1 text-sm bg-slate-50 border border-slate-300 rounded-lg p-2 focus:ring-1 focus:ring-red-500 outline-none text-slate-800 font-medium"
                >
                  {classes.map(c => (
                    <option key={c.name} value={c.name}>{c.name}</option>
                  ))}
                </select>
                <button
                  onClick={() => setIsCreatingClass(true)}
                  className="p-2 bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 rounded-lg transition-colors flex items-center justify-center"
                  title="Create New Class"
                  id="create-new-class-sidebar-btn"
                >
                  <Plus className="h-4 w-4" />
                </button>
              </div>
            )}
            <p className="text-[10px] text-slate-400 font-mono">Current Roster size: {activeClass.students.length} students</p>
          </div>

          {/* Delete Active Section Controller */}
          {classes.length > 1 && (
            <div className="flex justify-end -mt-3.5" id="sidebar-delete-class-box">
              {!confirmDeleteClass ? (
                <button
                  onClick={() => setConfirmDeleteClass(true)}
                  className="inline-flex items-center gap-1 text-[10px] text-slate-400 hover:text-rose-600 transition-colors font-medium border border-slate-100 hover:border-slate-200 px-2 py-0.5 rounded bg-slate-50/50"
                >
                  <Trash2 className="h-2.5 w-2.5" />
                  Delete active section
                </button>
              ) : (
                <div className="flex flex-col gap-1.5 p-2 rounded-md bg-slate-50 border border-slate-200 text-left w-full">
                  <p className="text-[9px] text-slate-600 leading-tight">
                    Confirm clearing <strong>'{activeClass.name}'</strong>? All attendance and grades will be deleted.
                  </p>
                  <div className="flex items-center gap-1.5 justify-end">
                    <button
                      onClick={() => setConfirmDeleteClass(false)}
                      className="text-[9px] bg-white border border-slate-200 rounded text-slate-500 px-1.5 py-0.5 hover:bg-slate-100 shrink-0"
                    >
                      Cancel
                    </button>
                    <button
                      onClick={handleDeleteActiveSection}
                      className="text-[9px] bg-rose-600 text-white rounded px-2 py-0.5 hover:bg-rose-700 font-bold shrink-0 shadow-xs"
                    >
                      Yes, Delete
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          <hr className="border-slate-100" />

          {/* Roster Ingestion Upload */}
          <div className="flex flex-col gap-2" id="sidebar-roster-uploader">
            <h3 className="text-xs font-extrabold text-slate-600 uppercase tracking-widest flex items-center gap-1">
              <Upload className="h-3 w-3 text-red-500" />
              Upload Student Roster
            </h3>
            <p className="text-[10.5px] text-slate-500 leading-normal">
              Accepts .csv or .xlsx with 'First Name', 'Last Name', and 'ID Number'.
            </p>
            
            <label className="border-2 border-dashed border-slate-200 hover:border-red-300 rounded-xl p-4 flex flex-col items-center justify-center text-center cursor-pointer transition-colors bg-slate-50 hover:bg-red-50">
              <input 
                type="file" 
                accept=".csv, .xlsx" 
                onChange={handleRosterFileUpload} 
                className="hidden" 
              />
              <FileSpreadsheet className="h-6 w-6 text-slate-400 mb-1.5" />
              <span className="text-xs font-semibold text-slate-700">Choose file or drag here</span>
              <span className="text-[9.5px] text-slate-400 mt-1">Strips outer spacing and title-cases</span>
            </label>

            {fileError && (
              <div className="mt-1.5 text-[10px] bg-rose-50 text-rose-700 p-2 rounded border border-rose-200 flex items-start gap-1">
                <AlertCircle className="h-3 w-3 shrink-0 mt-0.5" />
                <span>{fileError}</span>
              </div>
            )}

            {fileSuccess && (
              <div className="mt-1.5 text-[10px] bg-emerald-50 text-emerald-800 p-2 rounded border border-emerald-100 flex items-start gap-1">
                <Check className="h-3 w-3 shrink-0 mt-0.5 text-emerald-600" />
                <span>{fileSuccess}</span>
              </div>
            )}

            {/* Quick Demo Roster Seed helper */}
            <button
              onClick={() => {
                setClasses(prev => prev.map(cl => {
                  if (cl.name === activeClass.name) {
                    return {
                      ...cl,
                      students: [
                        { id: "101", firstName: "Sophia", lastName: "Alvarez", fullName: "Sophia Alvarez" },
                        { id: "102", firstName: "Benjamin", lastName: "Chen", fullName: "Benjamin Chen" },
                        { id: "103", firstName: "Emma", lastName: "Dmitriev", fullName: "Emma Dmitriev" },
                        { id: "9001", firstName: "James", lastName: "Bond", fullName: "James Bond" },
                        { id: "404", firstName: "Ada", lastName: "Lovelace", fullName: "Ada Lovelace" }
                      ].sort((a,b)=> a.firstName.localeCompare(b.firstName)),
                      attendance: {
                        "101": { "2026-05-24": "Attended" },
                        "102": { "2026-05-24": "Absent" },
                        "103": { "2026-05-24": "Not Taken" },
                        "9001": { "2026-05-24": "Attended" }
                      },
                      grades: {
                        "101": { "Quiz 1": 95, "Midterm": 88 },
                        "102": { "Quiz 1": 82, "Midterm": 79 },
                        "103": { "Quiz 1": 90, "Midterm": 92 }
                      }
                    };
                  }
                  return cl;
                }));
                setFileSuccess("Seeded mock roster details to current working class!");
              }}
              className="text-[10px] text-slate-500 text-right hover:text-slate-800 hover:underline"
              type="button"
            >
              ⚡ Fill Quick Mock Student Listing
            </button>
          </div>

          <hr className="border-slate-100" />

          {/* Manual Add Input components */}
          <div className="flex flex-col gap-3" id="sidebar-manual-add">
            <h3 className="text-xs font-extrabold text-slate-600 uppercase tracking-widest flex items-center gap-1.5">
              <UserPlus className="h-3.5 w-3.5 text-red-500" />
              Add Student Record
            </h3>

            <form onSubmit={handleManualAddStudent} className="flex flex-col gap-2">
              <input 
                type="text" 
                value={studentIdInput}
                onChange={e => setStudentIdInput(e.target.value)}
                placeholder="ID Number (e.g. 108)" 
                className="w-full text-xs px-2.5 py-1.5 rounded border border-slate-300 bg-slate-50 text-slate-800"
              />
              <div className="grid grid-cols-2 gap-2">
                <input 
                  type="text" 
                  value={studentFirstInput}
                  onChange={e => setStudentFirstInput(e.target.value)}
                  placeholder="First Name" 
                  className="w-full text-xs px-2.5 py-1.5 rounded border border-slate-300 bg-slate-50 text-slate-800"
                />
                <input 
                  type="text" 
                  value={studentLastInput}
                  onChange={e => setStudentLastInput(e.target.value)}
                  placeholder="Last Name" 
                  className="w-full text-xs px-2.5 py-1.5 rounded border border-slate-300 bg-slate-50 text-slate-800"
                />
              </div>

              <button
                type="submit"
                className="w-full bg-slate-800 hover:bg-slate-900 border border-slate-900 text-white font-bold text-xs py-2 rounded-lg transition-colors cursor-pointer"
                id="add-student-btn-sidebar"
              >
                Add Student to Roster
              </button>
            </form>

            {studentAddError && (
              <div className="text-[10px] text-rose-700 bg-rose-50 border border-rose-100 p-2 rounded">
                {studentAddError}
              </div>
            )}

            {studentAddSuccess && (
              <div className="text-[10px] text-emerald-800 bg-emerald-50 border border-emerald-100 p-2 rounded">
                {studentAddSuccess}
              </div>
            )}
          </div>

          <hr className="border-slate-100" />

          {/* Remove Student Section */}
          <div className="flex flex-col gap-3" id="sidebar-manual-removal">
            <h3 className="text-xs font-extrabold text-slate-600 uppercase tracking-widest flex items-center gap-1.5">
              <Trash2 className="h-3.5 w-3.5 text-red-500" />
              Remove Student
            </h3>

            {activeClass.students.length === 0 ? (
              <p className="text-[10px] text-slate-400 italic leading-normal">
                No active students registered to select from.
              </p>
            ) : (
              <form onSubmit={handleRemoveStudentSubmit} className="flex flex-col gap-2">
                <select
                  value={studentToDeleteId}
                  onChange={e => setStudentToDeleteId(e.target.value)}
                  className="w-full text-xs bg-slate-50 border border-slate-300 rounded p-2 text-slate-700 focus:ring-1 outline-none font-medium"
                >
                  <option value="">-- Choose Student to Delete --</option>
                  {activeClass.students.map(s => (
                    <option key={s.id} value={s.id}>{s.id} - {s.fullName}</option>
                  ))}
                </select>

                {/* Radio options */}
                <div className="flex flex-col gap-1.5 mt-1 text-xs">
                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Scope of Deletion:</span>
                  <label className="flex items-center gap-2 text-slate-700 cursor-pointer">
                    <input 
                      type="radio" 
                      name="deletionScope" 
                      checked={deletionScope === 'current'}
                      onChange={() => setDeletionScope('current')}
                      className="accent-slate-800" 
                    />
                    <span>Remove from current class section</span>
                  </label>
                  <label className="flex items-center gap-2 text-slate-700 cursor-pointer">
                    <input 
                      type="radio" 
                      name="deletionScope" 
                      checked={deletionScope === 'all'}
                      onChange={() => setDeletionScope('all')}
                      className="accent-slate-800" 
                    />
                    <span>Remove from ALL classes globally</span>
                  </label>
                </div>

                <button
                  type="submit"
                  disabled={!studentToDeleteId}
                  className={`w-full text-xs font-bold py-2 rounded transition-colors flex items-center justify-center gap-1 border ${
                    studentToDeleteId 
                      ? 'bg-rose-600 hover:bg-rose-700 text-white border-rose-700 cursor-pointer' 
                      : 'bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed'
                  }`}
                  id="delete-student-submit-btn"
                >
                  <Trash2 className="h-3 w-3" />
                  Execute Deletion
                </button>
              </form>
            )}

            {deletionSuccess && (
              <div className="text-[10px] text-emerald-800 bg-emerald-50 border border-emerald-100 p-2 rounded flex items-start gap-1">
                <Check className="h-3 w-3 shrink-0 mt-0.5 text-emerald-600" />
                <span>{deletionSuccess}</span>
              </div>
            )}
          </div>

          <hr className="border-slate-100" />

          {/* Export Data section in sidebar */}
          <div className="mt-auto pt-2 border-t border-slate-100 flex flex-col gap-3" id="sidebar-exports">
            <h3 className="text-xs font-extrabold text-slate-600 uppercase tracking-widest flex items-center gap-1">
              <Download className="h-3 w-3 text-red-500" />
              Download Ledgers
            </h3>
            
            <div className="flex flex-col gap-2">
              <span className="text-[9px] font-extrabold text-slate-400 tracking-wider uppercase">Active Class:</span>
              <button
                onClick={handleExportGrades}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-semibold py-2 px-3 rounded-lg border text-left flex items-center justify-between gap-2 shadow-xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700 hover:text-black cursor-pointer'
                  : 'bg-slate-50 text-slate-400 border-slate-200 cursor-not-allowed'
                }`}
                id="export-grades-excel-sidebar"
              >
                <span className="flex items-center gap-2">
                  <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                  Download Grades (Excel)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
              </button>

              <button
                onClick={handleExportAttendance}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-semibold py-2 px-3 rounded-lg border text-left flex items-center justify-between gap-2 shadow-xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700 hover:text-black cursor-pointer'
                  : 'bg-slate-50 text-slate-400 border-slate-200 cursor-not-allowed'
                }`}
                id="export-attendance-excel-sidebar"
              >
                <span className="flex items-center gap-2">
                  <Calendar className="h-4 w-4 text-sky-600" />
                  Download Attendance Log
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
              </button>

              {/* Merged Exports System */}
              <span className="text-[9px] font-extrabold text-slate-400 tracking-wider uppercase mt-2">Combined All Classes (with 2 empty row breaks):</span>
              <button
                onClick={() => handleExportMerged('grades')}
                className="w-full text-xs font-semibold py-2 px-3 rounded-lg border text-left flex items-center justify-between gap-2 shadow-xs bg-indigo-50 border-indigo-100 hover:bg-indigo-100 text-indigo-800 transition-all cursor-pointer"
                id="export-merged-grades-sidebar"
              >
                <span className="flex items-center gap-2">
                  <FileSpreadsheet className="h-4 w-4 text-indigo-600" />
                  Download Merged Grades
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-indigo-400" />
              </button>

              <button
                onClick={() => handleExportMerged('attendance')}
                className="w-full text-xs font-semibold py-2 px-3 rounded-lg border text-left flex items-center justify-between gap-2 shadow-xs bg-teal-50 border-teal-100 hover:bg-teal-100 text-teal-800 transition-all cursor-pointer"
                id="export-merged-attendance-sidebar"
              >
                <span className="flex items-center gap-2">
                  <Calendar className="h-4 w-4 text-teal-600" />
                  Download Merged Attendance
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-teal-400" />
              </button>
            </div>
            
            <span className="text-[10px] text-slate-400 leading-normal">
              Attendance exports keep text values representing status colors. Saved in xlsx.
            </span>
          </div>
        </aside>

        {/* MAIN DASHBOARD PANEL */}
        <main className="flex-1 p-6 md:p-8 flex flex-col gap-6 overflow-x-auto" id="streamlit-main">
          {/* Active section ribbon banner */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col sm:flex-row justify-between sm:items-center gap-4" id="main-ribbon">
            <div>
              <div className="text-[10px] font-bold text-slate-400 tracking-widest uppercase">active working context</div>
              <h2 className="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2 mt-1">
                {activeClass.name}
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                All changes reflect immediately under simulated <strong>st.session_state</strong> inside React client memory.
              </p>
            </div>
            <div className="flex gap-4">
              <div className="bg-slate-50 px-4 py-2 border border-slate-100 rounded-lg text-center min-w-24">
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Students</div>
                <div className="text-lg font-bold text-slate-800">{activeClass.students.length}</div>
              </div>
              <div className="bg-slate-50 px-4 py-2 border border-slate-100 rounded-lg text-center min-w-24">
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Assessments</div>
                <div className="text-lg font-bold text-slate-800">{activeClass.assessments.length}</div>
              </div>
            </div>
          </div>

          {/* PYTHON SOURCE CODE CONTAINER (Toggled open or on Tab trigger) */}
          {activeTab === 'code' ? (
            <div className="bg-slate-900 text-slate-100 rounded-xl shadow-md border border-slate-800 flex flex-col overflow-hidden" id="python-source-viewer">
              <div className="bg-slate-800 px-4 py-3 border-b border-slate-700 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Code className="h-4 w-4 text-amber-500" />
                  <span className="text-xs font-mono font-bold text-slate-200">/app.py</span>
                  <span className="text-[10px] bg-slate-900 text-slate-400 border border-slate-700 px-1.5 py-0.5 rounded">Python 3.10+ Streamlit</span>
                </div>
                <button
                  onClick={copyToClipboard}
                  className="bg-slate-700 hover:bg-slate-600 text-slate-200 hover:text-white px-3 py-1 rounded text-xs inline-flex items-center gap-1 transition-colors"
                  id="copy-python-code-btn"
                >
                  <Copy className="h-3.5 w-3.5" />
                  {copiedCode ? 'Copied!' : 'Copy Code'}
                </button>
              </div>
              <pre className="p-6 text-xs text-slate-300 font-mono overflow-auto max-h-[600px] leading-relaxed select-all">
                {PYTHON_STREAMLIT_CODE}
              </pre>
            </div>
          ) : (
            <div className="flex flex-col gap-6" id="app-working-views">
              
              {/* STAGED SHORT ID VERIFICATION CONTAINER */}
              {pendingImport && (
                <div className="bg-amber-50 rounded-xl border border-amber-200 shadow-sm p-6 flex flex-col gap-4 animate-fade-in" id="short-id-compliance-container">
                  <div className="flex items-start gap-3">
                    <AlertCircle className="h-6 w-6 text-amber-600 shrink-0 mt-0.5" />
                    <div>
                      <h3 className="text-md font-bold text-amber-900 tracking-tight">
                        Roster Verification Required: Short Student IDs Detected
                      </h3>
                      <p className="text-xs text-amber-700 mt-1 leading-normal">
                        We located student entries inside import with ID numbers consisting of <strong>6 or less digits</strong>.
                        Choose which students you want to whitelist and keep in your roster. Deselected whitelists will be omitted:
                      </p>
                    </div>
                  </div>

                  {/* Grid lists with checkbox whitelisting */}
                  <div className="bg-white border border-amber-100 rounded-lg max-h-56 overflow-y-auto divide-y divide-slate-100">
                    {pendingImport.short.map(st => {
                      const isChecked = selectedShortIds.includes(st.id);
                      return (
                        <div key={st.id} className="flex items-center justify-between p-3 text-xs hover:bg-amber-50/50 transition-colors">
                          <div className="flex items-center gap-3">
                            <input 
                              type="checkbox" 
                              checked={isChecked}
                              onChange={() => {
                                if (isChecked) {
                                  setSelectedShortIds(prev => prev.filter(x => x !== st.id));
                                } else {
                                  setSelectedShortIds(prev => [...prev, st.id]);
                                }
                              }}
                              className="h-4 w-4 text-amber-600 focus:ring-amber-500 rounded border-slate-300 cursor-pointer"
                            />
                            <div>
                              <span className="font-mono bg-amber-100 text-amber-900 px-1.5 py-0.5 rounded font-semibold text-[10.5px]">ID: {st.id}</span>
                              <span className="font-semibold text-slate-700 ml-2.5">{st.fullName}</span>
                            </div>
                          </div>
                          <span className="text-[10px] text-slate-400 font-medium">({st.id.length} digits)</span>
                        </div>
                      );
                    })}
                  </div>

                  {/* Checkbox triggers control buttons */}
                  <div className="flex items-center gap-3 flex-wrap">
                    <button
                      onClick={handleConfirmPendingImport}
                      className="bg-amber-600 hover:bg-amber-700 border border-amber-700 text-white font-bold text-xs px-4 py-2 rounded-lg shadow-sm transition-colors cursor-pointer"
                    >
                      Keep whitelisted ({selectedShortIds.length}) & Import All
                    </button>
                    <button
                      onClick={handleDiscardPendingImportShortIds}
                      className="bg-white hover:bg-amber-100 border border-amber-200 text-amber-850 text-xs font-semibold px-4 py-2 rounded-lg transition-colors cursor-pointer"
                    >
                      Discard All {pendingImport.short.length} Short IDs & Save Others
                    </button>
                    <button
                      onClick={() => setPendingImport(null)}
                      className="bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-semibold px-4 py-2 rounded-lg transition-colors cursor-pointer"
                    >
                      Abort Import
                    </button>
                  </div>
                </div>
              )}

              {/* GLOBAL CROSS-SECTION STUDENT SEARCH CARD */}
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col gap-4" id="global-student-search-container">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-1">
                  <div className="flex items-center gap-2">
                    <span className="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg">
                      <Search className="h-4.5 w-4.5" />
                    </span>
                    <div>
                      <h3 className="font-extrabold text-slate-900 text-sm md:text-md tracking-tight">Global Student Search</h3>
                      <p className="text-[11px] text-slate-400">Search student directory across all sections and modify exam marks instantly</p>
                    </div>
                  </div>
                  {globalSearchQuery && (
                    <button 
                      onClick={() => setGlobalSearchQuery('')}
                      className="text-[10px] text-indigo-600 hover:text-indigo-800 font-bold bg-indigo-50 hover:bg-indigo-100 px-2.5 py-1 rounded-md transition-all self-start sm:self-auto cursor-pointer"
                    >
                      Clear Search
                    </button>
                  )}
                </div>

                <div className="relative">
                  <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                    <Search className="h-4 w-4" />
                  </span>
                  <input
                    type="text"
                    value={globalSearchQuery}
                    onChange={e => setGlobalSearchQuery(e.target.value)}
                    placeholder="Type Student Name or ID (e.g. Liam, Chen, or 105) to search all sections..."
                    className="w-full pl-10 pr-4 py-2.5 text-xs md:text-sm bg-slate-50 border border-slate-300 rounded-xl outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 focus:bg-white text-slate-800 placeholder-slate-400 font-medium transition-all shadow-sm"
                  />
                </div>

                {globalSearchQuery.trim() !== '' && (
                  <div className="mt-1 flex flex-col gap-3">
                    <div className="flex items-center justify-between border-b border-slate-100 pb-1">
                      <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                        Search Results ({globalSearchResults.length} matches)
                      </span>
                    </div>

                    {globalSearchResults.length === 0 ? (
                      <div className="p-8 text-center bg-slate-50 border border-dashed rounded-xl text-slate-450 text-xs">
                        No student directory registrations matching <strong className="text-slate-700">"{globalSearchQuery}"</strong> located in any class sections.
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        {globalSearchResults.map(({ student, classSection }) => {
                          const assessments = classSection.assessments;
                          const finalGrade = calculateFinalGrade(student.id, classSection);
                          const colorInfo = getStudentHighlightClasses(student.color);
                          const cardBorder = student.color
                            ? `border border-slate-200 border-l-[5px] ${colorInfo.border?.replace('border-l-4 ', '') || ''}`
                            : 'border border-slate-200 hover:border-slate-350';

                          return (
                            <div 
                              key={`${classSection.name}-${student.id}`} 
                              className={`${colorInfo.bg || 'bg-slate-50/50 hover:bg-slate-50/80'} ${cardBorder} rounded-xl p-4 flex flex-col gap-3.5 relative transition-all shadow-xs`}
                            >
                              {/* Header details */}
                              <div className="flex justify-between items-start gap-2">
                                <div>
                                  <div className="flex items-center gap-2 flex-wrap">
                                    <h4 className="font-extrabold text-slate-900 text-sm tracking-tight">{student.fullName}</h4>
                                    <span className="text-[10px] font-mono text-slate-500 bg-white border border-slate-200 px-1.5 py-0.5 rounded font-bold">
                                      ID: {student.id}
                                    </span>
                                  </div>
                                  <div className="mt-1.5 flex flex-col gap-1.5">
                                    <div className="flex items-center gap-1.5">
                                      <span className="text-[10px] text-slate-400 font-medium">Enrolled in:</span>
                                      <button
                                        onClick={() => {
                                          setActiveClassName(classSection.name);
                                          // Scroll view smooth to main ribbon
                                          const ribbon = document.getElementById('main-ribbon');
                                          if (ribbon) ribbon.scrollIntoView({ behavior: 'smooth' });
                                        }}
                                        className="text-[10.5px] font-bold text-indigo-600 hover:text-indigo-800 bg-indigo-50 hover:bg-indigo-100 px-2 py-0.5 rounded transition-all cursor-pointer inline-flex items-center gap-0.5"
                                        title="Click to jump to this class section"
                                      >
                                        {classSection.name}
                                        <ChevronRight className="h-3 w-3" />
                                      </button>
                                    </div>
                                    
                                    <div className="flex items-center gap-1.5 mt-0.5">
                                      <span className="text-[10px] text-slate-400 font-bold">Highlight:</span>
                                      <div className="flex items-center gap-1 bg-white p-1 rounded-full border border-slate-200/60 w-fit shadow-3xs">
                                        <button
                                          onClick={() => setStudentHighlightColor(student.id, classSection.name, 'none')}
                                          className={`w-3 h-3 rounded-full border border-slate-300 flex items-center justify-center transition-all bg-slate-50 hover:bg-slate-200 cursor-pointer ${
                                            !student.color ? 'ring-2 ring-indigo-550 scale-105' : 'opacity-85'
                                          }`}
                                          title="Clear Highlight"
                                        >
                                          <X className="h-1 w-1 text-slate-500" />
                                        </button>
                                        {HIGHLIGHT_COLORS.map(c => {
                                          const isSelected = student.color === c.id;
                                          return (
                                            <button
                                              key={c.id}
                                              onClick={() => setStudentHighlightColor(student.id, classSection.name, c.id)}
                                              className={`w-3 h-3 rounded-full ${c.dotClass} border border-black/10 transition-all hover:scale-125 cursor-pointer ${
                                                isSelected ? 'ring-2 ring-indigo-400 ring-offset-0 scale-110' : 'opacity-80'
                                              }`}
                                              title={`Highlight ${c.name}`}
                                            />
                                          );
                                        })}
                                      </div>
                                    </div>
                                  </div>
                                </div>
                                <div className="text-right shrink-0">
                                  <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Final Grade</div>
                                  <div className="text-xs font-black text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-md border border-emerald-100 inline-block mt-0.5 shadow-3xs">
                                    {finalGrade.toFixed(2)}
                                  </div>
                                </div>
                              </div>

                              <div className="border-t border-slate-100 pt-3">
                                <span className="text-[9px] font-extrabold text-slate-400 uppercase tracking-widest block mb-2">
                                  Exam & Assessment Scores
                                </span>

                                {assessments.length === 0 ? (
                                  <p className="text-[10.5px] text-slate-400 italic bg-white p-2.5 rounded border border-slate-100 text-center">
                                    No grading metrics defined for {classSection.name}.
                                  </p>
                                ) : (
                                  <div className="grid grid-cols-2 gap-2">
                                    {assessments.map(ass => {
                                      const score = classSection.grades[student.id]?.[ass.name] ?? 0;
                                      return (
                                        <div key={ass.name} className="flex flex-col gap-1 bg-white p-2 rounded-lg border border-slate-200/80 shadow-3xs">
                                          <div className="text-[10px] font-extrabold text-slate-500 truncate" title={ass.name}>
                                            {ass.name}
                                          </div>
                                          <div className="flex items-center gap-1">
                                            <input
                                              type="number"
                                              min={0}
                                              value={score}
                                              onChange={e => handleScoreChangeCrossSection(student.id, ass.name, e.target.value, classSection.name)}
                                              className="w-full text-xs font-black text-slate-800 bg-slate-50 hover:bg-slate-100 focus:bg-white transition-all outline-none py-0.5 px-1.5 rounded border border-transparent focus:border-indigo-400"
                                            />
                                            <span className="text-[9px] text-slate-400 font-mono select-none">pts</span>
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* ST-STYLE SELECTABLE TABS */}
              <div className="flex border-b border-slate-200 shrink-0 gap-1 overflow-x-auto bg-white p-1 rounded-xl shadow-xs" id="tabs-ribbon">
                <button
                  onClick={() => setActiveTab('students')}
                  className={`flex items-center gap-2 px-4 py-3 text-xs md:text-sm font-bold rounded-lg border transition-all ${
                    activeTab === 'students'
                    ? 'bg-slate-800 text-white border-slate-800 shadow-sm'
                    : 'bg-transparent text-slate-600 border-transparent hover:bg-slate-50'
                  }`}
                  id="tab-btn-students"
                >
                  <Users className="h-4 w-4" />
                  Students List
                </button>

                <button
                  onClick={() => setActiveTab('attendance')}
                  className={`flex items-center gap-2 px-4 py-3 text-xs md:text-sm font-bold rounded-lg border transition-all ${
                    activeTab === 'attendance'
                    ? 'bg-slate-800 text-white border-slate-800 shadow-sm'
                    : 'bg-transparent text-slate-600 border-transparent hover:bg-slate-50'
                  }`}
                  id="tab-btn-attendance"
                >
                  <CalendarCheck className="h-4 w-4" />
                  Attendance Registry
                </button>

                <button
                  onClick={() => setActiveTab('grades')}
                  className={`flex items-center gap-2 px-4 py-3 text-xs md:text-sm font-bold rounded-lg border transition-all ${
                    activeTab === 'grades'
                    ? 'bg-slate-800 text-white border-slate-800 shadow-sm'
                    : 'bg-transparent text-slate-600 border-transparent hover:bg-slate-50'
                  }`}
                  id="tab-btn-grades"
                >
                  <GraduationCap className="h-4 w-4" />
                  Academic Grades
                </button>
              </div>

              {/* ACTIVE WORKING VIEWPORTS */}
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm" id="tab-viewport">
                
                {/* STUDENTS VIEWPORT */}
                {activeTab === 'students' && (
                  <div className="flex flex-col gap-4">
                    <div className="flex justify-between items-center pb-2 border-b border-slate-100">
                      <h3 className="font-extrabold text-slate-900 text-md tracking-tight">
                        Class Roster Listing: {activeClass.name}
                      </h3>
                      <span className="text-xs bg-slate-100 px-2.5 py-1 text-slate-600 font-bold rounded-md">
                        Total Count: {activeClass.students.length} students
                      </span>
                    </div>

                    {activeClass.students.length === 0 ? (
                      <div className="py-12 flex flex-col items-center justify-center text-center">
                        <Users className="h-12 w-12 text-slate-300 mb-3" />
                        <h4 className="font-bold text-slate-800 text-sm">Roster list is currently empty</h4>
                        <p className="text-xs text-slate-400 mt-1 max-w-sm leading-normal">
                          Insert single records using the custom forms, click 'Fill Quick Mock', or upload spreadsheets directly inside the sidebar tools.
                        </p>
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse" id="students-table-matrix">
                          <thead>
                            <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-widest border-b border-slate-100">
                              <th className="px-6 py-3">ID Number</th>
                              <th className="px-6 py-3">Full Name</th>
                              <th className="px-6 py-3 text-right">Actions</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || '';
                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-slate-50/50'} transition-all`}>
                                  <td className={`px-6 py-3 text-slate-500 font-mono font-semibold ${rowBgClass} ${colorInfo.border || ''}`}>
                                    <div className="flex items-center gap-3">
                                      <div className="inline-flex items-center gap-1.5 bg-white px-1.5 py-0.5 rounded-full shadow-3xs border border-slate-200 shrink-0">
                                        {/* Clear / No Color selection */}
                                        <button
                                          onClick={() => setStudentHighlightColor(s.id, activeClass.name, 'none')}
                                          className={`w-3 h-3 rounded-full border border-slate-200 flex items-center justify-center transition-all bg-slate-50 hover:bg-slate-200 cursor-pointer ${
                                            !s.color ? 'ring-1.5 ring-indigo-500 scale-105' : 'opacity-60 hover:opacity-100'
                                          }`}
                                          title="Clear Highlight"
                                        >
                                          <X className="h-1.5 w-1.5 text-slate-400" />
                                        </button>
                                        {HIGHLIGHT_COLORS.map(c => {
                                          const isSelected = s.color === c.id;
                                          return (
                                            <button
                                              key={c.id}
                                              onClick={() => setStudentHighlightColor(s.id, activeClass.name, c.id)}
                                              className={`w-3 h-3 rounded-full ${c.dotClass} border border-black/10 transition-all hover:scale-125 cursor-pointer ${
                                                isSelected ? 'ring-1.5 ring-indigo-500 scale-110' : 'opacity-85 hover:opacity-100'
                                              }`}
                                              title={`Highlight ${c.name}`}
                                            />
                                          );
                                        })}
                                      </div>
                                      <span>{s.id}</span>
                                    </div>
                                  </td>
                                  <td className={`px-6 py-3 font-semibold text-slate-800 ${rowBgClass}`}>{s.fullName}</td>
                                  <td className={`px-6 py-3 text-right ${rowBgClass}`}>
                                    <button
                                      onClick={() => {
                                        setClasses(prev => prev.map(cl => {
                                          if (cl.name === activeClass.name) {
                                            return {
                                              ...cl,
                                              students: cl.students.filter(x => x.id !== s.id)
                                            };
                                          }
                                          return cl;
                                        }));
                                      }}
                                      className="text-rose-500 hover:text-rose-700 hover:underline inline-flex items-center gap-1 font-bold"
                                    >
                                      <Trash2 className="h-3 w-3" /> Dismiss
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* ATTENDANCE VIEWPORT */}
                {activeTab === 'attendance' && (
                  <div className="flex flex-col gap-5">
                    <div className="pb-3 border-b border-slate-100">
                      <h3 className="font-extrabold text-slate-900 text-md tracking-tight">Daily Attendance Registry</h3>
                      <p className="text-xs text-slate-400 mt-1">
                        Define dates and track student presence. Column values represent custom styled markers matching status colors.
                      </p>
                    </div>

                    {/* Choose and configure date columns */}
                    <div className="p-4 bg-slate-50 rounded-xl border border-slate-100 grid md:grid-cols-3 gap-4 items-center">
                      <div className="flex flex-col gap-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Attendance date</label>
                        <input
                          type="date"
                          value={attendanceDate}
                          onChange={e => setAttendanceDate(e.target.value)}
                          className="text-xs bg-white border border-slate-300 rounded-lg p-2 focus:ring-1 outline-none"
                        />
                      </div>
                      <div className="md:col-span-2 pt-4 md:pt-0">
                        <button
                          onClick={handleAddAttendanceColumn}
                          className="bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs px-4 py-2 rounded-lg shadow-sm shrink-0"
                          id="take-attendance-trigger-btn"
                        >
                          Take Attendance (Add Column)
                        </button>
                      </div>
                    </div>

                    {takeAttendanceMsg && (
                      <div className={`text-xs p-3 rounded-lg border flex items-center gap-2 ${
                        takeAttendanceMsg.type === 'success' 
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-100'
                        : 'bg-amber-50 text-amber-900 border-amber-100'
                      }`}>
                        <Check className="h-4 w-4 shrink-0 text-emerald-600" />
                        <span>{takeAttendanceMsg.text}</span>
                      </div>
                    )}

                    {activeClass.students.length === 0 ? (
                      <div className="py-12 text-center text-slate-400 italic">
                        No registered student lists detected. Setup student information first.
                      </div>
                    ) : getAttendanceDates(activeClass).length === 0 ? (
                      <div className="p-8 text-center bg-slate-50 border border-dashed rounded-lg text-slate-400">
                        No attendance date logs found. Pick a test date above and click 'Take Attendance' to initiate your registry!
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse" id="attendance-table-matrix">
                          <thead>
                            <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-widest border-b border-slate-100">
                              <th className="px-5 py-3.5 whitespace-nowrap">ID Number</th>
                              <th className="px-5 py-3.5 whitespace-nowrap">Full Name</th>
                              {getAttendanceDates(activeClass).map(d => (
                                <th key={d} className="px-5 py-3.5 whitespace-nowrap text-center">
                                  {d}
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || '';
                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-slate-50/50'} transition-all`}>
                                  <td className={`px-5 py-3.5 text-slate-500 font-mono font-medium ${rowBgClass} ${colorInfo.border || ''}`}>{s.id}</td>
                                  <td className={`px-5 py-3.5 font-bold text-slate-800 whitespace-nowrap ${rowBgClass}`}>{s.fullName}</td>
                                  {getAttendanceDates(activeClass).map(d => {
                                    const status = activeClass.attendance[s.id]?.[d] || 'Not Taken';
                                    
                                    let pillClass = 'bg-slate-100 text-slate-500 border-slate-200';
                                    if (status === 'Attended') pillClass = 'bg-emerald-50 text-emerald-700 border-emerald-200';
                                    else if (status === 'Absent') pillClass = 'bg-rose-50 text-rose-700 border-rose-200';

                                    return (
                                      <td key={d} className={`px-5 py-3.5 text-center whitespace-nowrap ${rowBgClass}`}>
                                        <button
                                          onClick={() => handleToggleAttendanceCell(s.id, d, status)}
                                          className={`inline-flex items-center px-3 py-1 rounded-full text-[10.5px] font-bold border transition-colors cursor-pointer ${pillClass}`}
                                        >
                                          {status}
                                        </button>
                                      </td>
                                    );
                                  })}
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* ACADEMIC GRADES VIEWPORT */}
                {activeTab === 'grades' && (
                  <div className="flex flex-col gap-6">
                    <div className="pb-3 border-b border-slate-100">
                      <h3 className="font-extrabold text-slate-900 text-md tracking-tight">Academic Grading Sheets</h3>
                      <p className="text-xs text-slate-400 mt-1">
                        Register dynamic assessments and record scores. Custom Excel exporter highlights whitelists and Final Grade calculations automatically.
                      </p>
                    </div>

                    {/* Define an assessment form */}
                    <form onSubmit={handleAddAssessment} className="p-4 bg-slate-50 rounded-xl border border-slate-100 flex flex-col md:flex-row gap-4 items-end">
                      <div className="flex-1 flex flex-col gap-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Exam / Assessment Name</label>
                        <input
                          type="text"
                          value={assessNameInput}
                          onChange={e => setAssessNameInput(e.target.value)}
                          placeholder="e.g. Midterm, Project 1"
                          className="text-xs bg-white border border-slate-300 rounded-lg p-2 focus:ring-1 outline-none text-slate-800 placeholder-slate-400"
                        />
                      </div>

                      <div className="w-full md:w-36 flex flex-col gap-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Max Mark Target</label>
                        <input
                          type="number"
                          min={0}
                          value={assessWeightInput}
                          onChange={e => setAssessWeightInput(parseInt(e.target.value) || 0)}
                          className="text-xs bg-white border border-slate-300 rounded-lg p-2 focus:ring-1 outline-none text-slate-800"
                        />
                      </div>

                      <button
                        type="submit"
                        className="bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs px-4 py-2.5 rounded-lg shadow-sm shrink-0 whitespace-nowrap cursor-pointer"
                        id="register-assessment-trigger-btn"
                      >
                        Register Assessment
                      </button>
                    </form>

                    {assessFeedback && (
                      <div className={`text-xs p-3 rounded-lg border ${
                        assessFeedback.type === 'success' 
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-100'
                        : 'bg-rose-50 text-rose-700 border-rose-10 border-rose-100'
                      }`}>
                        {assessFeedback.text}
                      </div>
                    )}

                    {activeClass.students.length === 0 ? (
                      <div className="py-12 text-center text-slate-400 italic">
                        No active roster exists. Choose Fill Quick Mock or define students in the sidebar to begin.
                      </div>
                    ) : activeClass.assessments.length === 0 ? (
                      <div className="p-8 text-center bg-slate-50 border border-dashed rounded-lg text-slate-440">
                        No grading parameters set up. Create a test assessment metric above (e.g. "Homework 1") to build columns.
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse" id="grades-table-matrix">
                          <thead>
                            <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-widest border-b border-slate-100">
                              <th className="px-5 py-3.5 whitespace-nowrap">ID Number</th>
                              <th className="px-5 py-3.5 whitespace-nowrap">Full Name</th>
                              {activeClass.assessments.map(a => (
                                <th key={a.name} className="px-5 py-3.5 whitespace-nowrap">
                                  {a.name}
                                </th>
                              ))}
                              <th className="px-5 py-3.5 whitespace-nowrap text-right text-emerald-700 font-extrabold bg-emerald-50/50">
                                📊 Final Grade
                              </th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || '';
                              const finalGradeCellClass = s.color ? rowBgClass : 'bg-amber-50/30';
                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-slate-50/50'} transition-all`}>
                                  <td className={`px-5 py-3.5 text-slate-500 font-mono font-medium ${rowBgClass} ${colorInfo.border || ''}`}>{s.id}</td>
                                  <td className={`px-5 py-3.5 font-bold text-slate-800 whitespace-nowrap ${rowBgClass}`}>{s.fullName}</td>
                                  {activeClass.assessments.map(a => {
                                    let score = activeClass.grades[s.id]?.[a.name] ?? 0;
                                    return (
                                      <td key={a.name} className={`px-5 py-3.5 whitespace-nowrap ${rowBgClass}`}>
                                        <input 
                                          type="number"
                                          min={0}
                                          max={120}
                                          value={score}
                                          onChange={e => handleScoreChange(s.id, a.name, e.target.value)}
                                          className={`w-16 text-xs ${s.color ? 'bg-white/80' : 'bg-slate-100'} hover:bg-slate-200/60 focus:bg-white rounded border border-transparent hover:border-slate-300 p-1 font-bold text-slate-800 animate-none`}
                                        />
                                      </td>
                                    );
                                  })}
                                  <td className={`px-5 py-3.5 text-right font-black text-amber-900 whitespace-nowrap ${finalGradeCellClass}`}>
                                    {calculateFinalGrade(s.id, activeClass).toFixed(2)}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

              </div>
            </div>
          )}
        </main>
      </div>

      {/* footer credits */}
      <footer className="bg-slate-100 px-6 py-3 border-t border-slate-200 text-center text-[10px] text-slate-500 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2" id="streamlit-footer">
        <span>Offline-Ready Class Manager whitelists digits limits in custom .env templates.</span>
        <span className="font-mono text-slate-400">Streamlit Sandbox Applet Context</span>
      </footer>
    </div>
  );
}
