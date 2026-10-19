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
  Lock,
  Settings
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { Student, Assessment, ClassSection } from './types';

// Import Firebase config & handlers
import { auth, db, googleProvider, handleFirestoreError, OperationType } from './firebase';
import { signInWithPopup, signOut, onAuthStateChanged, User, signInWithEmailAndPassword, createUserWithEmailAndPassword } from 'firebase/auth';
import { doc, setDoc, onSnapshot } from 'firebase/firestore';

// Initial classes set
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
    attendanceDates: ['2026-05-24'],
    attendance: {
      '101': { '2026-05-24': 'Attended' },
      '102': { '2026-05-24': 'Absent' },
      '103': { '2026-05-24': 'Not Taken' },
      '104': { '2026-05-24': 'Attended' },
      '105': { '2026-05-24': 'Attended' },
    },
    assessments: [
      { name: 'Quiz 1', maxMark: 100, weight: 30 },
      { name: 'Midterm', maxMark: 100, weight: 70 }
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
    attendanceDates: [],
    attendance: {},
    assessments: [],
    grades: {}
  }
];

// Date formatter: converts date format from left to right (D/M/YYYY, e.g. 14/5/2026)
export const formatDisplayDate = (d: string): string => {
  if (!d) return '';
  // Check YYYY-MM-DD
  const ymdMatch = d.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (ymdMatch) {
    const year = ymdMatch[1];
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    return `${day}/${month}/${year}`;
  }
  // Check YYYY/MM/DD
  const ymdSlash = d.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (ymdSlash) {
    const year = ymdSlash[1];
    const month = parseInt(ymdSlash[2], 10);
    const day = parseInt(ymdSlash[3], 10);
    return `${day}/${month}/${year}`;
  }
  return d;
};

const HIGHLIGHT_COLORS = [
  { id: 'emerald', name: 'Green', bgClass: 'bg-emerald-50', hoverClass: 'hover:bg-emerald-100/60', borderClass: 'border-l-emerald-500', dotClass: 'bg-emerald-500' },
  { id: 'sky', name: 'Blue', bgClass: 'bg-sky-50', hoverClass: 'hover:bg-sky-100/60', borderClass: 'border-l-sky-500', dotClass: 'bg-sky-500' },
  { id: 'orange', name: 'Orange', bgClass: 'bg-orange-50', hoverClass: 'hover:bg-orange-100/60', borderClass: 'border-l-orange-500', dotClass: 'bg-orange-500' },
  { id: 'rose', name: 'Red', bgClass: 'bg-rose-50', hoverClass: 'hover:bg-rose-100/60', borderClass: 'border-l-rose-500', dotClass: 'bg-rose-500' },
  { id: 'pink', name: 'Pink', bgClass: 'bg-pink-50', hoverClass: 'hover:bg-pink-100/60', borderClass: 'border-l-pink-500', dotClass: 'bg-pink-500' }
];

const getStudentHighlightClasses = (colorId?: string) => {
  if (!colorId || colorId === 'none') return { bg: '', hover: 'hover:bg-slate-50/50', border: '', text: 'text-slate-800' };
  const found = HIGHLIGHT_COLORS.find(c => c.id === colorId);
  if (!found) return { bg: '', hover: 'hover:bg-slate-50/50', border: '', text: 'text-slate-800' };
  return {
    bg: found.bgClass,
    hover: found.hoverClass,
    border: `border-l-4 ${found.borderClass}`,
    text: 'text-slate-800'
  };
};

function StudentGearBox({
  student,
  className,
  isOpen,
  onToggle,
  onClose,
  onSelectColor,
  onDeleteRequest
}: {
  student: Student;
  className: string;
  isOpen: boolean;
  onToggle: () => void;
  onClose: () => void;
  onSelectColor: (colorId: string) => void;
  onDeleteRequest: () => void;
}) {
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isOpen) return;
    const handleOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose();
      }
    };
    document.addEventListener('mousedown', handleOutside);
    return () => document.removeEventListener('mousedown', handleOutside);
  }, [isOpen, onClose]);

  return (
    <div className="relative inline-flex items-center" ref={menuRef}>
      <button
        type="button"
        onClick={(e) => {
          e.stopPropagation();
          onToggle();
        }}
        className={`p-1.5 rounded-lg border transition-all cursor-pointer ${
          isOpen
            ? 'bg-indigo-100 text-indigo-700 border-indigo-300 ring-2 ring-indigo-400/30'
            : 'bg-white hover:bg-slate-100 text-slate-500 hover:text-slate-800 border-slate-200 shadow-3xs'
        }`}
        title={`Settings for ${student.fullName}`}
        aria-label={`Settings for ${student.fullName}`}
      >
        <Settings className="h-3.5 w-3.5" />
      </button>

      {isOpen && (
        <div 
          onClick={(e) => e.stopPropagation()}
          className="absolute left-0 top-full mt-1.5 z-50 w-56 bg-white rounded-2xl shadow-xl border border-slate-200 p-3 flex flex-col gap-2.5 animate-in fade-in zoom-in-95 duration-150 text-left"
        >
          {/* Header */}
          <div className="flex items-center justify-between border-b border-slate-100 pb-2">
            <div className="min-w-0 pr-1">
              <span className="text-xs font-bold text-slate-800 truncate block">{student.fullName}</span>
              <span className="text-[10px] text-slate-400 font-mono">ID: {student.id}</span>
            </div>
            <button
              type="button"
              onClick={onClose}
              className="p-1 text-slate-400 hover:text-slate-600 rounded-lg hover:bg-slate-100 cursor-pointer"
            >
              <X className="h-3 w-3" />
            </button>
          </div>

          {/* 1. HIGHLIGHT COLOR (ABOVE DELETE OPTION) */}
          <div className="flex flex-col gap-1.5">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
              Highlight Color:
            </span>
            <div className="grid grid-cols-2 gap-1.5">
              <button
                type="button"
                onClick={() => {
                  onSelectColor('none');
                }}
                className={`flex items-center gap-1.5 px-2 py-1 rounded-lg text-[10px] font-bold border transition-all cursor-pointer ${
                  !student.color || student.color === 'none'
                    ? 'bg-slate-200 text-slate-800 border-slate-300 ring-1 ring-slate-400 font-extrabold'
                    : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                }`}
              >
                <X className="h-3 w-3 text-slate-500" />
                <span>None</span>
              </button>

              {HIGHLIGHT_COLORS.map(c => {
                const isSelected = student.color === c.id;
                return (
                  <button
                    key={c.id}
                    type="button"
                    onClick={() => {
                      onSelectColor(c.id);
                    }}
                    className={`flex items-center gap-1.5 px-2 py-1 rounded-lg text-[10px] font-bold border transition-all cursor-pointer ${
                      isSelected
                        ? `${c.bgClass} text-slate-900 border-slate-400 ring-2 ring-indigo-500 font-black shadow-3xs`
                        : `${c.bgClass} text-slate-700 border-slate-200 hover:brightness-95`
                    }`}
                  >
                    <span className={`w-2.5 h-2.5 rounded-full ${c.dotClass} shrink-0 border border-black/10`} />
                    <span>{c.name}</span>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="h-px bg-slate-100" />

          {/* 2. DELETE OPTION (BELOW HIGHLIGHT OPTION) */}
          <div className="flex flex-col gap-1">
            <button
              type="button"
              onClick={onDeleteRequest}
              className="w-full flex items-center justify-center gap-1.5 px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 hover:text-rose-900 border border-rose-200 rounded-xl text-[11px] font-bold transition-all cursor-pointer active:scale-95"
            >
              <Trash2 className="h-3.5 w-3.5 text-rose-600" />
              <span>Delete Student</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  // Academic classes & sections state
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
  const [activeTab, setActiveTab] = useState<'students' | 'attendance' | 'grades'>('students');

  // Sidebar Controls
  const [newClassNameInput, setNewClassNameInput] = useState('');
  const [isCreatingClass, setIsCreatingClass] = useState(false);
  const [confirmDeleteClass, setConfirmDeleteClass] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [expandedSearchStudentId, setExpandedSearchStudentId] = useState<string | null>(null);

  // Student Gear Box & Warning States
  const [activeGearStudent, setActiveGearStudent] = useState<{ studentId: string; className: string } | null>(null);
  const [confirmDeleteStudentTarget, setConfirmDeleteStudentTarget] = useState<{ studentId: string; studentName: string; className: string } | null>(null);

  // Attendance state
  const [attendanceDate, setAttendanceDate] = useState<string>('2026-05-25');
  const [takeAttendanceMsg, setTakeAttendanceMsgMessage] = useState<{type: 'success'|'warning', text: string} | null>(null);

  // Grade Assessment Modal (Add & Edit) states
  const [isAssessmentModalOpen, setIsAssessmentModalOpen] = useState(false);
  const [assessmentModalMode, setAssessmentModalMode] = useState<'add' | 'edit'>('add');
  const [editingAssessmentOriginalName, setEditingAssessmentOriginalName] = useState<string | null>(null);
  const [assessNameInput, setAssessNameInput] = useState('');
  const [assessMaxMarkInput, setAssessMaxMarkInput] = useState<number>(100);
  const [assessWeightInput, setAssessWeightInput] = useState<number>(20);
  const [assessModalError, setAssessModalError] = useState<string | null>(null);

  // Final Grade Breakdown & Weights Configuration Modal
  const [isFinalGradeModalOpen, setIsFinalGradeModalOpen] = useState(false);
  const [tempWeights, setTempWeights] = useState<{ [assessName: string]: number }>({});
  const [finalGradeModalError, setFinalGradeModalError] = useState<string | null>(null);

  // Attendance Date Deletion state
  const [confirmDeleteDateTarget, setConfirmDeleteDateTarget] = useState<{ date: string; displayDate: string } | null>(null);

  // Feedback notifications
  const [assessFeedback, setAssessFeedback] = useState<{type: 'error' | 'success', text: string} | null>(null);

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

  // States for importing student roster ONLY during new class creation modal
  const [createdClassStudents, setCreatedClassStudents] = useState<Student[]>([]);
  const [createdClassFileError, setCreatedClassFileError] = useState<string | null>(null);
  const [createdClassFileSuccess, setCreatedClassFileSuccess] = useState<string | null>(null);
  const [isDragOver, setIsDragOver] = useState(false);

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

  const totalAllocatedWeight = activeClass.assessments.reduce((sum, a) => sum + (Number(a.weight) || 0), 0);

  // Helper metrics for final grade calculating: (rawScore / maxMark) * weight
  const calculateFinalGrade = (studentId: string, currentClass: ClassSection) => {
    if (!currentClass.assessments || currentClass.assessments.length === 0) return 0;

    let totalWeightedScore = 0;
    currentClass.assessments.forEach(ass => {
      const rawScore = currentClass.grades[studentId]?.[ass.name];
      if (rawScore !== undefined && rawScore !== null && !isNaN(rawScore)) {
        const maxMark = ass.maxMark && ass.maxMark > 0 ? ass.maxMark : 100;
        const weight = ass.weight || 0;
        const contribution = (rawScore / maxMark) * weight;
        totalWeightedScore += contribution;
      }
    });

    return Math.round(totalWeightedScore * 100) / 100;
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

  // Creation of a New Class Section database row (including optional roster import)
  const handleCreateClass = (e: React.FormEvent) => {
    e.preventDefault();
    const formattedName = newClassNameInput.trim();
    if (!formattedName) return;

    if (classes.some(c => c.name.toLowerCase() === formattedName.toLowerCase())) {
      alert(`Class '${formattedName}' already exists.`);
      return;
    }

    const initialAttendance: any = {};
    const initialGrades: any = {};
    createdClassStudents.forEach(st => {
      initialAttendance[st.id] = {};
      initialGrades[st.id] = {};
    });

    const newClass: ClassSection = {
      name: formattedName,
      students: createdClassStudents,
      attendance: initialAttendance,
      assessments: [],
      grades: initialGrades
    };

    setClasses(prev => [...prev, newClass]);
    setActiveClassName(formattedName);
    
    // Reset states
    setNewClassNameInput('');
    setCreatedClassStudents([]);
    setCreatedClassFileError(null);
    setCreatedClassFileSuccess(null);
    setIsCreatingClass(false);
  };

  // Parsing helper for file upload inside class creation modal
  const handleModalRosterFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCreatedClassFileError(null);
    setCreatedClassFileSuccess(null);
    const file = e.target.files?.[0];
    if (!file) return;

    processModalRosterFile(file);
  };

  const processModalRosterFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        let jsonRows: any[] = [];

        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheet = workbook.SheetNames[0];
        jsonRows = XLSX.utils.sheet_to_json<any>(workbook.Sheets[firstSheet]);

        if (!jsonRows || jsonRows.length === 0) {
          setCreatedClassFileError("The uploaded file contains zero records or is corrupted.");
          return;
        }

        const normalizedRows = jsonRows.map(row => {
          const keys = Object.keys(row);
          const normalized: any = {};
          keys.forEach(k => {
            normalized[k.trim().toLowerCase().replace(/\s+/g, '')] = row[k];
          });
          return { original: row, norm: normalized };
        });

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
          setCreatedClassFileError(`Columns matching 'ID Number', 'First Name', and 'Last Name' were not located.`);
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
          setCreatedClassFileError("Failed to parse valid student records containing names and ID numbers.");
          return;
        }

        const sortedStudents = parsedStudents.sort((a, b) => 
          a.firstName.localeCompare(b.firstName)
        );

        setCreatedClassStudents(sortedStudents);
        setCreatedClassFileSuccess(`Successfully recognized ${sortedStudents.length} student records from file.`);
      } catch (err: any) {
        setCreatedClassFileError(`Error parsing file: ${err?.message || 'Unsupported format'}`);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleModalRosterDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleModalRosterDragLeave = () => {
    setIsDragOver(false);
  };

  const handleModalRosterDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      processModalRosterFile(file);
    }
  };

  // Deletion of the active Class Section (works even if only 1 left)
  const handleDeleteActiveSection = () => {
    const remainingClasses = classes.filter(c => c.name !== activeClass.name);
    setClasses(remainingClasses);
    setActiveClassName(remainingClasses.length > 0 ? remainingClasses[0].name : '');
    setConfirmDeleteClass(false);
  };

  // Deletion of a specific Student from a class
  const handleDeleteStudentFromClass = (studentId: string, className: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === className) {
        const nextAtt = { ...cl.attendance };
        delete nextAtt[studentId];

        const nextGrades = { ...cl.grades };
        delete nextGrades[studentId];

        return {
          ...cl,
          students: cl.students.filter(s => s.id !== studentId),
          attendance: nextAtt,
          grades: nextGrades
        };
      }
      return cl;
    }));
    setConfirmDeleteStudentTarget(null);
    setActiveGearStudent(null);
  };

  // Add Attendance Date Column (new column appears right next to student names)
  const handleAddAttendanceColumn = () => {
    setTakeAttendanceMsgMessage(null);
    const dateStr = attendanceDate;
    if (!dateStr) return;

    // Check if the current section already has this column
    const dates = getAttendanceDates(activeClass);
    if (dates.includes(dateStr)) {
      setTakeAttendanceMsgMessage({
        type: 'warning',
        text: `The attendance column '${formatDisplayDate(dateStr)}' already exists in the table.`
      });
      return;
    }

    // Prepend new date so it appears immediately next to names, shifting older columns to the right
    const nextDates = [dateStr, ...dates];

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextAtt = { ...cl.attendance };
        cl.students.forEach(st => {
          if (!nextAtt[st.id]) nextAtt[st.id] = {};
          nextAtt[st.id][dateStr] = 'Not Taken';
        });
        return { 
          ...cl, 
          attendanceDates: nextDates,
          attendance: nextAtt 
        };
      }
      return cl;
    }));

    setTakeAttendanceMsgMessage({
      type: 'success',
      text: `Added attendance date column '${formatDisplayDate(dateStr)}' next to student names.`
    });
  };

  // Delete a specific Attendance Date Column
  const handleDeleteAttendanceDate = (dateToDelete: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextDates = getAttendanceDates(cl).filter(d => d !== dateToDelete);
        const nextAtt = { ...cl.attendance };
        cl.students.forEach(st => {
          if (nextAtt[st.id]) {
            const studentCopy = { ...nextAtt[st.id] };
            delete studentCopy[dateToDelete];
            nextAtt[st.id] = studentCopy;
          }
        });
        return {
          ...cl,
          attendanceDates: nextDates,
          attendance: nextAtt
        };
      }
      return cl;
    }));
    setConfirmDeleteDateTarget(null);
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

  // ASSESSMENT MODAL HANDLERS (Add, Edit, Reopen from header)
  const handleOpenAddAssessment = () => {
    setAssessmentModalMode('add');
    setEditingAssessmentOriginalName(null);
    setAssessNameInput('');
    setAssessMaxMarkInput(100);
    const currentTotal = activeClass.assessments.reduce((sum, a) => sum + (Number(a.weight) || 0), 0);
    const remaining = Math.max(0, 100 - currentTotal);
    setAssessWeightInput(remaining > 0 ? remaining : 0);
    setAssessModalError(null);
    setIsAssessmentModalOpen(true);
  };

  const handleOpenEditAssessment = (assessment: Assessment) => {
    setAssessmentModalMode('edit');
    setEditingAssessmentOriginalName(assessment.name);
    setAssessNameInput(assessment.name);
    setAssessMaxMarkInput(assessment.maxMark || 100);
    setAssessWeightInput(assessment.weight || 0);
    setAssessModalError(null);
    setIsAssessmentModalOpen(true);
  };

  const handleSaveAssessmentModal = (e: React.FormEvent) => {
    e.preventDefault();
    setAssessModalError(null);
    const formName = assessNameInput.trim();
    if (!formName) {
      setAssessModalError('Please enter an assessment name.');
      return;
    }

    const maxMark = assessMaxMarkInput > 0 ? assessMaxMarkInput : 100;
    const weight = Math.max(0, assessWeightInput || 0);

    // Sum of other weights
    const otherWeights = activeClass.assessments
      .filter(a => assessmentModalMode === 'edit' ? a.name !== editingAssessmentOriginalName : true)
      .reduce((sum, a) => sum + (Number(a.weight) || 0), 0);

    if (otherWeights + weight > 100) {
      const allowed = Math.max(0, 100 - otherWeights);
      setAssessModalError(`Total assessment weights cannot exceed 100%. Other assessments use ${otherWeights}%, so maximum weight for this test is ${allowed}%.`);
      return;
    }

    if (assessmentModalMode === 'add') {
      if (activeClass.assessments.some(a => a.name.toLowerCase() === formName.toLowerCase())) {
        setAssessModalError(`An assessment named '${formName}' is already registered.`);
        return;
      }

      const newAssessment: Assessment = {
        name: formName,
        maxMark,
        weight
      };

      // When adding new assessment, prepend it so it appears right next to names!
      setClasses(prev => prev.map(cl => {
        if (cl.name === activeClass.name) {
          const nextGrades = { ...cl.grades };
          cl.students.forEach(st => {
            if (!nextGrades[st.id]) nextGrades[st.id] = {};
            nextGrades[st.id][formName] = 0;
          });

          return {
            ...cl,
            assessments: [newAssessment, ...cl.assessments],
            grades: nextGrades
          };
        }
        return cl;
      }));

      setIsAssessmentModalOpen(false);
      setAssessFeedback({
        type: 'success',
        text: `Added '${formName}' (Full mark: ${maxMark}, Weight: ${weight}%) next to student names.`
      });
    } else {
      // Edit mode
      const oldName = editingAssessmentOriginalName;
      if (!oldName) return;

      if (formName.toLowerCase() !== oldName.toLowerCase() &&
          activeClass.assessments.some(a => a.name.toLowerCase() === formName.toLowerCase())) {
        setAssessModalError(`An assessment named '${formName}' is already registered.`);
        return;
      }

      setClasses(prev => prev.map(cl => {
        if (cl.name === activeClass.name) {
          const nextAssessments = cl.assessments.map(a => {
            if (a.name === oldName) {
              return { name: formName, maxMark, weight };
            }
            return a;
          });

          const nextGrades = { ...cl.grades };
          if (formName !== oldName) {
            cl.students.forEach(st => {
              if (nextGrades[st.id] && nextGrades[st.id][oldName] !== undefined) {
                const prevScore = nextGrades[st.id][oldName];
                delete nextGrades[st.id][oldName];
                nextGrades[st.id][formName] = prevScore;
              }
            });
          }

          return {
            ...cl,
            assessments: nextAssessments,
            grades: nextGrades
          };
        }
        return cl;
      }));

      setIsAssessmentModalOpen(false);
      setAssessFeedback({
        type: 'success',
        text: `Updated '${formName}' parameters (Full mark: ${maxMark}, Weight: ${weight}%).`
      });
    }
  };

  const handleDeleteAssessment = (nameToDelete: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextAssessments = cl.assessments.filter(a => a.name !== nameToDelete);
        const nextGrades = { ...cl.grades };
        cl.students.forEach(st => {
          if (nextGrades[st.id]) {
            delete nextGrades[st.id][nameToDelete];
          }
        });
        return {
          ...cl,
          assessments: nextAssessments,
          grades: nextGrades
        };
      }
      return cl;
    }));
    setIsAssessmentModalOpen(false);
  };

  // FINAL GRADE WEIGHTS POPUP HANDLERS
  const handleOpenFinalGradeWeights = () => {
    const weightsMap: { [name: string]: number } = {};
    activeClass.assessments.forEach(a => {
      weightsMap[a.name] = a.weight || 0;
    });
    setTempWeights(weightsMap);
    setFinalGradeModalError(null);
    setIsFinalGradeModalOpen(true);
  };

  const handleSaveFinalGradeWeights = () => {
    const totalSum = (Object.values(tempWeights) as number[]).reduce((sum: number, w: number) => sum + (Number(w) || 0), 0);
    if (totalSum > 100) {
      setFinalGradeModalError(`Total weight is ${totalSum}%, which exceeds 100%. Total weight cannot exceed 100%.`);
      return;
    }

    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextAssessments = cl.assessments.map(a => ({
          ...a,
          weight: tempWeights[a.name] !== undefined ? Math.max(0, Number(tempWeights[a.name])) : a.weight
        }));
        return {
          ...cl,
          assessments: nextAssessments
        };
      }
      return cl;
    }));

    setIsFinalGradeModalOpen(false);
    setAssessFeedback({
      type: 'success',
      text: `Updated assessment weights for '${activeClass.name}'.`
    });
  };

  const handleAutoDistributeWeights = () => {
    if (activeClass.assessments.length === 0) return;
    const count = activeClass.assessments.length;
    const baseWeight = Math.floor(100 / count);
    const remainder = 100 - (baseWeight * count);

    const newTemp: { [name: string]: number } = {};
    activeClass.assessments.forEach((a, idx) => {
      newTemp[a.name] = baseWeight + (idx === 0 ? remainder : 0);
    });
    setTempWeights(newTemp);
    setFinalGradeModalError(null);
  };

  // Edit Single score in table matrix
  const handleScoreChange = (studentId: string, assessmentName: string, scoreStr: string) => {
    let scoreNum = parseFloat(scoreStr);
    if (isNaN(scoreNum)) {
      scoreNum = 0;
    }
    const currentAss = activeClass.assessments.find(a => a.name === assessmentName);
    const maxAllowed = currentAss?.maxMark && currentAss.maxMark > 0 ? currentAss.maxMark : 100;
    if (scoreNum < 0) scoreNum = 0;
    if (scoreNum > maxAllowed) scoreNum = maxAllowed;

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

        // Always set pendingImport so we can preview the full checklist, showcasing short IDs at the top as priority
        setPendingImport({
          className: activeClass.name,
          normal: normalIds,
          short: shortIds
        });
        if (shortIds.length > 0) {
          setFileSuccess(`Verification check required: Found ${shortIds.length} short student IDs. Previewing roster.`);
        } else {
          setFileSuccess(`Previewing imported student roster containing ${sortedStudents.length} records.`);
        }

      } catch (err: any) {
        setFileError(`Error parsing workbook: ${err?.message || 'Unsupported format'}`);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Commit Staging Roster with Checkbox Selection
  const handleConfirmPendingImport = () => {
    if (!pendingImport) return;

    // Filter students based on what is checked
    const approvedShortStudents = pendingImport.short.filter(st => 
      selectedShortIds.includes(st.id)
    );
    const approvedNormalStudents = pendingImport.normal.filter(st => 
      selectedShortIds.includes(st.id)
    );

    const consolidatedList = [...approvedNormalStudents, ...approvedShortStudents].sort((a, b) => 
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
    setFileSuccess(`Successfully imported roster containing ${consolidatedList.length} records.`);
  };

  // Deselects all short IDs from check list instantly
  const handleDeselectShortIDsChecklist = () => {
    if (!pendingImport) return;
    const shortIdsList = pendingImport.short.map(s => s.id);
    setSelectedShortIds(prev => prev.filter(id => !shortIdsList.includes(id)));
  };

  // Selects all short and normal IDs in checklist instantly
  const handleSelectAllChecklist = () => {
    if (!pendingImport) return;
    const allIds = [...pendingImport.short, ...pendingImport.normal].map(s => s.id);
    setSelectedShortIds(allIds);
  };

  const handleDiscardPendingImportShortIds = () => {
    if (!pendingImport) return;

    // Direct save with only normal entries (fully skip short IDs)
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

  // Extract date columns from attendance ledger (preserves custom date ordering)
  const getAttendanceDates = (cl: ClassSection): string[] => {
    if (cl.attendanceDates && Array.isArray(cl.attendanceDates)) {
      return cl.attendanceDates;
    }
    const datesSet = new Set<string>();
    Object.values(cl.attendance || {}).forEach(record => {
      Object.keys(record || {}).forEach(d => datesSet.add(d));
    });
    return Array.from(datesSet).sort((a, b) => b.localeCompare(a));
  };

  // Single Class Exporters
  const handleExportGrades = () => {
    const tableData = activeClass.students.map(s => {
      const row: any = {
        'ID Number': s.id,
        'Full Name': s.fullName
      };
      
      activeClass.assessments.forEach(ass => {
        const colTitle = `${ass.name} (Max ${ass.maxMark || 100}, ${ass.weight}%)`;
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
        row[formatDisplayDate(d)] = activeClass.attendance[s.id]?.[d] ?? 'Not Taken';
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
        const headers = ['ID Number', 'Full Name', ...dates.map(d => formatDisplayDate(d))];
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
    <div className="min-h-screen bg-slate-50 flex flex-col text-slate-800 font-sans antialiased" id="omni-class-app-container">
      {/* Omni Class Brand Accent Strip */}
      <div className="h-1 bg-gradient-to-r from-indigo-500 via-purple-500 to-indigo-600 w-full" id="omni-accent-line"></div>

      {/* Header Banner */}
      <header className="bg-[#f3f4f6]/90 backdrop-blur-md border-b border-slate-200/80 px-6 py-4 flex flex-col md:flex-row md:items-center md:justify-between gap-4" id="omni-header">
        <div className="flex items-center gap-3.5">
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className="p-3 bg-white hover:bg-slate-100 rounded-full text-slate-700 transition-all cursor-pointer flex items-center justify-center border border-slate-200 shadow-3xs active:scale-95"
            id="toggle-sidebar-hamburger"
            title={isSidebarOpen ? "Minimize Navigation Drawer" : "Expand Navigation Drawer"}
          >
            <Menu className="h-5 w-5" />
          </button>
          
          <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-2xl border border-indigo-100/50 shadow-3xs">
            <GraduationCap className="h-6 w-6" />
          </div>

          <div>
            <h1 className="text-lg md:text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
              Omni Class
            </h1>
            <p className="text-[11px] font-medium text-slate-500 leading-none mt-1">Class Operations Hub</p>
          </div>
        </div>
        
        <div className="flex items-center gap-2.5 flex-wrap">
          {user ? (
            <span className="inline-flex bg-emerald-50 text-emerald-800 font-mono text-[10px] uppercase font-bold items-center px-3 py-1.5 rounded-full border border-emerald-200 gap-1.5 shadow-3xs select-none" id="data-sync-badge">
              <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse"></span>
              Data Sync ON
            </span>
          ) : (
            <span className="inline-flex bg-slate-100 text-slate-500 font-mono text-[10px] uppercase font-bold items-center px-3 py-1.5 rounded-full border border-slate-200 gap-1.5 shadow-3xs select-none" id="data-sync-badge">
              <span className="w-1.5 h-1.5 bg-slate-400 rounded-full"></span>
              Data Sync OFF
            </span>
          )}
        </div>
      </header>

      {/* Interface workspace */}
      <div className="flex-1 flex flex-col lg:flex-row relative" id="omni-workspace">
        
        {/* Mobile slide-out overlay backdrop */}
        {isSidebarOpen && (
          <div 
            onClick={() => setIsSidebarOpen(false)}
            className="fixed inset-0 bg-slate-900/40 z-40 lg:hidden backdrop-blur-3xs transition-opacity duration-300"
            id="sidebar-mobile-backdrop"
          />
        )}

        {/* SIDEBAR WIDGET PANELS */}
        <aside className={`
          fixed inset-y-0 left-0 z-50 w-80 bg-slate-50 p-5 flex flex-col shrink-0 gap-5 border-r border-slate-200 shadow-xl overflow-y-auto transition-transform duration-300
          lg:self-start lg:sticky lg:top-0 lg:max-h-screen lg:translate-x-0 lg:shadow-none lg:w-80 lg:flex lg:border-r lg:border-slate-200
          ${isSidebarOpen ? 'translate-x-0 flex' : '-translate-x-full hidden lg:flex'}
        `} id="omni-sidebar">
          
          {/* Mobile close button drawer header */}
          <div className="flex items-center justify-between lg:hidden border-b border-slate-200 pb-3" id="sidebar-drawer-mobile-header">
            <span className="text-sm font-bold text-slate-800 tracking-tight flex items-center gap-1.5">
              <GraduationCap className="h-5 w-5 text-indigo-600" />
              Navigation Panel
            </span>
            <button
              onClick={() => setIsSidebarOpen(false)}
              className="p-1.5 hover:bg-slate-200 rounded-full text-slate-500 border border-slate-200 transition-colors cursor-pointer"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Firebase Sync Widget */}
          <div className="bg-slate-100 p-4 rounded-3xl border border-slate-200 shadow-3xs flex flex-col gap-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                <Cloud className={`h-4 w-4 ${user ? 'text-indigo-600' : 'text-slate-400'}`} />
                Cloud Database
              </span>
              
              {user && (
                <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold ${
                  isSyncing ? 'bg-amber-100 text-amber-800 animate-pulse' : 'bg-emerald-100 text-emerald-800'
                }`}>
                  <span className={`w-1 h-1 rounded-full mr-1 ${isSyncing ? 'bg-amber-500 animate-ping' : 'bg-emerald-500'}`} />
                  {isSyncing ? "Syncing..." : "Live"}
                </span>
              )}
            </div>

            {authLoading ? (
              <div className="flex items-center gap-2 text-xs text-slate-400 font-medium py-1">
                <RefreshCw className="h-3 w-3 animate-spin text-indigo-500" />
                <span>Checking credentials...</span>
              </div>
            ) : user ? (
              <div className="flex flex-col gap-2">
                <div className="text-[11px] text-slate-600 bg-white p-2.5 rounded-2xl border border-slate-200 flex flex-col gap-1 shadow-3xs">
                  <div className="font-semibold text-slate-800 truncate flex items-center gap-1.5" title={user.email || ""}>
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                    <span className="truncate">{user.displayName || user.email?.split('@')[0] || "Professor"}</span>
                  </div>
                  <div className="text-[10px] text-slate-400 truncate font-mono">{user.email}</div>
                </div>

                <button
                  onClick={handleSignOut}
                  className="w-full flex items-center justify-center gap-1.5 px-3 py-2 bg-rose-50 hover:bg-rose-100 border border-rose-150 hover:border-rose-200 text-rose-700 rounded-full text-xs font-bold transition-all cursor-pointer shadow-3xs"
                >
                  <LogOut className="h-3.5 w-3.5" />
                  Sign Out
                </button>
              </div>
            ) : (
              <div className="flex flex-col gap-2.5">
                <p className="text-[11px] text-slate-500 leading-normal font-medium">
                  Connect cloud database to access your class data from any device securely.
                </p>

                <button
                  type="button"
                  onClick={handleGoogleSignIn}
                  disabled={isSyncing}
                  className="w-full flex items-center justify-center gap-2.5 px-3 py-2.5 bg-white hover:bg-slate-50 active:bg-slate-100 text-slate-700 hover:text-slate-900 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-3xs border border-slate-200 active:scale-[0.98] disabled:opacity-50"
                  id="google-login-btn"
                >
                  {isSyncing ? (
                    <RefreshCw className="h-4 w-4 animate-spin text-indigo-600" />
                  ) : (
                    <svg className="w-4 h-4 shrink-0" viewBox="0 0 24 24">
                      <path
                        fill="#4285F4"
                        d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"
                      />
                      <path
                        fill="#34A853"
                        d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
                      />
                      <path
                        fill="#FBBC05"
                        d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z"
                      />
                      <path
                        fill="#EA4335"
                        d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z"
                      />
                    </svg>
                  )}
                  <span>Login with Google</span>
                </button>
              </div>
            )}

            {syncErrorMessage && (
              <div className="bg-rose-50 text-rose-700 border border-rose-100 p-2.5 rounded-xl text-[10px] font-medium leading-normal flex items-start gap-1.5">
                <AlertCircle className="h-3.5 w-3.5 shrink-0 mt-0.5" />
                <span>{syncErrorMessage}</span>
              </div>
            )}
          </div>

          {/* Active Section Selector */}
          <div className="flex flex-col gap-2" id="sidebar-class-selector">
            <label className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest">
              Select a class
            </label>
            
            <div className="flex gap-2 items-center">
              <select
                value={classes.length === 0 ? '' : activeClassName}
                onChange={e => {
                  setActiveClassName(e.target.value);
                  setConfirmDeleteClass(false);
                }}
                disabled={classes.length === 0}
                className={`flex-1 text-xs rounded-xl p-2.5 outline-none font-semibold shadow-3xs transition-all ${
                  classes.length === 0
                    ? 'bg-slate-200/80 text-slate-400 border border-slate-300 cursor-not-allowed italic'
                    : 'bg-white border border-slate-300 text-slate-800 cursor-pointer focus:ring-2 focus:ring-indigo-500/20'
                }`}
              >
                {classes.length === 0 ? (
                  <option value="" disabled>No Active Sections</option>
                ) : (
                  classes.map(c => (
                    <option key={c.name} value={c.name}>{c.name}</option>
                  ))
                )}
              </select>
              <button
                type="button"
                onClick={() => {
                  setNewClassNameInput('');
                  setCreatedClassStudents([]);
                  setCreatedClassFileError(null);
                  setCreatedClassFileSuccess(null);
                  setIsCreatingClass(true);
                }}
                className="w-10 h-10 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white rounded-xl transition-all flex items-center justify-center cursor-pointer shadow-sm active:scale-95 shrink-0 border border-indigo-700"
                title="Create New Class (with optional import)"
                id="create-new-class-sidebar-btn"
                aria-label="Add new class section"
              >
                <Plus className="h-5 w-5 text-white" strokeWidth={2.5} />
              </button>
            </div>
            <p className="text-[10px] text-slate-400 font-mono pl-1">
              {classes.length === 0 ? 'No section selected' : `Roster size: ${activeClass.students.length} students enrolled`}
            </p>
          </div>

          {/* Delete Active Section Controller */}
          <div className="flex flex-col gap-2 pt-0.5" id="sidebar-delete-class-box">
            {classes.length === 0 ? (
              <div className="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-[11px] font-medium text-slate-400 bg-slate-100 border border-slate-200 rounded-xl select-none" title="No active sections to delete">
                <Trash2 className="h-3.5 w-3.5 text-slate-300" />
                <span>No sections to delete</span>
              </div>
            ) : confirmDeleteClass ? (
              <div className="flex flex-col gap-2.5 p-3 rounded-2xl bg-rose-50 border-2 border-rose-300 text-left w-full shadow-sm">
                <div className="flex items-start gap-2">
                  <Trash2 className="h-4 w-4 text-rose-600 shrink-0 mt-0.5" />
                  <p className="text-xs text-rose-950 leading-snug font-medium">
                    Are you sure you want to delete <strong className="font-bold underline text-rose-900">'{activeClass.name}'</strong>? {classes.length === 1 ? 'This is your only class section. All' : 'All'} student records, grades, and attendance will be permanently removed.
                  </p>
                </div>
                <div className="flex items-center gap-2 justify-end pt-1">
                  <button
                    type="button"
                    onClick={() => setConfirmDeleteClass(false)}
                    className="text-xs bg-white border border-slate-300 rounded-lg text-slate-700 px-3 py-1.5 hover:bg-slate-100 font-semibold transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={handleDeleteActiveSection}
                    className="text-xs bg-rose-600 hover:bg-rose-700 text-white rounded-lg px-3.5 py-1.5 font-bold shadow-sm transition-all cursor-pointer active:scale-95 flex items-center gap-1.5"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    Confirm Delete
                  </button>
                </div>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => setConfirmDeleteClass(true)}
                className="w-full inline-flex items-center justify-center gap-2 px-3 py-2 text-xs font-bold text-rose-700 bg-rose-100/90 hover:bg-rose-200 active:bg-rose-300/80 border border-rose-300 rounded-xl transition-all cursor-pointer shadow-3xs"
                title={`Delete '${activeClass.name}'`}
                id="delete-class-btn"
              >
                <Trash2 className="h-3.5 w-3.5 text-rose-600" strokeWidth={2.5} />
                <span className="truncate">Delete Class ({activeClass.name})</span>
              </button>
            )}
          </div>

          {/* GLOBAL STUDENT SEARCH (Shifted up directly under Class Selection) */}
          <div className="flex flex-col gap-2 bg-slate-100 p-4 rounded-3xl border border-slate-200 shadow-3xs" id="sidebar-global-student-search">
            <div className="flex items-center justify-between pb-0.5">
              <span className="text-[11px] font-black text-slate-700 uppercase tracking-wide flex items-center gap-1.5">
                <Search className="h-3.5 w-3.5 text-indigo-600 font-bold" />
                Global Search
              </span>
              {globalSearchQuery && (
                <button 
                  onClick={() => setGlobalSearchQuery('')}
                  className="text-[10px] text-indigo-600 hover:text-indigo-800 font-bold transition-all cursor-pointer"
                >
                  Clear
                </button>
              )}
            </div>

            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                <Search className="h-3.5 w-3.5" />
              </span>
              <input
                type="text"
                value={globalSearchQuery}
                onChange={e => setGlobalSearchQuery(e.target.value)}
                placeholder="Type Student Name or ID..."
                className="w-full pl-9 pr-3 py-2 text-xs bg-white border border-slate-300 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 text-slate-800 placeholder-slate-400 font-medium transition-all"
              />
            </div>

            {globalSearchQuery.trim() !== '' && (
              <div className="max-h-80 overflow-y-auto mt-1 flex flex-col gap-2.5 divide-y divide-slate-200 border-t border-slate-200/50 pt-2" id="sidebar-search-results">
                {globalSearchResults.length === 0 ? (
                  <div className="py-5 text-center text-slate-400 text-[10.5px]">
                    No students found matching <strong className="text-slate-600">"{globalSearchQuery}"</strong>
                  </div>
                ) : (
                  globalSearchResults.map(({ student, classSection }) => {
                    const assessments = classSection.assessments;
                    const finalGrade = calculateFinalGrade(student.id, classSection);
                    const isExpanded = expandedSearchStudentId === `${classSection.name}-${student.id}`;
                    
                    return (
                      <div key={`${classSection.name}-${student.id}`} className="pt-2.5 first:pt-0 flex flex-col gap-2">
                        
                        <div className="flex justify-between items-start gap-1">
                          <div className="min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <StudentGearBox
                                student={student}
                                className={classSection.name}
                                isOpen={activeGearStudent?.studentId === student.id && activeGearStudent?.className === classSection.name}
                                onToggle={() => {
                                  setActiveGearStudent(prev => 
                                    prev?.studentId === student.id && prev?.className === classSection.name
                                      ? null
                                      : { studentId: student.id, className: classSection.name }
                                  );
                                }}
                                onClose={() => setActiveGearStudent(null)}
                                onSelectColor={(colorId) => setStudentHighlightColor(student.id, classSection.name, colorId)}
                                onDeleteRequest={() => {
                                  setConfirmDeleteStudentTarget({
                                    studentId: student.id,
                                    studentName: student.fullName,
                                    className: classSection.name
                                  });
                                  setActiveGearStudent(null);
                                }}
                              />
                              <span className="font-mono text-[10px] font-bold text-slate-700 bg-slate-100/90 px-1.5 py-0.5 rounded border border-slate-200 shadow-3xs shrink-0">
                                {student.id}
                              </span>
                              <span className="font-bold text-slate-900 text-xs truncate leading-tight">
                                {student.fullName}
                              </span>
                              <button
                                onClick={() => {
                                  setActiveClassName(classSection.name);
                                  setIsSidebarOpen(false);
                                }}
                                className="text-[9px] font-bold text-indigo-600 hover:text-indigo-800 bg-indigo-50 px-1.5 py-0.5 rounded hover:bg-indigo-100 transition-all cursor-pointer border border-indigo-100 ml-auto shrink-0"
                              >
                                {classSection.name}
                              </button>
                            </div>
                          </div>

                          <div className="text-right shrink-0">
                            <span className="text-[9px] font-bold text-slate-400 uppercase tracking-tight block">Grade</span>
                            <span className="text-[10.5px] font-black text-emerald-800 bg-emerald-50 px-1.5 py-0.2 rounded-md border border-emerald-100 inline-block">{finalGrade.toFixed(1)}</span>
                          </div>
                        </div>

                        {/* Collapsible Exam Scores */}
                        <div className="flex flex-col gap-1.5 bg-white p-2 rounded-xl border border-slate-200">
                          <button
                            type="button"
                            onClick={() => setExpandedSearchStudentId(isExpanded ? null : `${classSection.name}-${student.id}`)}
                            className="text-[9.5px] font-bold text-slate-500 hover:text-indigo-600 transition-colors inline-flex items-center gap-0.5 bg-slate-50 hover:bg-slate-100 px-2 py-0.5 border border-slate-200 rounded-md w-full justify-center cursor-pointer"
                          >
                            {isExpanded ? 'Hide Exam Scores' : 'Edit Exam Scores'}
                          </button>

                          {isExpanded && (
                            <div className="border-t border-slate-100 pt-2 mt-1 flex flex-col gap-1.5">
                              {assessments.length === 0 ? (
                                <span className="text-[9px] text-slate-400 italic text-center block py-1">No assessments in {classSection.name}</span>
                              ) : (
                                <div className="grid grid-cols-1 gap-1.5">
                                  {assessments.map(ass => {
                                    const score = classSection.grades[student.id]?.[ass.name] ?? 0;
                                    return (
                                      <div key={ass.name} className="flex items-center justify-between bg-slate-50 px-2 py-1 rounded-lg border border-slate-150">
                                        <span className="text-[10px] text-slate-550 truncate font-semibold w-1/2">{ass.name}</span>
                                        <div className="flex items-center gap-1 w-20">
                                          <input
                                            type="number"
                                            min={0}
                                            value={score}
                                            onChange={e => handleScoreChangeCrossSection(student.id, ass.name, e.target.value, classSection.name)}
                                            className="w-full text-right text-[11px] font-black bg-white border border-slate-300 rounded px-1.5 py-0.5 outline-none focus:border-indigo-500"
                                          />
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            )}
          </div>

          {/* Export Data Ledger System */}
          <div className="bg-slate-100 p-4 rounded-3xl border border-slate-200 shadow-3xs flex flex-col gap-2.5" id="sidebar-exports">
            <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <Download className="h-3.5 w-3.5 text-indigo-550" />
              Download Ledgers
            </h3>
            
            <div className="flex flex-col gap-2">
              <span className="text-[9px] font-bold text-slate-400 tracking-wider uppercase pl-0.5">Active Class:</span>
              <button
                onClick={handleExportGrades}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-white hover:bg-slate-50 border-slate-250 text-slate-700 hover:text-indigo-650 cursor-pointer'
                  : 'bg-slate-50 text-slate-450 border-slate-200 cursor-not-allowed'
                }`}
                id="export-grades-excel-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <FileSpreadsheet className="h-4 w-4 text-emerald-600 shrink-0" />
                  Grades Workbook (Excel)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-slate-400 shrink-0" />
              </button>

              <button
                onClick={handleExportAttendance}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-white hover:bg-slate-50 border-slate-250 text-slate-700 hover:text-indigo-650 cursor-pointer'
                  : 'bg-slate-50 text-slate-450 border-slate-200 cursor-not-allowed'
                }`}
                id="export-attendance-excel-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <Calendar className="h-4 w-4 text-sky-600 shrink-0" />
                  Attendance Log (Excel)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-slate-400 shrink-0" />
              </button>

              {/* Merged Exports System */}
              <span className="text-[9px] font-bold text-slate-400 tracking-wider uppercase pl-0.5 mt-1">Cross-Section Consolidated:</span>
              <button
                onClick={() => handleExportMerged('grades')}
                className="w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs bg-white hover:bg-indigo-50/50 border-slate-250 text-indigo-750 hover:text-indigo-900 transition-all cursor-pointer"
                id="export-merged-grades-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <FileSpreadsheet className="h-4 w-4 text-indigo-600 shrink-0" />
                  Merged Grades (All)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-indigo-400 shrink-0" />
              </button>

              <button
                onClick={() => handleExportMerged('attendance')}
                className="w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs bg-white hover:bg-teal-50/50 border-slate-250 text-teal-750 hover:text-teal-900 transition-all cursor-pointer"
                id="export-merged-attendance-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <Calendar className="h-4 w-4 text-teal-600 shrink-0" />
                  Merged Attendance (All)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-teal-400 shrink-0" />
              </button>
            </div>
          </div>
        </aside>

        {/* MAIN DASHBOARD PANEL */}
        <main className="flex-1 p-6 md:p-8 flex flex-col gap-6 overflow-x-auto" id="streamlit-main">
          {/* Active section ribbon banner */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col sm:flex-row justify-between sm:items-center gap-4" id="main-ribbon">
            <div>
              <div className="text-[10px] font-bold text-slate-400 tracking-widest uppercase">active working context</div>
              <h2 className="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2 mt-1">
                {classes.length === 0 ? "No Active Sections" : activeClass.name}
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                {classes.length === 0 ? (
                  "Create a class section using the '+' button in the sidebar or upload a roster."
                ) : (
                  <>Academic records and student rosters are synchronized and ready.</>
                )}
              </p>
            </div>
            <div className="flex gap-4">
              <div className="bg-slate-50 px-4 py-2 border border-slate-100 rounded-lg text-center min-w-24">
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Students</div>
                <div className="text-lg font-bold text-slate-800">{classes.length === 0 ? 0 : activeClass.students.length}</div>
              </div>
              <div className="bg-slate-50 px-4 py-2 border border-slate-100 rounded-lg text-center min-w-24">
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Assessments</div>
                <div className="text-lg font-bold text-slate-800">{classes.length === 0 ? 0 : activeClass.assessments.length}</div>
              </div>
            </div>
          </div>

          <div className="flex flex-col gap-6" id="app-working-views">
            {/* STAGED ROSTER VERIFICATION & PREVIEW HUB */}
              {pendingImport && (
                <div className={`rounded-xl border shadow-sm p-6 flex flex-col gap-5 animate-fade-in ${
                  pendingImport.short.length > 0
                    ? 'bg-amber-50 border-amber-200'
                    : 'bg-indigo-50/50 border-indigo-150'
                }`} id="short-id-compliance-container">
                  
                  <div className="flex items-start gap-3">
                    {pendingImport.short.length > 0 ? (
                      <AlertCircle className="h-6 w-6 text-amber-600 shrink-0 mt-0.5" />
                    ) : (
                      <Check className="h-6 w-6 text-emerald-600 shrink-0 mt-0.5" />
                    )}
                    <div>
                      <h3 className={`text-md font-extrabold tracking-tight ${
                        pendingImport.short.length > 0 ? 'text-amber-900' : 'text-slate-900'
                      }`}>
                        {pendingImport.short.length > 0 
                          ? 'Roster Verification Required: Short Student IDs Detected' 
                          : 'Roster Import Preview & Selection Checklist'
                        }
                      </h3>
                      <p className={`text-xs mt-1 leading-normal ${
                        pendingImport.short.length > 0 ? 'text-amber-700' : 'text-slate-500'
                      }`}>
                        {pendingImport.short.length > 0 
                          ? 'We detected student records with short ID numbers (6 or less digits). Decide which candidates you wish to discard or whitelist in your class directory below:'
                          : 'Review the list of students initialized from the uploading template. Choose what to keep and what to discard before committing:'
                        }
                      </p>
                    </div>
                  </div>

                  {/* Consolidated Checklist Preview Scrollbar Box */}
                  <div className="bg-white border border-slate-200 rounded-xl max-h-72 overflow-y-auto divide-y divide-slate-100 shadow-3xs text-left">
                    
                    {/* PRIORITY SECTION: Short Student IDs */}
                    {pendingImport.short.length > 0 && (
                      <div>
                        <div className="bg-amber-100/65 text-amber-950 font-bold px-3.5 py-2 text-[10px] uppercase tracking-wide flex items-center justify-between sticky top-0 z-10 border-b border-amber-200/55 backdrop-blur-xs">
                          <span className="flex items-center gap-1">
                            ⚠️ Priority Review: Short IDs ({pendingImport.short.length})
                          </span>
                          <span className="text-[9px] bg-amber-200 text-amber-900 font-bold px-1.5 py-0.5 rounded-md">
                            Discard Recommendation
                          </span>
                        </div>
                        <div className="divide-y divide-slate-100 bg-amber-50/15">
                          {pendingImport.short.map(st => {
                            const isChecked = selectedShortIds.includes(st.id);
                            return (
                              <div key={st.id} className="flex items-center justify-between p-3.5 text-xs hover:bg-amber-50/70 transition-colors">
                                <div className="flex items-center gap-3">
                                  <input 
                                    type="checkbox" 
                                    checked={isChecked}
                                    id={`chk-import-${st.id}`}
                                    onChange={() => {
                                      if (isChecked) {
                                        setSelectedShortIds(prev => prev.filter(x => x !== st.id));
                                      } else {
                                        setSelectedShortIds(prev => [...prev, st.id]);
                                      }
                                    }}
                                    className="h-4.5 w-4.5 text-amber-600 focus:ring-amber-500 rounded border-slate-300 cursor-pointer"
                                  />
                                  <label htmlFor={`chk-import-${st.id}`} className="flex flex-col sm:flex-row sm:items-center gap-1.5 cursor-pointer select-none">
                                    <span className="font-mono bg-amber-100 text-amber-900 px-1.5 py-0.5 rounded font-black text-[10.5px]">ID: {st.id}</span>
                                    <span className="font-bold text-slate-900">{st.fullName}</span>
                                  </label>
                                </div>
                                <span className="text-[10px] text-amber-600 font-mono font-bold">⚠️ {st.id.length} digits</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* GENERAL SECTION: Normal Student Records */}
                    {pendingImport.normal.length > 0 && (
                      <div>
                        <div className="bg-slate-55 text-slate-650 font-bold px-3.5 py-2 text-[10px] uppercase tracking-wide flex items-center justify-between sticky top-0 z-10 border-b border-slate-200 backdrop-blur-xs">
                          <span>✅ General Roster List ({pendingImport.normal.length})</span>
                          <span className="text-[9px] bg-slate-100 text-slate-500 font-bold px-1.5 py-0.5 rounded-md">
                            Normal ID Formats
                          </span>
                        </div>
                        <div className="divide-y divide-slate-100">
                          {pendingImport.normal.map(st => {
                            const isChecked = selectedShortIds.includes(st.id);
                            return (
                              <div key={st.id} className="flex items-center justify-between p-3.5 text-xs hover:bg-slate-50/50 transition-colors">
                                <div className="flex items-center gap-3">
                                  <input 
                                    type="checkbox" 
                                    checked={isChecked}
                                    id={`chk-import-${st.id}`}
                                    onChange={() => {
                                      if (isChecked) {
                                        setSelectedShortIds(prev => prev.filter(x => x !== st.id));
                                      } else {
                                        setSelectedShortIds(prev => [...prev, st.id]);
                                      }
                                    }}
                                    className="h-4.5 w-4.5 text-indigo-600 focus:ring-indigo-500 rounded border-slate-300 cursor-pointer"
                                  />
                                  <label htmlFor={`chk-import-${st.id}`} className="flex flex-col sm:flex-row sm:items-center gap-1.5 cursor-pointer select-none">
                                    <span className="font-mono bg-slate-100 text-slate-650 px-1.5 py-0.5 rounded font-bold text-[10.5px]">ID: {st.id}</span>
                                    <span className="font-bold text-slate-850">{st.fullName}</span>
                                  </label>
                                </div>
                                <span className="text-[10px] text-slate-400 font-mono">({st.id.length} digits)</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                  </div>

                  {/* Summary of checked items */}
                  <div className="text-[11px] font-bold text-slate-500 flex items-center gap-1 px-1">
                    <span>Selected to Keep:</span>
                    <span className="text-emerald-700 bg-emerald-50 px-2.5 py-0.5 rounded-full border border-emerald-100">
                      {selectedShortIds.length} of {pendingImport.short.length + pendingImport.normal.length} Students Active
                    </span>
                  </div>

                  {/* Interactive Buttons Bar */}
                  <div className="flex items-center gap-2.5 flex-wrap border-t border-slate-205/50 pt-3">
                    <button
                      onClick={handleConfirmPendingImport}
                      className="bg-indigo-600 hover:bg-indigo-750 text-white font-bold text-xs px-4.5 py-2.5 rounded-full shadow-md transition-all cursor-pointer active:scale-95 inline-flex items-center gap-1"
                    >
                      <span>Save & Complete Import ({selectedShortIds.length})</span>
                    </button>
                    
                    {pendingImport.short.length > 0 && (
                      <button
                        onClick={handleDeselectShortIDsChecklist}
                        className="bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-250 text-xs font-bold px-4.5 py-2.5 rounded-full transition-all cursor-pointer"
                        title="Unchecks all student records possessing 6 or less digits"
                      >
                        Deselect All Short IDs
                      </button>
                    )}

                    <button
                      onClick={handleSelectAllChecklist}
                      className="bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 text-xs font-bold px-4 py-2.5 rounded-full transition-all cursor-pointer"
                    >
                      Select All ({pendingImport.short.length + pendingImport.normal.length})
                    </button>

                    <button
                      onClick={() => setPendingImport(null)}
                      className="bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-bold px-4 py-2.5 rounded-full transition-colors cursor-pointer ml-auto"
                    >
                      Abort
                    </button>
                  </div>
                </div>
              )}

              {/* MATERIAL DESIGN SEGMENTED TIMELINE TABS */}
              <div className="flex bg-slate-100 p-1 rounded-full border border-slate-200 shrink-0 gap-1 overflow-x-auto max-w-full md:w-fit" id="tabs-ribbon">
                <button
                  onClick={() => setActiveTab('students')}
                  className={`flex items-center justify-center gap-2 px-5 py-2.5 text-xs md:text-sm font-bold rounded-full transition-all duration-200 select-none cursor-pointer whitespace-nowrap ${
                    activeTab === 'students'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'bg-transparent text-slate-650 hover:bg-slate-205 hover:text-slate-800'
                  }`}
                  id="tab-btn-students"
                >
                  <Users className="h-4 w-4 shrink-0" />
                  <span>Student Roster</span>
                </button>

                <button
                  onClick={() => setActiveTab('attendance')}
                  className={`flex items-center justify-center gap-2 px-5 py-2.5 text-xs md:text-sm font-bold rounded-full transition-all duration-200 select-none cursor-pointer whitespace-nowrap ${
                    activeTab === 'attendance'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'bg-transparent text-slate-650 hover:bg-slate-205 hover:text-slate-800'
                  }`}
                  id="tab-btn-attendance"
                >
                  <CalendarCheck className="h-4 w-4 shrink-0" />
                  <span>Attendance Registry</span>
                </button>

                <button
                  onClick={() => setActiveTab('grades')}
                  className={`flex items-center justify-center gap-2 px-5 py-2.5 text-xs md:text-sm font-bold rounded-full transition-all duration-200 select-none cursor-pointer whitespace-nowrap ${
                    activeTab === 'grades'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'bg-transparent text-slate-650 hover:bg-slate-250 hover:text-slate-800'
                  }`}
                  id="tab-btn-grades"
                >
                  <GraduationCap className="h-4 w-4 shrink-0" />
                  <span>Academic Grades</span>
                </button>
              </div>

              {/* ACTIVE WORKING VIEWPORTS */}
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm" id="tab-viewport">
                
                {/* STUDENTS VIEWPORT */}
                {activeTab === 'students' && (
                  <div className="flex flex-col gap-4">
                    <div className="flex justify-between items-center pb-2 border-b border-slate-100">
                      <h3 className="font-extrabold text-slate-900 text-md tracking-tight">
                        {classes.length === 0 ? "Class Roster: No Active Sections" : `Class Roster Listing: ${activeClass.name}`}
                      </h3>
                      <span className="text-xs bg-slate-100 px-2.5 py-1 text-slate-600 font-bold rounded-md">
                        Total Count: {classes.length === 0 ? 0 : activeClass.students.length} students
                      </span>
                    </div>

                    {classes.length === 0 ? (
                      <div className="py-14 flex flex-col items-center justify-center text-center">
                        <Users className="h-12 w-12 text-slate-300 mb-3" />
                        <h4 className="font-bold text-slate-800 text-sm">No Active Sections</h4>
                        <p className="text-xs text-slate-400 mt-1 max-w-sm leading-normal">
                          All class sections have been removed. Click the '+' button in the sidebar or below to create a new class section.
                        </p>
                        <button
                          type="button"
                          onClick={() => {
                            setNewClassNameInput('');
                            setCreatedClassStudents([]);
                            setCreatedClassFileError(null);
                            setCreatedClassFileSuccess(null);
                            setIsCreatingClass(true);
                          }}
                          className="mt-4 inline-flex items-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer"
                        >
                          <Plus className="h-4 w-4 text-white" strokeWidth={2.5} />
                          <span>Create New Class</span>
                        </button>
                      </div>
                    ) : activeClass.students.length === 0 ? (
                      <div className="py-12 flex flex-col items-center justify-center text-center">
                        <Users className="h-12 w-12 text-slate-300 mb-3" />
                        <h4 className="font-bold text-slate-800 text-sm">Roster list is currently empty</h4>
                        <p className="text-xs text-slate-400 mt-1 max-w-sm leading-normal">
                          Create a class section and import a spreadsheet roster to populate enrolled students.
                        </p>
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse" id="students-table-matrix">
                          <thead>
                            <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-widest border-b border-slate-100">
                              <th className="pl-4 pr-1.5 py-3 w-28 whitespace-nowrap">ID Number</th>
                              <th className="pl-1.5 pr-4 py-3 w-44 whitespace-nowrap">Student Name</th>
                              <th className="px-4 py-3">Academic Sneak Peek (Semester Progress)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || '';
                              const totalDates = getAttendanceDates(activeClass);
                              const attendedCount = totalDates.filter(d => activeClass.attendance[s.id]?.[d] === 'Attended').length;
                              const finalGrade = calculateFinalGrade(s.id, activeClass);

                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-slate-50/50'} transition-all`}>
                                  <td className={`pl-4 pr-1.5 py-3 whitespace-nowrap font-medium ${rowBgClass} ${colorInfo.border || ''}`}>
                                    <div className="flex items-center gap-1.5">
                                      <StudentGearBox
                                        student={s}
                                        className={activeClass.name}
                                        isOpen={activeGearStudent?.studentId === s.id && activeGearStudent?.className === activeClass.name}
                                        onToggle={() => {
                                          setActiveGearStudent(prev => 
                                            prev?.studentId === s.id && prev?.className === activeClass.name
                                              ? null
                                              : { studentId: s.id, className: activeClass.name }
                                          );
                                        }}
                                        onClose={() => setActiveGearStudent(null)}
                                        onSelectColor={(colorId) => setStudentHighlightColor(s.id, activeClass.name, colorId)}
                                        onDeleteRequest={() => {
                                          setConfirmDeleteStudentTarget({
                                            studentId: s.id,
                                            studentName: s.fullName,
                                            className: activeClass.name
                                          });
                                          setActiveGearStudent(null);
                                        }}
                                      />
                                      <span className="font-mono text-xs font-bold text-slate-700 bg-slate-100/90 px-1.5 py-0.5 rounded border border-slate-200 shadow-3xs">
                                        {s.id}
                                      </span>
                                    </div>
                                  </td>
                                  <td className={`pl-1.5 pr-4 py-3 font-bold text-slate-900 whitespace-nowrap ${rowBgClass}`}>
                                    <span>{s.fullName}</span>
                                  </td>
                                  <td className={`px-4 py-3 ${rowBgClass}`}>
                                    <div className="flex items-center gap-2 flex-wrap">
                                      {/* Attendance days attended / total lessons given (e.g. 12/16) */}
                                      <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold bg-sky-50 text-sky-800 border border-sky-200/70 shadow-3xs">
                                        <CalendarCheck className="h-3.5 w-3.5 text-sky-600 shrink-0" />
                                        <span>
                                          Attendance: <strong className="font-bold text-sky-950 font-mono">({attendedCount}/{totalDates.length})</strong>
                                        </span>
                                        {totalDates.length > 0 && (
                                          <span className="text-[10px] text-sky-700 font-mono font-bold">
                                            {Math.round((attendedCount / totalDates.length) * 100)}%
                                          </span>
                                        )}
                                      </div>

                                      {/* Assessment names with score next to it */}
                                      {activeClass.assessments.map(ass => {
                                        const score = activeClass.grades[s.id]?.[ass.name];
                                        const hasScore = score !== undefined && score !== null;
                                        return (
                                          <div key={ass.name} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs bg-white border border-slate-200 shadow-3xs">
                                            <span className="text-slate-500 font-medium text-[11px] truncate max-w-28">{ass.name}:</span>
                                            <span className={`font-mono text-[11px] font-bold ${hasScore ? 'text-slate-900 font-extrabold' : 'text-slate-300 italic font-normal'}`}>
                                              {hasScore ? score : '-'}
                                            </span>
                                          </div>
                                        );
                                      })}

                                      {/* Final Grade Badge */}
                                      {activeClass.assessments.length > 0 && (
                                        <div className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-extrabold bg-emerald-50 text-emerald-800 border border-emerald-200 shadow-3xs">
                                          <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-600">Final:</span>
                                          <span className="font-mono">{finalGrade.toFixed(1)}</span>
                                        </div>
                                      )}

                                      {/* Empty fallback when semester is just beginning */}
                                      {totalDates.length === 0 && activeClass.assessments.length === 0 && (
                                        <span className="text-xs text-slate-400 italic">
                                          No attendance or grades recorded yet
                                        </span>
                                      )}
                                    </div>
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

                    {classes.length === 0 ? (
                      <div className="py-12 text-center text-slate-400 italic">
                        No active class sections. Please create a section to take attendance.
                      </div>
                    ) : activeClass.students.length === 0 ? (
                      <div className="py-12 text-center text-slate-400 italic">
                        No registered student lists detected. Setup student information first.
                      </div>
                    ) : getAttendanceDates(activeClass).length === 0 ? (
                      <div className="p-8 text-center bg-slate-50 border border-dashed rounded-lg text-slate-400">
                        No attendance date logs found. Pick an attendance date above and click 'Take Attendance' to initiate your registry.
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse" id="attendance-table-matrix">
                          <thead>
                            <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-widest border-b border-slate-100">
                              <th className="pl-4 pr-1.5 py-3.5 whitespace-nowrap w-28">ID Number</th>
                              <th className="pl-1.5 pr-4 py-3.5 whitespace-nowrap w-44">Full Name</th>
                              {getAttendanceDates(activeClass).map(d => (
                                <th key={d} className="px-4 py-3.5 whitespace-nowrap text-center">
                                  <div className="inline-flex items-center gap-1.5 justify-center">
                                    <span className="font-mono text-xs font-bold text-slate-800">
                                      {formatDisplayDate(d)}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={() => setConfirmDeleteDateTarget({ date: d, displayDate: formatDisplayDate(d) })}
                                      title={`Delete attendance column for ${formatDisplayDate(d)}`}
                                      className="p-1 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors cursor-pointer"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </div>
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
                                  <td className={`pl-4 pr-1.5 py-3.5 text-slate-500 font-mono font-medium ${rowBgClass} ${colorInfo.border || ''}`}>{s.id}</td>
                                  <td className={`pl-1.5 pr-4 py-3.5 font-bold text-slate-800 whitespace-nowrap ${rowBgClass}`}>{s.fullName}</td>
                                  {getAttendanceDates(activeClass).map(d => {
                                    const status = activeClass.attendance[s.id]?.[d] || 'Not Taken';
                                    
                                    let pillClass = 'bg-slate-100 text-slate-500 border-slate-200';
                                    if (status === 'Attended') pillClass = 'bg-emerald-50 text-emerald-700 border-emerald-200';
                                    else if (status === 'Absent') pillClass = 'bg-rose-50 text-rose-700 border-rose-200';

                                    return (
                                      <td key={d} className={`px-4 py-3.5 text-center whitespace-nowrap ${rowBgClass}`}>
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

                    {/* Action Toolbar: Clean button with plus sign to Add Assessment & Total Weight breakdown */}
                    <div className="flex flex-wrap items-center justify-between gap-3 p-3.5 bg-slate-50 rounded-2xl border border-slate-200">
                      <div className="flex items-center gap-3">
                        <button
                          type="button"
                          onClick={handleOpenAddAssessment}
                          className="inline-flex items-center gap-2 bg-slate-900 hover:bg-indigo-600 active:scale-95 text-white font-bold text-xs px-4 py-2.5 rounded-xl shadow-xs transition-all cursor-pointer"
                          id="open-add-assessment-btn"
                        >
                          <Plus className="h-4 w-4" />
                          <span>Add Assessment</span>
                        </button>
                        <span className="text-xs text-slate-500 hidden sm:inline">
                          Click to specify assessment name, full mark, and weight.
                        </span>
                      </div>

                      <button
                        type="button"
                        onClick={handleOpenFinalGradeWeights}
                        className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-bold border transition-colors cursor-pointer bg-white hover:bg-emerald-50 border-slate-200 hover:border-emerald-300 text-slate-700 hover:text-emerald-800 shadow-3xs"
                        title="Click to view and adjust all assessment weights"
                        id="final-grade-weights-breakdown-btn"
                      >
                        <Settings className="h-3.5 w-3.5 text-emerald-600" />
                        <span>Weight Sum:</span>
                        <span className={`font-mono px-2 py-0.5 rounded text-[11px] font-extrabold ${
                          totalAllocatedWeight === 100 
                            ? 'bg-emerald-100 text-emerald-800' 
                            : totalAllocatedWeight < 100 
                              ? 'bg-amber-100 text-amber-800' 
                              : 'bg-rose-100 text-rose-800'
                        }`}>
                          {totalAllocatedWeight}% / 100%
                        </span>
                      </button>
                    </div>

                    {assessFeedback && (
                      <div className={`text-xs p-3 rounded-lg border ${
                        assessFeedback.type === 'success' 
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-100'
                        : 'bg-rose-50 text-rose-700 border-rose-10 border-rose-100'
                      }`}>
                        {assessFeedback.text}
                      </div>
                    )}

                    {classes.length === 0 ? (
                      <div className="py-12 text-center text-slate-400 italic">
                        No active class sections. Please create a section to record grades.
                      </div>
                    ) : activeClass.students.length === 0 ? (
                      <div className="py-12 text-center text-slate-400 italic">
                        No active roster exists. Create a class with students or import a roster to begin recording grades.
                      </div>
                    ) : activeClass.assessments.length === 0 ? (
                      <div className="p-8 text-center bg-slate-50 border border-dashed rounded-lg text-slate-400 flex flex-col items-center gap-2">
                        <p>No assessments registered yet. Click 'Add Assessment' above to create test columns.</p>
                        <button
                          type="button"
                          onClick={handleOpenAddAssessment}
                          className="mt-1 inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-slate-800 hover:bg-indigo-600 text-white rounded-lg transition-colors cursor-pointer"
                        >
                          <Plus className="h-3.5 w-3.5" />
                          <span>Add Assessment</span>
                        </button>
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse" id="grades-table-matrix">
                          <thead>
                            <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-widest border-b border-slate-100">
                              <th className="pl-4 pr-1.5 py-3.5 whitespace-nowrap w-28">ID Number</th>
                              <th className="pl-1.5 pr-4 py-3.5 whitespace-nowrap w-44">Full Name</th>
                              {activeClass.assessments.map(a => (
                                <th 
                                  key={a.name} 
                                  onClick={() => handleOpenEditAssessment(a)}
                                  className="px-4 py-3 whitespace-nowrap cursor-pointer hover:bg-slate-100/90 transition-colors group select-none"
                                  title="Click to edit test name, max mark, or weight"
                                >
                                  <div className="flex flex-col items-start gap-0.5">
                                    <div className="flex items-center gap-1.5 font-bold text-slate-800 text-xs group-hover:text-indigo-600 transition-colors">
                                      <span>{a.name}</span>
                                      <Settings className="h-3 w-3 text-slate-400 group-hover:text-indigo-500 opacity-60 group-hover:opacity-100 transition-opacity" />
                                    </div>
                                    <div className="flex items-center gap-1.5 text-[10px] text-slate-400 font-normal">
                                      <span>Full Mark: <strong className="font-semibold text-slate-600">{a.maxMark || 100}</strong></span>
                                      <span>•</span>
                                      <span>Weight: <strong className="font-semibold text-indigo-600">{a.weight}%</strong></span>
                                    </div>
                                  </div>
                                </th>
                              ))}
                              <th 
                                onClick={handleOpenFinalGradeWeights}
                                className="px-5 py-3.5 whitespace-nowrap text-right text-emerald-700 font-extrabold bg-emerald-50/60 hover:bg-emerald-100/80 transition-colors cursor-pointer select-none group"
                                title="Click to view and adjust all assessment weights"
                              >
                                <div className="flex flex-col items-end gap-0.5">
                                  <div className="inline-flex items-center gap-1.5 justify-end">
                                    <span>📊 Final Grade</span>
                                    <Settings className="h-3.5 w-3.5 text-emerald-600 group-hover:rotate-45 transition-transform" />
                                  </div>
                                  <span className="text-[10px] text-emerald-600 font-bold uppercase tracking-wider">
                                    Weight: {totalAllocatedWeight}% / 100%
                                  </span>
                                </div>
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
                                  <td className={`pl-4 pr-1.5 py-3.5 text-slate-500 font-mono font-medium ${rowBgClass} ${colorInfo.border || ''}`}>{s.id}</td>
                                  <td className={`pl-1.5 pr-4 py-3.5 font-bold text-slate-800 whitespace-nowrap ${rowBgClass}`}>{s.fullName}</td>
                                  {activeClass.assessments.map(a => {
                                    let score = activeClass.grades[s.id]?.[a.name] ?? 0;
                                    const maxMark = a.maxMark || 100;
                                    return (
                                      <td key={a.name} className={`px-4 py-3.5 whitespace-nowrap ${rowBgClass}`}>
                                        <div className="flex items-center gap-1">
                                          <input 
                                            type="number"
                                            min={0}
                                            max={maxMark}
                                            value={score}
                                            onChange={e => handleScoreChange(s.id, a.name, e.target.value)}
                                            className={`w-16 text-xs text-right ${s.color ? 'bg-white/80' : 'bg-slate-100'} hover:bg-slate-200/60 focus:bg-white rounded border border-transparent hover:border-slate-300 p-1 font-bold text-slate-800 animate-none`}
                                          />
                                          <span className="text-[10px] text-slate-400 font-mono">/{maxMark}</span>
                                        </div>
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
          </main>
        </div>

      {/* GOOGLE MATERIAL DESIGN CLASS CREATION MODAL WITH ISOLATED INTEGRATED ROSTER IMPORT */}
      {isCreatingClass && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-fade-in" id="class-creation-modal-overlay">
          <div className="bg-white rounded-3xl p-6 md:p-8 max-w-lg w-full shadow-2xl flex flex-col gap-5 border border-slate-150 max-h-[90vh] overflow-y-auto transform scale-100 transition-all duration-300">
            
            <div className="flex items-center justify-between border-b border-indigo-50 pb-3.5 animate-none" id="modal-header">
              <span className="text-base font-bold text-slate-900 tracking-tight flex items-center gap-1.5 font-sans">
                <Plus className="h-5 w-5 text-indigo-600" />
                Create New Class Section
              </span>
              <button
                type="button"
                onClick={() => {
                  setIsCreatingClass(false);
                  setNewClassNameInput('');
                  setCreatedClassStudents([]);
                  setCreatedClassFileError(null);
                  setCreatedClassFileSuccess(null);
                }}
                className="p-1.5 hover:bg-slate-100 rounded-full text-slate-400 border border-slate-200 transition-all cursor-pointer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleCreateClass} className="flex flex-col gap-4">
              
              {/* Class Name Input */}
              <div className="flex flex-col gap-1">
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Class Section Name</label>
                <input
                  type="text"
                  value={newClassNameInput}
                  onChange={e => setNewClassNameInput(e.target.value)}
                  placeholder="e.g. Physics 301, Section B"
                  className="w-full text-sm px-3.5 py-2.5 rounded-xl border border-slate-300 bg-white shadow-3xs focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 text-slate-800 transition-all"
                  autoFocus
                  required
                />
              </div>

              {/* Roster File Import Box */}
              <div className="flex flex-col gap-2 border-t border-slate-100 pt-4" id="modal-roster-uploader">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                    <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                    Student Roster (Optional Spreadsheet)
                  </label>
                </div>
                <p className="text-[11px] text-slate-400 leading-normal">
                  Upload a spreadsheet roster (.csv or .xlsx) to automatically populate student names and IDs.
                </p>

                <div 
                  onDragOver={handleModalRosterDragOver}
                  onDragLeave={handleModalRosterDragLeave}
                  onDrop={handleModalRosterDrop}
                  className={`border-2 border-dashed rounded-2xl p-5 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${
                    isDragOver 
                      ? 'border-indigo-500 bg-indigo-50/50 scale-[1.01]' 
                      : 'border-slate-200 hover:border-indigo-300 bg-slate-50 hover:bg-indigo-50/30'
                  }`}
                >
                  <input
                    type="file"
                    accept=".csv, .xlsx"
                    onChange={handleModalRosterFileUpload}
                    className="hidden"
                    id="modal-roster-file-input"
                  />
                  <label htmlFor="modal-roster-file-input" className="w-full h-full cursor-pointer flex flex-col items-center">
                    <Upload className="h-7 w-7 text-slate-400 mb-2" />
                    <span className="text-xs font-bold text-slate-700">Choose file or drag and drop here</span>
                    <span className="text-[10px] text-slate-400 mt-1">Accepts CSV or Excel workbooks</span>
                  </label>
                </div>

                {/* Feedback state notifications inside modal */}
                {createdClassFileError && (
                  <div className="text-[10.5px] bg-rose-50 text-rose-700 p-2.5 rounded-xl border border-rose-150 flex items-start gap-1.5 mt-1.5">
                    <AlertCircle className="h-3.5 w-3.5 shrink-0 mt-0.5 text-rose-600" />
                    <span>{createdClassFileError}</span>
                  </div>
                )}

                {createdClassFileSuccess && (
                  <div className="text-[10.5px] bg-emerald-50 text-emerald-800 p-2.5 rounded-xl border border-emerald-100 flex items-start gap-1.5 mt-1.5 font-bold">
                    <Check className="h-3.5 w-3.5 shrink-0 mt-0.5 text-emerald-600" />
                    <span>{createdClassFileSuccess}</span>
                  </div>
                )}

                {createdClassStudents.length > 0 && (
                  <div className="mt-2 p-3 bg-amber-50 rounded-2xl border border-amber-200 text-xs text-amber-800 flex flex-col gap-2">
                    <div className="flex items-start gap-2">
                      <AlertCircle className="h-4 w-4 shrink-0 mt-0.5 text-amber-600" />
                      <div>
                        <span className="font-bold">Short Student IDs Detected:</span>
                        <p className="text-[11px] leading-relaxed mt-0.5 text-amber-700">
                          There are <strong className="text-amber-900">{createdClassStudents.filter(st => st.id.length <= 6).length} students</strong> with ID numbers consisting of 6 or less digits.
                        </p>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setCreatedClassStudents(prev => prev.filter(st => st.id.length > 6));
                          setCreatedClassFileSuccess("Omitted short ID numbers from roster template!");
                        }}
                        className="bg-amber-600 hover:bg-amber-700 text-white font-bold px-3 py-1.5 rounded-xl text-[10px] shadow-3xs cursor-pointer transition-all active:scale-95"
                      >
                        Discard All Short IDs
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setCreatedClassFileSuccess("Kept short ID records in roster.");
                        }}
                        className="bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 font-semibold px-3 py-1.5 rounded-xl text-[10px] shadow-3xs cursor-pointer transition-all"
                      >
                        Keep All
                      </button>
                    </div>
                  </div>
                )}

                {createdClassStudents.length > 0 && (
                  <div className="mt-1.5 p-2.5 bg-indigo-50/30 rounded-xl border border-indigo-100/50 text-[11px] text-slate-600 overflow-hidden text-left">
                    <div className="flex items-center justify-between border-b border-indigo-150 pb-1 mb-1.5">
                      <span className="font-bold text-slate-700 block">Import Preview ({createdClassStudents.length}):</span>
                      <button
                        type="button"
                        onClick={() => setCreatedClassStudents([])}
                        className="text-[10px] text-rose-500 hover:text-rose-700 font-bold hover:underline"
                      >
                        Clear All
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto p-1 bg-white/50 rounded-lg">
                      {createdClassStudents.map(st => {
                        const isShort = st.id.length <= 6;
                        return (
                          <span 
                            key={st.id} 
                            className={`inline-flex items-center gap-1 bg-white border px-2 py-0.5 rounded-full text-[10px] font-semibold text-slate-700 font-mono shadow-3xs hover:bg-slate-50 transition-colors ${
                              isShort ? 'border-amber-300 bg-amber-50/30 text-amber-900' : 'border-slate-205'
                            }`}
                          >
                            <span>{st.firstName} ({st.id})</span>
                            <button
                              type="button"
                              onClick={() => {
                                setCreatedClassStudents(prev => prev.filter(x => x.id !== st.id));
                              }}
                              className="p-0.5 rounded-full hover:bg-slate-200 text-slate-450 hover:text-slate-700 flex items-center justify-center cursor-pointer transition-colors"
                              title="Discard student"
                            >
                              <X className="h-2 w-2" />
                            </button>
                          </span>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>

              {/* Action buttons */}
              <div className="flex items-center gap-2.5 justify-end mt-4 border-t border-slate-100 pt-4" id="modal-actions">
                <button
                  type="button"
                  onClick={() => {
                    setIsCreatingClass(false);
                    setNewClassNameInput('');
                    setCreatedClassStudents([]);
                    setCreatedClassFileError(null);
                    setCreatedClassFileSuccess(null);
                  }}
                  className="bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold px-4.5 py-2.5 rounded-full transition-all cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="bg-indigo-600 hover:bg-indigo-750 active:scale-[0.98] text-white text-xs font-bold px-5 py-2.5 rounded-full transition-all cursor-pointer shadow-md inline-flex items-center gap-1.5"
                >
                  <Plus className="h-4 w-4" />
                  Create Class
                </button>
              </div>

            </form>
          </div>
        </div>
      )}

      {/* Assessment Modal (Add New Assessment & Edit Assessment Parameters) */}
      {isAssessmentModalOpen && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200" id="assessment-configuration-modal">
          <div className="bg-white rounded-3xl p-6 md:p-7 max-w-md w-full shadow-2xl flex flex-col gap-5 border border-slate-200 animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="font-extrabold text-slate-900 text-base">
                  {assessmentModalMode === 'add' ? 'Add New Assessment' : 'Edit Assessment'}
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  {assessmentModalMode === 'add' 
                    ? `Create an assessment for ${activeClass.name}`
                    : `Update parameters for ${editingAssessmentOriginalName}`}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setIsAssessmentModalOpen(false)}
                className="p-1.5 hover:bg-slate-100 rounded-full text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleSaveAssessmentModal} className="flex flex-col gap-4">
              {/* Assessment Name */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Assessment Name</label>
                <input
                  type="text"
                  value={assessNameInput}
                  onChange={e => setAssessNameInput(e.target.value)}
                  placeholder="e.g. Test 1, Quiz 2, Final Project"
                  required
                  autoFocus
                  className="text-xs bg-slate-50 border border-slate-300 rounded-xl p-2.5 outline-none focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 text-slate-800 placeholder-slate-400 font-semibold transition-all"
                />
              </div>

              {/* Full Mark (Max Score) */}
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Full Mark (Max Score)</label>
                  <span className="text-[10px] text-slate-400">Total points possible</span>
                </div>
                <input
                  type="number"
                  min={1}
                  max={1000}
                  value={assessMaxMarkInput}
                  onChange={e => setAssessMaxMarkInput(parseInt(e.target.value) || 0)}
                  required
                  placeholder="e.g. 20, 50, 100"
                  className="text-xs bg-slate-50 border border-slate-300 rounded-xl p-2.5 outline-none focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 text-slate-800 font-mono font-bold transition-all"
                />
                <span className="text-[10.5px] text-slate-400">
                  Example: A test can be out of 20 or 100, while its weight determines final contribution.
                </span>
              </div>

              {/* Weight Out of 100% */}
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Weight (% of Final Grade)</label>
                  <span className="text-[10px] font-mono text-indigo-600 font-bold">Max total: 100%</span>
                </div>
                <div className="relative flex items-center">
                  <input
                    type="number"
                    min={0}
                    max={100}
                    value={assessWeightInput}
                    onChange={e => setAssessWeightInput(parseInt(e.target.value) || 0)}
                    required
                    className="w-full text-xs bg-slate-50 border border-slate-300 rounded-xl p-2.5 pr-8 outline-none focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 text-slate-800 font-mono font-bold transition-all"
                  />
                  <span className="absolute right-3 font-bold text-slate-400 text-xs">%</span>
                </div>

                {/* Real-time Weight Balance Preview */}
                {(() => {
                  const otherWeights = activeClass.assessments
                    .filter(a => assessmentModalMode === 'edit' ? a.name !== editingAssessmentOriginalName : true)
                    .reduce((sum, a) => sum + (Number(a.weight) || 0), 0);
                  const prospectiveTotal = otherWeights + (assessWeightInput || 0);
                  const isOver = prospectiveTotal > 100;
                  const availableForThis = Math.max(0, 100 - otherWeights);

                  return (
                    <div className={`p-2.5 rounded-xl border text-[11px] flex flex-col gap-1 mt-1 ${
                      isOver 
                        ? 'bg-rose-50 border-rose-200 text-rose-800' 
                        : prospectiveTotal === 100 
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                          : 'bg-indigo-50/70 border-indigo-200 text-indigo-900'
                    }`}>
                      <div className="flex items-center justify-between font-bold">
                        <span>Total Projected Weight:</span>
                        <span className="font-mono text-xs">{prospectiveTotal}% / 100%</span>
                      </div>
                      <div className="text-[10.5px]">
                        {isOver ? (
                          <span className="font-semibold text-rose-700">
                            ⚠️ Exceeds 100% by {prospectiveTotal - 100}%. Maximum allowed for this assessment is {availableForThis}%.
                          </span>
                        ) : prospectiveTotal === 100 ? (
                          <span className="font-semibold text-emerald-700">
                            ✓ Perfectly balanced! Reaches exactly 100% of final grade.
                          </span>
                        ) : (
                          <span>
                            Other assessments use {otherWeights}%. Remaining unallocated: {100 - prospectiveTotal}%.
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}
              </div>

              {assessModalError && (
                <div className="text-xs p-2.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-700 font-medium">
                  {assessModalError}
                </div>
              )}

              {/* Action buttons */}
              <div className="flex items-center justify-between gap-2 pt-3 border-t border-slate-100 mt-2">
                {assessmentModalMode === 'edit' && editingAssessmentOriginalName ? (
                  <button
                    type="button"
                    onClick={() => handleDeleteAssessment(editingAssessmentOriginalName)}
                    className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-rose-600 hover:text-rose-700 hover:bg-rose-50 rounded-xl transition-colors cursor-pointer"
                  >
                    <Trash2 className="h-4 w-4" />
                    <span>Delete Test</span>
                  </button>
                ) : <div />}

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setIsAssessmentModalOpen(false)}
                    className="px-4 py-2 text-xs font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-xl transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl shadow-sm transition-all cursor-pointer inline-flex items-center gap-1.5 active:scale-95"
                  >
                    <Check className="h-4 w-4" />
                    <span>{assessmentModalMode === 'add' ? 'Add Assessment' : 'Save Changes'}</span>
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Final Grade Breakdown & Weights Configuration Modal */}
      {isFinalGradeModalOpen && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200" id="final-grade-weights-modal">
          <div className="bg-white rounded-3xl p-6 md:p-7 max-w-md w-full shadow-2xl flex flex-col gap-5 border border-slate-200 animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="font-extrabold text-slate-900 text-base flex items-center gap-2">
                  <span>📊 Final Grade Weight Distribution</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  Set how much each test weighs. All tests combined must weigh 100%.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setIsFinalGradeModalOpen(false)}
                className="p-1.5 hover:bg-slate-100 rounded-full text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {activeClass.assessments.length === 0 ? (
              <div className="py-8 text-center text-slate-400 italic text-xs">
                No assessments registered yet for this class section.
              </div>
            ) : (
              <div className="flex flex-col gap-3">
                <div className="flex items-center justify-between text-[11px] text-slate-500 font-bold uppercase tracking-wider px-1">
                  <span>Assessment Name & Max Mark</span>
                  <span>Weight (%)</span>
                </div>

                <div className="max-h-60 overflow-y-auto flex flex-col gap-2 p-1">
                  {activeClass.assessments.map(a => {
                    const currentVal = tempWeights[a.name] !== undefined ? tempWeights[a.name] : a.weight;
                    return (
                      <div key={a.name} className="flex items-center justify-between p-3 bg-slate-50 rounded-2xl border border-slate-200 shadow-3xs">
                        <div className="min-w-0 pr-2">
                          <span className="font-bold text-slate-800 text-xs block truncate">{a.name}</span>
                          <span className="text-[10px] text-slate-400 font-mono">Full Mark: {a.maxMark || 100} pts</span>
                        </div>
                        <div className="relative flex items-center shrink-0">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={currentVal}
                            onChange={e => {
                              const val = parseInt(e.target.value) || 0;
                              setTempWeights(prev => ({ ...prev, [a.name]: val }));
                            }}
                            className="w-20 text-xs bg-white border border-slate-300 rounded-xl p-2 pr-6 text-right font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600"
                          />
                          <span className="absolute right-2.5 font-bold text-slate-400 text-xs">%</span>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Auto distribute helper button */}
                {activeClass.assessments.length > 1 && (
                  <div className="flex justify-end">
                    <button
                      type="button"
                      onClick={handleAutoDistributeWeights}
                      className="text-[11px] font-bold text-indigo-600 hover:text-indigo-800 hover:underline cursor-pointer"
                    >
                      Auto-Distribute Evenly
                    </button>
                  </div>
                )}

                {/* Weight Total Status Banner */}
                {(() => {
                  const currentSum = Object.values(tempWeights).reduce((sum, w) => sum + (Number(w) || 0), 0);
                  const isOver = currentSum > 100;
                  const isBalanced = currentSum === 100;

                  return (
                    <div className={`p-3 rounded-2xl border flex flex-col gap-1 text-xs ${
                      isOver
                        ? 'bg-rose-50 border-rose-200 text-rose-800'
                        : isBalanced
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                          : 'bg-amber-50 border-amber-200 text-amber-800'
                    }`}>
                      <div className="flex items-center justify-between font-extrabold text-sm">
                        <span>Combined Weights Total:</span>
                        <span className="font-mono">{currentSum}% / 100%</span>
                      </div>
                      <div className="text-[11px]">
                        {isOver ? (
                          <span className="font-semibold text-rose-700">
                            ⚠️ Total exceeds 100% by {currentSum - 100}%. Total weight cannot exceed 100%.
                          </span>
                        ) : isBalanced ? (
                          <span className="font-semibold text-emerald-700">
                            ✓ Balanced: exactly 100% of the final grade is allocated.
                          </span>
                        ) : (
                          <span className="font-semibold text-amber-700">
                            Notice: {100 - currentSum}% of the final grade remains unallocated.
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}

                {finalGradeModalError && (
                  <div className="text-xs p-2.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-700 font-medium">
                    {finalGradeModalError}
                  </div>
                )}

                {/* Action buttons */}
                <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setIsFinalGradeModalOpen(false)}
                    className="px-4 py-2 text-xs font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-xl transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={handleSaveFinalGradeWeights}
                    disabled={Object.values(tempWeights).reduce((sum, w) => sum + (Number(w) || 0), 0) > 100}
                    className="px-5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl shadow-sm transition-all cursor-pointer inline-flex items-center gap-1.5 active:scale-95"
                  >
                    <Check className="h-4 w-4" />
                    <span>Save Weights</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Delete Attendance Date Column Confirmation Modal */}
      {confirmDeleteDateTarget && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-3xs animate-in fade-in duration-150" id="delete-attendance-date-modal">
          <div className="bg-white rounded-2xl max-w-sm w-full p-5 shadow-2xl border border-rose-200 flex flex-col gap-4 animate-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-rose-100 text-rose-600 shrink-0">
                <Trash2 className="h-5 w-5" />
              </div>
              <div className="flex-1 min-w-0">
                <h4 className="text-sm font-bold text-slate-900">Delete Attendance Date Column</h4>
                <p className="text-xs text-slate-600 mt-1 leading-normal">
                  Are you sure you want to delete the attendance date column for <strong className="text-slate-900 font-bold">{confirmDeleteDateTarget.displayDate}</strong> in <strong className="text-slate-900 font-bold">{activeClass.name}</strong>?
                </p>
                <p className="text-[11px] text-rose-700 bg-rose-50 border border-rose-150 rounded-lg p-2 mt-2 font-medium">
                  ⚠️ Warning: All attendance check logs for this date will be permanently deleted.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-1 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setConfirmDeleteDateTarget(null)}
                className="px-3.5 py-1.5 text-xs font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => handleDeleteAttendanceDate(confirmDeleteDateTarget.date)}
                className="px-3.5 py-1.5 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-lg shadow-sm transition-all cursor-pointer flex items-center gap-1.5 active:scale-95"
              >
                <Trash2 className="h-3.5 w-3.5" />
                <span>Delete Date Column</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete Student Warning Confirmation Modal */}
      {confirmDeleteStudentTarget && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-3xs animate-in fade-in duration-150" id="delete-student-warning-modal">
          <div className="bg-white rounded-2xl max-w-sm w-full p-5 shadow-2xl border border-rose-200 flex flex-col gap-4 animate-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-rose-100 text-rose-600 shrink-0">
                <AlertCircle className="h-5 w-5" />
              </div>
              <div className="flex-1 min-w-0">
                <h4 className="text-sm font-bold text-slate-900">Delete Student Confirmation</h4>
                <p className="text-xs text-slate-600 mt-1 leading-normal">
                  Are you sure you want to remove <strong className="text-slate-900 font-bold">{confirmDeleteStudentTarget.studentName}</strong> (ID: {confirmDeleteStudentTarget.studentId}) from <strong className="text-slate-900 font-bold">{confirmDeleteStudentTarget.className}</strong>?
                </p>
                <p className="text-[11px] text-rose-700 bg-rose-50 border border-rose-150 rounded-lg p-2 mt-2 font-medium">
                  ⚠️ Warning: All attendance records and graded scores for this student in this class will be permanently removed.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-1 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setConfirmDeleteStudentTarget(null)}
                className="px-3.5 py-1.5 text-xs font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => handleDeleteStudentFromClass(confirmDeleteStudentTarget.studentId, confirmDeleteStudentTarget.className)}
                className="px-3.5 py-1.5 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-lg shadow-sm transition-all cursor-pointer flex items-center gap-1.5 active:scale-95"
              >
                <Trash2 className="h-3.5 w-3.5" />
                <span>Confirm Delete</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* footer credits */}
      <footer className="bg-slate-100 px-6 py-3 border-t border-slate-200 text-center text-[10px] text-slate-500 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2" id="omni-class-footer">
        <span>Omni Class — Academic Roster, Attendance & Grade Records Manager</span>
        <span className="font-mono text-slate-400">Offline-Ready & Cloud-Connected</span>
      </footer>
    </div>
  );
}
