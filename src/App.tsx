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
  ChevronLeft,
  ChevronDown,
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
  Settings,
  Pencil,
  Clock,
  CalendarDays,
  CalendarPlus,
  Filter,
  Tag,
  FileText,
  List,
  LayoutGrid
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { Student, Assessment, ClassSection, DayOfWeek, ClassScheduleSlot, CalendarEvent, CalendarEventType } from './types';
import { 
  parseAnyExcelFile, 
  applyAutoColumnWidths, 
  applyAutoColumnWidthsToTable,
  ParsedClassSectionData 
} from './excelUtils';

// Import Firebase config & handlers
import { auth, db, googleProvider, handleFirestoreError, OperationType } from './firebase';
import { signInWithPopup, signOut, onAuthStateChanged, User, signInWithEmailAndPassword, createUserWithEmailAndPassword } from 'firebase/auth';
import { doc, setDoc, onSnapshot } from 'firebase/firestore';

// Initial classes set with sample weekly lecture schedules
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
    },
    schedule: [
      { id: 'slot-1', day: 'Sunday', startTime: '08:30 AM', endTime: '11:20 AM' },
      { id: 'slot-2', day: 'Tuesday', startTime: '01:00 PM', endTime: '02:00 PM' },
      { id: 'slot-3', day: 'Thursday', startTime: '09:00 AM', endTime: '10:30 AM' }
    ]
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
    grades: {},
    schedule: [
      { id: 'slot-4', day: 'Monday', startTime: '09:30 AM', endTime: '11:00 AM' },
      { id: 'slot-5', day: 'Wednesday', startTime: '01:00 PM', endTime: '03:00 PM' }
    ]
  }
];

export const DAYS_OF_WEEK: DayOfWeek[] = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday'
];

export const formatDurationMinutes = (minutes: number): string => {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (h > 0 && m > 0) return `${h}h ${m}m`;
  if (h > 0) return `${h}h`;
  return `${m}m`;
};

export const parseTimeToMinutes = (timeStr: string): number => {
  if (!timeStr) return 0;
  const trimmed = timeStr.trim().toLowerCase();
  const isPM = trimmed.includes('pm');
  const isAM = trimmed.includes('am');
  const clean = trimmed.replace(/[ap]m/g, '').trim();
  const parts = clean.split(':');
  let hours = parseInt(parts[0], 10) || 0;
  const minutes = parseInt(parts[1], 10) || 0;
  if (isPM && hours < 12) hours += 12;
  if (isAM && hours === 12) hours = 0;
  return hours * 60 + minutes;
};

export const TIME_SLOTS = [
  { label: '08:00 AM', startMin: 8 * 60, endMin: 9 * 60, display: '08:00 - 09:00 AM' },
  { label: '09:00 AM', startMin: 9 * 60, endMin: 10 * 60, display: '09:00 - 10:00 AM' },
  { label: '10:00 AM', startMin: 10 * 60, endMin: 11 * 60, display: '10:00 - 11:00 AM' },
  { label: '11:00 AM', startMin: 11 * 60, endMin: 12 * 60, display: '11:00 - 12:00 PM' },
  { label: '12:00 PM', startMin: 12 * 60, endMin: 13 * 60, display: '12:00 - 01:00 PM' },
  { label: '01:00 PM', startMin: 13 * 60, endMin: 14 * 60, display: '01:00 - 02:00 PM' },
  { label: '02:00 PM', startMin: 14 * 60, endMin: 15 * 60, display: '02:00 - 03:00 PM' },
  { label: '03:00 PM', startMin: 15 * 60, endMin: 16 * 60, display: '03:00 - 04:00 PM' },
  { label: '04:00 PM', startMin: 16 * 60, endMin: 17 * 60, display: '04:00 - 05:00 PM' },
  { label: '05:00 PM', startMin: 17 * 60, endMin: 18 * 60, display: '05:00 - 06:00 PM' },
];

export const DAY_OF_WEEK_NAMES: DayOfWeek[] = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

export const INITIAL_CALENDAR_EVENTS: CalendarEvent[] = [
  {
    id: 'evt-1',
    date: '2026-11-11',
    title: 'Midterm Exam - Class Section A',
    time: '10:00 AM',
    className: 'Class Section A',
    type: 'exam',
    description: 'Midterm examination on Chapters 1 to 5. Bring scientific calculators and formula sheets.',
    color: '#E11D48'
  },
  {
    id: 'evt-2',
    date: '2026-11-04',
    title: 'Attendance Check & Quiz Review',
    time: '01:30 PM',
    className: 'Class Section A',
    type: 'note',
    description: 'Review registry records and prep questions for Friday assessment.',
    color: '#4D5A8F'
  },
  {
    id: 'evt-3',
    date: '2026-11-16',
    title: 'Department Faculty Meeting',
    time: '02:00 PM',
    type: 'meeting',
    description: 'Academic council term evaluation in Conference Room B.',
    color: '#2563EB'
  },
  {
    id: 'evt-4',
    date: '2026-11-20',
    title: 'Term Project Submissions Due',
    time: '11:59 PM',
    className: 'Class Section A',
    type: 'deadline',
    description: 'Final submission deadline for group reports and project presentations.',
    color: '#D97706'
  }
];

export const CLASS_COLOR_PALETTES = [
  { bg: 'bg-[#E3EAFA]', border: 'border-[#8F9DCC]', text: 'text-[#252B3B]', badge: 'bg-[#4D5A8F] text-white', hover: 'hover:border-[#4D5A8F] hover:bg-[#CED8F2] hover:shadow-xs', dot: 'bg-[#4D5A8F]' },
  { bg: 'bg-[#F5F8FF]', border: 'border-[#B3C0E6]', text: 'text-[#252B3B]', badge: 'bg-[#6D7CB0] text-white', hover: 'hover:border-[#6D7CB0] hover:bg-[#E3EAFA] hover:shadow-xs', dot: 'bg-[#6D7CB0]' },
  { bg: 'bg-[#CED8F2]/40', border: 'border-[#6D7CB0]', text: 'text-[#252B3B]', badge: 'bg-[#4D5A8F] text-white', hover: 'hover:border-[#4D5A8F] hover:bg-[#CED8F2]/80 hover:shadow-xs', dot: 'bg-[#4D5A8F]' },
  { bg: 'bg-white', border: 'border-[#8F9DCC]', text: 'text-[#252B3B]', badge: 'bg-[#252B3B] text-white', hover: 'hover:border-[#4D5A8F] hover:bg-[#F5F8FF] hover:shadow-xs', dot: 'bg-[#6D7CB0]' },
  { bg: 'bg-[#E3EAFA]/70', border: 'border-[#4D5A8F]/40', text: 'text-[#252B3B]', badge: 'bg-[#4D5A8F] text-white', hover: 'hover:border-[#4D5A8F] hover:bg-[#CED8F2] hover:shadow-xs', dot: 'bg-[#4D5A8F]' },
  { bg: 'bg-[#F0F4FD]', border: 'border-[#B3C0E6]', text: 'text-[#252B3B]', badge: 'bg-[#6D7CB0] text-white', hover: 'hover:border-[#4D5A8F] hover:bg-[#E3EAFA] hover:shadow-xs', dot: 'bg-[#8F9DCC]' },
];

// Date formatter: converts date format everywhere to dd-mm-yy (e.g. 14-05-26)
export const formatDisplayDate = (d: string): string => {
  if (!d) return '';
  const trimmed = d.trim();

  // 1. Matches YYYY-MM-DD or YYYY/MM/DD
  const ymdMatch = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (ymdMatch) {
    const year = ymdMatch[1].slice(-2);
    const month = String(parseInt(ymdMatch[2], 10)).padStart(2, '0');
    const day = String(parseInt(ymdMatch[3], 10)).padStart(2, '0');
    return `${day}-${month}-${year}`;
  }

  // 2. Matches DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const day = String(parseInt(dmyMatch[1], 10)).padStart(2, '0');
    const month = String(parseInt(dmyMatch[2], 10)).padStart(2, '0');
    const year = dmyMatch[3].slice(-2);
    return `${day}-${month}-${year}`;
  }

  // 3. Matches DD-MM-YY or DD/MM/YY
  const dmyShortMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/);
  if (dmyShortMatch) {
    const day = String(parseInt(dmyShortMatch[1], 10)).padStart(2, '0');
    const month = String(parseInt(dmyShortMatch[2], 10)).padStart(2, '0');
    const year = dmyShortMatch[3];
    return `${day}-${month}-${year}`;
  }

  // 4. Fallback: Parse with Date object
  const parsed = new Date(trimmed);
  if (!isNaN(parsed.getTime())) {
    const day = String(parsed.getDate()).padStart(2, '0');
    const month = String(parsed.getMonth() + 1).padStart(2, '0');
    const year = String(parsed.getFullYear()).slice(-2);
    return `${day}-${month}-${year}`;
  }

  return trimmed;
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
  const buttonRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [coords, setCoords] = useState<{ top?: number; bottom?: number; left: number; openUpwards: boolean } | null>(null);

  useEffect(() => {
    if (isOpen && buttonRef.current) {
      const updatePosition = () => {
        if (!buttonRef.current) return;
        const rect = buttonRef.current.getBoundingClientRect();
        const popoverWidth = 230;
        const popoverHeight = 260;
        const spaceBelow = window.innerHeight - rect.bottom;
        const openUpwards = spaceBelow < popoverHeight && rect.top > popoverHeight;

        let left = rect.left;
        if (left + popoverWidth > window.innerWidth - 12) {
          left = Math.max(12, window.innerWidth - popoverWidth - 12);
        }

        setCoords({
          top: openUpwards ? undefined : rect.bottom + 6,
          bottom: openUpwards ? window.innerHeight - rect.top + 6 : undefined,
          left: Math.max(8, left),
          openUpwards
        });
      };

      updatePosition();
    } else {
      setCoords(null);
    }
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;
    const handleOutside = (e: MouseEvent) => {
      if (
        menuRef.current && !menuRef.current.contains(e.target as Node) &&
        buttonRef.current && !buttonRef.current.contains(e.target as Node)
      ) {
        onClose();
      }
    };
    const handleScrollOrResize = (e: Event) => {
      // If scroll happens inside the menu itself, do not close
      if (menuRef.current && e.target instanceof Node && menuRef.current.contains(e.target)) {
        return;
      }
      onClose();
    };

    document.addEventListener('mousedown', handleOutside);
    window.addEventListener('scroll', handleScrollOrResize, true);
    window.addEventListener('resize', handleScrollOrResize);
    return () => {
      document.removeEventListener('mousedown', handleOutside);
      window.removeEventListener('scroll', handleScrollOrResize, true);
      window.removeEventListener('resize', handleScrollOrResize);
    };
  }, [isOpen, onClose]);

  return (
    <div className="inline-flex items-center">
      <button
        ref={buttonRef}
        type="button"
        onClick={(e) => {
          e.stopPropagation();
          onToggle();
        }}
        className={`p-1.5 rounded-lg border transition-all cursor-pointer ${
          isOpen
            ? 'bg-[#CED8F2] text-[#4D5A8F] border-[#4D5A8F] ring-2 ring-[#4D5A8F]/20'
            : 'bg-white hover:bg-[#F5F8FF] text-[#252B3B] border-[#E8ECF5] shadow-3xs'
        }`}
        title={`Settings for ${student.fullName}`}
        aria-label={`Settings for ${student.fullName}`}
      >
        <Settings className="h-3.5 w-3.5" />
      </button>

      {isOpen && coords && (
        <div 
          ref={menuRef}
          onClick={(e) => e.stopPropagation()}
          style={{
            position: 'fixed',
            top: coords.top !== undefined ? `${coords.top}px` : undefined,
            bottom: coords.bottom !== undefined ? `${coords.bottom}px` : undefined,
            left: `${coords.left}px`,
            zIndex: 9999,
          }}
          className="w-58 bg-white rounded-2xl shadow-xl border border-[#E8ECF5] p-3 flex flex-col gap-2.5 animate-in fade-in zoom-in-95 duration-150 text-left"
        >
          {/* Header */}
          <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-2">
            <div className="min-w-0 pr-1">
              <span className="text-xs font-bold text-[#252B3B] truncate block">{student.fullName}</span>
              <span className="text-[10px] text-[#6D7CB0] font-mono font-bold">ID: {student.id}</span>
            </div>
            <button
              type="button"
              onClick={onClose}
              className="p-1 text-[#747D91] hover:text-[#252B3B] rounded-lg hover:bg-[#E3EAFA] cursor-pointer transition-colors"
            >
              <X className="h-3 w-3" />
            </button>
          </div>

          {/* 1. HIGHLIGHT COLOR (ABOVE DELETE OPTION) */}
          <div className="flex flex-col gap-1.5">
            <span className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider">
              Highlight Color:
            </span>
            <div className="grid grid-cols-2 gap-1.5">
              <button
                type="button"
                onClick={() => onSelectColor('none')}
                className={`flex items-center gap-1.5 px-2 py-1 rounded-lg text-[10px] font-bold border transition-all cursor-pointer ${
                  !student.color || student.color === 'none'
                    ? 'bg-[#E3EAFA] text-[#4D5A8F] border-[#4D5A8F] ring-1 ring-[#4D5A8F] font-extrabold'
                    : 'bg-white text-[#252B3B] border-[#E8ECF5] hover:bg-[#F5F8FF]'
                }`}
              >
                <X className="h-3 w-3 text-[#252B3B]" />
                <span>None</span>
              </button>

              {HIGHLIGHT_COLORS.map(c => {
                const isSelected = student.color === c.id;
                return (
                  <button
                    key={c.id}
                    type="button"
                    onClick={() => onSelectColor(c.id)}
                    className={`flex items-center gap-1.5 px-2 py-1 rounded-lg text-[10px] font-bold border transition-all cursor-pointer ${
                      isSelected
                        ? `${c.bgClass} text-slate-900 border-[#4D5A8F] ring-2 ring-[#4D5A8F] font-black shadow-3xs`
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

          <div className="h-px bg-[#E8ECF5]" />

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

interface GradeScoreInputProps {
  score: number | undefined;
  maxMark: number;
  highlightColor?: string;
  onSave: (val: number | undefined) => void;
  className?: string;
}

function GradeScoreInput({
  score,
  maxMark,
  highlightColor,
  onSave,
  className = 'w-16'
}: GradeScoreInputProps) {
  const isExplicitlySet = score !== undefined && score !== null;
  const [isFocused, setIsFocused] = useState(false);
  const [typedValue, setTypedValue] = useState<string>('');

  const handleFocus = (e: React.FocusEvent<HTMLInputElement>) => {
    setIsFocused(true);
    if (!isExplicitlySet) {
      // Default zero disappears when clicked so user can type freely without creating e.g. 50
      setTypedValue('');
    } else {
      // If already set, show current value and select all for quick typing
      setTypedValue(String(score));
      e.target.select();
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.value;
    // Allow empty string so user can clear out mark
    if (raw === '') {
      setTypedValue('');
      onSave(undefined);
      return;
    }

    // Only allow positive numbers (and optional decimal point)
    if (!/^\d*\.?\d*$/.test(raw)) return;

    let num = parseFloat(raw);
    if (!isNaN(num)) {
      if (num > maxMark) num = maxMark;
      if (num < 0) num = 0;
      setTypedValue(raw);
      onSave(num);
    } else {
      setTypedValue(raw);
    }
  };

  const handleBlur = () => {
    setIsFocused(false);
    if (typedValue.trim() === '') {
      onSave(undefined);
    } else {
      let num = parseFloat(typedValue);
      if (isNaN(num)) {
        onSave(undefined);
      } else {
        if (num > maxMark) num = maxMark;
        if (num < 0) num = 0;
        onSave(num);
      }
    }
  };

  // When not focused:
  // - If explicitly set (including a manually entered 0): show String(score) in bold dark font (not grayed out).
  // - If unset: show '0' in grayed-out italic styling.
  // When focused:
  // - Show typedValue (which is empty '' if was unset, so default zero disappeared!).
  const displayVal = isFocused ? typedValue : (isExplicitlySet ? String(score) : '0');
  const isGrayedOut = !isFocused && !isExplicitlySet;

  return (
    <input
      type="text"
      inputMode="decimal"
      value={displayVal}
      onFocus={handleFocus}
      onChange={handleChange}
      onBlur={handleBlur}
      className={`${className} text-xs text-right rounded border p-1 transition-all outline-none [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none ${
        highlightColor && highlightColor !== 'none' ? 'bg-white/90' : 'bg-slate-100/90'
      } ${
        isFocused
          ? 'bg-white ring-2 ring-indigo-500 border-indigo-400 font-bold text-slate-900 shadow-sm'
          : 'hover:bg-slate-200/60 border-transparent hover:border-slate-300'
      } ${
        isGrayedOut
          ? 'text-slate-400 font-normal italic opacity-70 select-all'
          : 'text-slate-900 font-extrabold'
      }`}
    />
  );
}

interface DDMMYYDatePickerProps {
  value: string;
  onChange: (val: string) => void;
  className?: string;
  id?: string;
}

// Helpers for DDMMYYDatePicker
const parseDatePickerParts = (d: string) => {
  if (!d) {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() };
  }
  const trimmed = d.trim();
  const ymd = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (ymd) {
    return { year: parseInt(ymd[1], 10), month: parseInt(ymd[2], 10), day: parseInt(ymd[3], 10) };
  }
  const dmy = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmy) {
    return { day: parseInt(dmy[1], 10), month: parseInt(dmy[2], 10), year: parseInt(dmy[3], 10) };
  }
  const dmyShort = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/);
  if (dmyShort) {
    return { day: parseInt(dmyShort[1], 10), month: parseInt(dmyShort[2], 10), year: 2000 + parseInt(dmyShort[3], 10) };
  }
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() };
};

const formatDateToDDMMYY = (year: number, month: number, day: number): string => {
  const yr = String(year).slice(-2);
  const m = String(month).padStart(2, '0');
  const d = String(day).padStart(2, '0');
  return `${d}-${m}-${yr}`;
};

const formatDateToYMD = (year: number, month: number, day: number): string => {
  const m = String(month).padStart(2, '0');
  const d = String(day).padStart(2, '0');
  return `${year}-${m}-${d}`;
};

const getTodayYMD = (): string => {
  const now = new Date();
  return formatDateToYMD(now.getFullYear(), now.getMonth() + 1, now.getDate());
};

export const formatFullDateString = (ymdStr: string): string => {
  if (!ymdStr) return '';
  const parts = ymdStr.split('-');
  if (parts.length !== 3) return ymdStr;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  const d = new Date(year, month, day);
  return d.toLocaleDateString('en-US', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });
};

export const getRelativeDateLabel = (ymdStr: string): string | null => {
  const today = getTodayYMD();
  if (ymdStr === today) return 'Today';
  
  const todayParts = today.split('-').map(Number);
  const targetParts = ymdStr.split('-').map(Number);
  const dToday = new Date(todayParts[0], todayParts[1] - 1, todayParts[2]);
  const dTarget = new Date(targetParts[0], targetParts[1] - 1, targetParts[2]);
  const diffDays = Math.round((dTarget.getTime() - dToday.getTime()) / (1000 * 60 * 60 * 24));
  
  if (diffDays === 1) return 'Tomorrow';
  if (diffDays === -1) return 'Yesterday';
  if (diffDays > 1 && diffDays <= 7) return `In ${diffDays} days`;
  if (diffDays < -1 && diffDays >= -7) return `${Math.abs(diffDays)} days ago`;
  return null;
};

function DDMMYYDatePicker({ value, onChange, className = '', id }: DDMMYYDatePickerProps) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  
  const selectedParts = parseDatePickerParts(value);
  const [viewYear, setViewYear] = useState<number>(selectedParts.year);
  const [viewMonth, setViewMonth] = useState<number>(selectedParts.month); // 1-12
  
  const [typedText, setTypedText] = useState<string>(
    formatDateToDDMMYY(selectedParts.year, selectedParts.month, selectedParts.day)
  );
  const [isInputFocused, setIsInputFocused] = useState(false);

  useEffect(() => {
    if (!isInputFocused) {
      const p = parseDatePickerParts(value);
      setTypedText(formatDateToDDMMYY(p.year, p.month, p.day));
      setViewYear(p.year);
      setViewMonth(p.month);
    }
  }, [value, isInputFocused]);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    if (isOpen) {
      document.addEventListener('mousedown', handleOutsideClick);
    }
    return () => {
      document.removeEventListener('mousedown', handleOutsideClick);
    };
  }, [isOpen]);

  const handleSelectDay = (day: number) => {
    const newYmd = formatDateToYMD(viewYear, viewMonth, day);
    onChange(newYmd);
    setTypedText(formatDateToDDMMYY(viewYear, viewMonth, day));
    setIsOpen(false);
  };

  const handlePrevMonth = () => {
    if (viewMonth === 1) {
      setViewMonth(12);
      setViewYear(prev => prev - 1);
    } else {
      setViewMonth(prev => prev - 1);
    }
  };

  const handleNextMonth = () => {
    if (viewMonth === 12) {
      setViewMonth(1);
      setViewYear(prev => prev + 1);
    } else {
      setViewMonth(prev => prev + 1);
    }
  };

  const handleSelectToday = () => {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth() + 1;
    const d = now.getDate();
    setViewYear(y);
    setViewMonth(m);
    onChange(formatDateToYMD(y, m, d));
    setTypedText(formatDateToDDMMYY(y, m, d));
    setIsOpen(false);
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setTypedText(e.target.value);
  };

  const handleInputBlur = () => {
    setIsInputFocused(false);
    const p = parseDatePickerParts(typedText);
    const ymd = formatDateToYMD(p.year, p.month, p.day);
    onChange(ymd);
    setTypedText(formatDateToDDMMYY(p.year, p.month, p.day));
    setViewYear(p.year);
    setViewMonth(p.month);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      (e.target as HTMLInputElement).blur();
    }
  };

  const daysInMonth = new Date(viewYear, viewMonth, 0).getDate();
  const firstDayIndex = new Date(viewYear, viewMonth - 1, 1).getDay();
  // Monday as starting day
  const startingOffset = (firstDayIndex + 6) % 7;

  const MONTH_NAMES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  const now = new Date();
  const isToday = (day: number) => 
    now.getFullYear() === viewYear && (now.getMonth() + 1) === viewMonth && now.getDate() === day;

  const isSelected = (day: number) =>
    selectedParts.year === viewYear && selectedParts.month === viewMonth && selectedParts.day === day;

  return (
    <div ref={containerRef} className="relative inline-block w-full">
      <div className="flex items-center relative">
        <input
          id={id}
          type="text"
          value={typedText}
          placeholder="dd-mm-yy"
          onFocus={() => {
            setIsInputFocused(true);
            setIsOpen(true);
          }}
          onChange={handleInputChange}
          onBlur={handleInputBlur}
          onKeyDown={handleKeyDown}
          className={`${className} w-full text-xs bg-white border border-[#E8ECF5] rounded-xl pl-8 pr-8 py-2 focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] outline-none font-mono font-bold text-[#252B3B] transition-all cursor-pointer`}
        />
        <Calendar 
          className="h-4 w-4 text-[#747D91] absolute left-2.5 pointer-events-none" 
        />
        <button
          type="button"
          onClick={() => setIsOpen(prev => !prev)}
          className="absolute right-2 p-1 text-[#747D91] hover:text-[#4D5A8F] transition-colors cursor-pointer"
          title="Toggle calendar"
        >
          <ChevronDown className={`h-3.5 w-3.5 transition-transform duration-150 ${isOpen ? 'rotate-180 text-[#4D5A8F]' : ''}`} />
        </button>
      </div>

      {isOpen && (
        <div className="absolute left-0 top-full mt-1.5 z-50 bg-white border border-[#E8ECF5] rounded-2xl shadow-xl p-3 w-72 animate-in fade-in zoom-in-95 duration-100">
          <div className="flex items-center justify-between mb-2 pb-2 border-b border-[#E8ECF5]">
            <button
              type="button"
              onClick={handlePrevMonth}
              className="p-1 rounded-lg hover:bg-[#F5F8FF] text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
              title="Previous Month"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
            <div className="text-xs font-extrabold text-[#252B3B] flex items-center gap-1">
              <span>{MONTH_NAMES[viewMonth - 1]}</span>
              <span className="text-[#747D91] font-mono">{viewYear}</span>
            </div>
            <button
              type="button"
              onClick={handleNextMonth}
              className="p-1 rounded-lg hover:bg-[#F5F8FF] text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
              title="Next Month"
            >
              <ChevronRight className="h-4 w-4" />
            </button>
          </div>

          <div className="grid grid-cols-7 gap-1 text-center mb-1">
            {['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su'].map(dayName => (
              <span key={dayName} className="text-[10px] font-bold text-[#747D91] uppercase py-0.5">
                {dayName}
              </span>
            ))}
          </div>

          <div className="grid grid-cols-7 gap-1 text-center">
            {Array.from({ length: startingOffset }).map((_, i) => (
              <div key={`offset-${i}`} className="h-7 w-7" />
            ))}
            {Array.from({ length: daysInMonth }).map((_, i) => {
              const day = i + 1;
              const selected = isSelected(day);
              const today = isToday(day);
              return (
                <button
                  key={day}
                  type="button"
                  onClick={() => handleSelectDay(day)}
                  className={`h-7 w-7 text-xs font-mono font-medium rounded-lg flex items-center justify-center transition-all cursor-pointer ${
                    selected
                      ? 'bg-[#4D5A8F] text-white font-extrabold shadow-xs'
                      : today
                      ? 'bg-[#E3EAFA] text-[#4D5A8F] font-bold hover:bg-[#CED8F2]'
                      : 'text-[#252B3B] hover:bg-[#F5F8FF]'
                  }`}
                >
                  {day}
                </button>
              );
            })}
          </div>

          <div className="flex items-center justify-between mt-3 pt-2 border-t border-[#E8ECF5]">
            <button
              type="button"
              onClick={handleSelectToday}
              className="text-[11px] font-bold text-[#4D5A8F] hover:text-[#252B3B] hover:underline cursor-pointer"
            >
              Today
            </button>
            <button
              type="button"
              onClick={() => setIsOpen(false)}
              className="text-[11px] font-medium text-[#747D91] hover:text-[#252B3B] cursor-pointer"
            >
              Close
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
    if (!saved) return INITIAL_CLASSES;
    try {
      const parsed: ClassSection[] = JSON.parse(saved);
      return parsed.map(c => {
        if (!c.schedule) {
          if (c.name === 'Class Section A') {
            return {
              ...c,
              schedule: [
                { id: 'slot-1', day: 'Sunday', startTime: '08:30 AM', endTime: '11:20 AM' },
                { id: 'slot-2', day: 'Tuesday', startTime: '01:00 PM', endTime: '02:00 PM' },
                { id: 'slot-3', day: 'Thursday', startTime: '09:00 AM', endTime: '10:30 AM' }
              ]
            };
          }
          if (c.name === 'Class Section B') {
            return {
              ...c,
              schedule: [
                { id: 'slot-4', day: 'Monday', startTime: '09:30 AM', endTime: '11:00 AM' },
                { id: 'slot-5', day: 'Wednesday', startTime: '01:00 PM', endTime: '03:00 PM' }
              ]
            };
          }
          return { ...c, schedule: [] };
        }
        // If Class Section A has old Monday slot, upgrade to Sunday 8:30-11:20
        if (c.name === 'Class Section A' && c.schedule && c.schedule.some(s => s.day === 'Monday' && s.startTime === '09:00 AM')) {
          return {
            ...c,
            schedule: [
              { id: 'slot-1', day: 'Sunday', startTime: '08:30 AM', endTime: '11:20 AM' },
              { id: 'slot-2', day: 'Tuesday', startTime: '01:00 PM', endTime: '02:00 PM' },
              { id: 'slot-3', day: 'Thursday', startTime: '09:00 AM', endTime: '10:30 AM' }
            ]
          };
        }
        return c;
      });
    } catch {
      return INITIAL_CLASSES;
    }
  });
  
  const setStudentHighlightColor = (studentId: string, className: string, colorId: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === className) {
        return {
          ...cl,
          students: cl.students.map(st => {
            if (st.id === studentId) {
              const updated = { ...st };
              if (!colorId || colorId === 'none') {
                delete updated.color;
              } else {
                updated.color = colorId;
              }
              return updated;
            }
            return st;
          })
        };
      }
      return cl;
    }));
  };
  
  const [activeClassName, setActiveClassName] = useState<string>('Class Section A');
  const [activeTab, setActiveTab] = useState<'students' | 'attendance' | 'grades' | 'timetable' | 'calendar'>('students');

  // Calendar & Agenda State
  const [calendarEvents, setCalendarEvents] = useState<CalendarEvent[]>(() => {
    try {
      const saved = localStorage.getItem('omni_calendar_events');
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.warn("Failed to read calendar events from localStorage", e);
    }
    return INITIAL_CALENDAR_EVENTS;
  });

  const [calendarViewDate, setCalendarViewDate] = useState<Date>(() => new Date());
  const [calendarSelectedDate, setCalendarSelectedDate] = useState<string>(getTodayYMD);
  const [calendarViewMode, setCalendarViewMode] = useState<'month' | 'agenda'>('month');
  const [calendarClassFilter, setCalendarClassFilter] = useState<string>('all');
  const [calendarTypeFilter, setCalendarTypeFilter] = useState<string>('all');

  // Event Creator / Editor Modal State
  const [isEventModalOpen, setIsEventModalOpen] = useState(false);
  const [editingEventId, setEditingEventId] = useState<string | null>(null);
  const [eventTitleInput, setEventTitleInput] = useState('');
  const [eventDateInput, setEventDateInput] = useState('');
  const [eventTimeInput, setEventTimeInput] = useState('10:00 AM');
  const [isEventAllDay, setIsEventAllDay] = useState(false);
  const [eventTypeInput, setEventTypeInput] = useState<CalendarEventType>('exam');
  const [eventClassInput, setEventClassInput] = useState<string>('all');
  const [eventDescriptionInput, setEventDescriptionInput] = useState('');
  const [eventModalError, setEventModalError] = useState<string | null>(null);
  const [confirmDeleteEventTarget, setConfirmDeleteEventTarget] = useState<CalendarEvent | null>(null);

  // Sidebar Controls (Responsive: closed by default on mobile phones/tablets)
  const [newClassNameInput, setNewClassNameInput] = useState('');
  const [isCreatingClass, setIsCreatingClass] = useState(false);
  const [confirmDeleteClass, setConfirmDeleteClass] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState<boolean>(() => {
    if (typeof window !== 'undefined' && window.innerWidth < 1024) {
      return false;
    }
    return true;
  });
  const [expandedSearchStudentId, setExpandedSearchStudentId] = useState<string | null>(null);

  // Responsive Tab Navigation Helper: Auto-closes mobile drawer when selecting destination
  const handleSelectTab = (tab: 'students' | 'attendance' | 'grades' | 'timetable' | 'calendar') => {
    setActiveTab(tab);
    if (typeof window !== 'undefined' && window.innerWidth < 1024) {
      setIsSidebarOpen(false);
    }
  };

  // Edit Class Name & Schedule Modal State
  const [editingClassTarget, setEditingClassTarget] = useState<{
    originalName: string;
    newName: string;
    schedule: ClassScheduleSlot[];
  } | null>(null);
  const [editingClassModalError, setEditingClassModalError] = useState<string | null>(null);
  const [editClassActiveTab, setEditClassActiveTab] = useState<'roster' | 'schedule'>('roster');
  const [editingClassAddSlotDays, setEditingClassAddSlotDays] = useState<DayOfWeek[]>(['Monday']);
  const [editingClassAddSlotStartHour, setEditingClassAddSlotStartHour] = useState('09:00');
  const [editingClassAddSlotStartPeriod, setEditingClassAddSlotStartPeriod] = useState<'AM' | 'PM'>('AM');
  const [editingClassAddSlotEndHour, setEditingClassAddSlotEndHour] = useState('10:30');
  const [editingClassAddSlotEndPeriod, setEditingClassAddSlotEndPeriod] = useState<'AM' | 'PM'>('AM');

  // Add Student to Class Modal State (opened via '+ Add a Student' button in Edit Class menu)
  const [isAddStudentModalOpen, setIsAddStudentModalOpen] = useState(false);
  const [addStudentIdInput, setAddStudentIdInput] = useState('');
  const [addStudentNameInput, setAddStudentNameInput] = useState('');
  const [addStudentModalError, setAddStudentModalError] = useState<string | null>(null);
  const [addStudentSuccessMsg, setAddStudentSuccessMsg] = useState<string | null>(null);

  // New Class Creation: Optional Schedule Slot Fields (Multi-day selection)
  const [newClassDays, setNewClassDays] = useState<DayOfWeek[]>([]);
  const [newClassStartHour, setNewClassStartHour] = useState('09:00');
  const [newClassStartPeriod, setNewClassStartPeriod] = useState<'AM' | 'PM'>('AM');
  const [newClassEndHour, setNewClassEndHour] = useState('10:30');
  const [newClassEndPeriod, setNewClassEndPeriod] = useState<'AM' | 'PM'>('AM');

  // Quick Add Timetable Slot Modal State (Multi-day selection)
  const [isAddScheduleModalOpen, setIsAddScheduleModalOpen] = useState(false);
  const [scheduleModalClass, setScheduleModalClass] = useState('');
  const [scheduleModalDays, setScheduleModalDays] = useState<DayOfWeek[]>(['Monday']);
  const [scheduleModalStartHour, setScheduleModalStartHour] = useState('09:00');
  const [scheduleModalStartPeriod, setScheduleModalStartPeriod] = useState<'AM' | 'PM'>('AM');
  const [scheduleModalEndHour, setScheduleModalEndHour] = useState('10:30');
  const [scheduleModalEndPeriod, setScheduleModalEndPeriod] = useState<'AM' | 'PM'>('AM');

  // Single Slot Time Modification Modal State (Pen icon in timetable bar)
  const [editingSlotTarget, setEditingSlotTarget] = useState<{
    className: string;
    slotId: string;
    day: DayOfWeek;
    startHour: string;
    startPeriod: 'AM' | 'PM';
    endHour: string;
    endPeriod: 'AM' | 'PM';
  } | null>(null);
  const [editingSlotError, setEditingSlotError] = useState<string | null>(null);

  // Student Gear Box & Warning States
  const [activeGearStudent, setActiveGearStudent] = useState<{ studentId: string; className: string } | null>(null);
  const [confirmDeleteStudentTarget, setConfirmDeleteStudentTarget] = useState<{ studentId: string; studentName: string; className: string } | null>(null);

  // Attendance state
  const [attendanceDate, setAttendanceDate] = useState<string>(getTodayYMD);
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

  // Staging state for list imports containing short ID values, attendance, grades, or multi-sections
  const [pendingImport, setPendingImport] = useState<{
    className: string;
    normal: Student[];
    short: Student[];
    attendanceDates?: string[];
    attendance?: Record<string, Record<string, 'Not Taken' | 'Attended' | 'Absent'>>;
    assessments?: Assessment[];
    grades?: Record<string, Record<string, number>>;
    isMultiSection?: boolean;
    allSections?: ParsedClassSectionData[];
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
        if (data && Array.isArray(data.calendarEvents)) {
          setCalendarEvents(data.calendarEvents);
        }
      } else {
        // First-time user: persist current client-side state into Firestore automatically
        try {
          const sanitizedClasses = JSON.parse(JSON.stringify(classes));
          const sanitizedEvents = JSON.parse(JSON.stringify(calendarEvents));
          setDoc(docRef, { classes: sanitizedClasses, calendarEvents: sanitizedEvents }, { merge: true })
            .then(() => {
              lastFirestoreDataRef.current = JSON.stringify(classes);
            })
            .catch((err) => {
              console.error("Failed to sync initial roster to Firestore:", err);
              setSyncErrorMessage("Failed to sync initial roster to Firestore.");
            });
        } catch (prepErr) {
          console.error("Failed to prepare initial roster for Firestore:", prepErr);
        }
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
      try {
        const sanitizedClasses = JSON.parse(currentClassesStr);
        setDoc(docRef, { classes: sanitizedClasses }, { merge: true })
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
      } catch (err) {
        setIsSyncing(false);
        console.error("Prep error: ", err);
      }
    }
  }, [classes, user]);

  // Sync calendar events to localStorage and Firestore
  useEffect(() => {
    try {
      localStorage.setItem('omni_calendar_events', JSON.stringify(calendarEvents));
    } catch (e) {
      console.warn("Could not save calendar events to localStorage", e);
    }

    if (!user || !isInitialDocFetched.current) return;
    const docRef = doc(db, 'users', user.uid);
    try {
      const sanitizedEvents = JSON.parse(JSON.stringify(calendarEvents));
      setDoc(docRef, { calendarEvents: sanitizedEvents }, { merge: true }).catch(err => {
        console.error("Firestore sync calendar events failed: ", err);
      });
    } catch (e) {
      console.error("Prep error for calendar events: ", e);
    }
  }, [calendarEvents, user]);

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
      // Explicitly prompt the user to select an existing browser account or sign in with another account
      googleProvider.setCustomParameters({
        prompt: 'select_account'
      });
      await signInWithPopup(auth, googleProvider);
    } catch (error: any) {
      console.error("Sign-in failed:", error);
      if (error.code !== 'auth/popup-closed-by-user') {
        setSyncErrorMessage(error.message || "Failed to sign in with Google.");
      }
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
  const handleScoreChangeCrossSection = (studentId: string, assessmentName: string, val: number | undefined, className: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === className) {
        const nextStudentScores = cl.grades[studentId] ? { ...cl.grades[studentId] } : {};
        if (val === undefined || isNaN(val)) {
          delete nextStudentScores[assessmentName];
        } else {
          const currentAss = cl.assessments.find(a => a.name === assessmentName);
          const maxAllowed = currentAss?.maxMark && currentAss.maxMark > 0 ? currentAss.maxMark : 100;
          let clamped = val;
          if (clamped < 0) clamped = 0;
          if (clamped > maxAllowed) clamped = maxAllowed;
          nextStudentScores[assessmentName] = clamped;
        }

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

    const scheduleSlots: ClassScheduleSlot[] = [];
    if (newClassDays.length > 0) {
      newClassDays.forEach((d, idx) => {
        scheduleSlots.push({
          id: `slot-${Date.now()}-${idx}`,
          day: d,
          startTime: `${newClassStartHour.trim() || '09:00'} ${newClassStartPeriod}`,
          endTime: `${newClassEndHour.trim() || '10:30'} ${newClassEndPeriod}`
        });
      });
    }

    const newClass: ClassSection = {
      name: formattedName,
      students: createdClassStudents,
      attendance: initialAttendance,
      assessments: [],
      grades: initialGrades,
      schedule: scheduleSlots
    };

    setClasses(prev => [...prev, newClass]);
    setActiveClassName(formattedName);
    
    // Reset states
    setNewClassNameInput('');
    setCreatedClassStudents([]);
    setCreatedClassFileError(null);
    setCreatedClassFileSuccess(null);
    setNewClassDays([]);
    setNewClassStartHour('09:00');
    setNewClassStartPeriod('AM');
    setNewClassEndHour('10:30');
    setNewClassEndPeriod('AM');
    setIsCreatingClass(false);
  };

  // Edit / Rename Class Section & Update its Schedule
  const handleSaveEditedClass = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingClassTarget) return;
    const trimmed = editingClassTarget.newName.trim();
    if (!trimmed) {
      setEditingClassModalError("Class section name cannot be empty.");
      return;
    }
    if (
      trimmed.toLowerCase() !== editingClassTarget.originalName.toLowerCase() &&
      classes.some(c => c.name.toLowerCase() === trimmed.toLowerCase())
    ) {
      setEditingClassModalError(`A class section named '${trimmed}' already exists.`);
      return;
    }

    const oldName = editingClassTarget.originalName;
    setClasses(prev => prev.map(cl => {
      if (cl.name === oldName) {
        return {
          ...cl,
          name: trimmed,
          schedule: editingClassTarget.schedule
        };
      }
      return cl;
    }));

    if (activeClassName === oldName) {
      setActiveClassName(trimmed);
    }
    setEditingClassTarget(null);
    setEditingClassModalError(null);
  };

  // Delete a specific schedule slot from a class
  const handleDeleteScheduleSlot = (className: string, slotIdOrKey: string) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === className && cl.schedule) {
        return {
          ...cl,
          schedule: cl.schedule.filter(s => (s.id || `${s.day}-${s.startTime}`) !== slotIdOrKey)
        };
      }
      return cl;
    }));
  };

  // Save changes to a single time slot (opened via pen icon on timetable bar)
  const handleSaveEditedSlot = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingSlotTarget) return;

    const startM = parseTimeToMinutes(`${editingSlotTarget.startHour} ${editingSlotTarget.startPeriod}`);
    const endM = parseTimeToMinutes(`${editingSlotTarget.endHour} ${editingSlotTarget.endPeriod}`);
    if (endM <= startM) {
      setEditingSlotError("Lecture end time must be after the start time.");
      return;
    }

    setClasses(prev => prev.map(cl => {
      if (cl.name === editingSlotTarget.className && cl.schedule) {
        return {
          ...cl,
          schedule: cl.schedule.map(s => {
            if ((s.id || `${s.day}-${s.startTime}`) === editingSlotTarget.slotId) {
              return {
                ...s,
                day: editingSlotTarget.day,
                startTime: `${editingSlotTarget.startHour.trim() || '09:00'} ${editingSlotTarget.startPeriod}`,
                endTime: `${editingSlotTarget.endHour.trim() || '10:30'} ${editingSlotTarget.endPeriod}`
              };
            }
            return s;
          })
        };
      }
      return cl;
    }));

    setEditingSlotTarget(null);
    setEditingSlotError(null);
  };

  // Add schedule slot(s) from Timetable modal (Supports selecting multiple days at once)
  const handleAddScheduleSlotSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const targetName = scheduleModalClass || activeClass.name;
    if (!targetName) return;

    const daysToSchedule = scheduleModalDays.length > 0 ? scheduleModalDays : (['Monday'] as DayOfWeek[]);
    const newSlots: ClassScheduleSlot[] = daysToSchedule.map((d, idx) => ({
      id: `slot-${Date.now()}-${idx}`,
      day: d,
      startTime: `${scheduleModalStartHour.trim() || '09:00'} ${scheduleModalStartPeriod}`,
      endTime: `${scheduleModalEndHour.trim() || '10:30'} ${scheduleModalEndPeriod}`
    }));

    setClasses(prev => prev.map(cl => {
      if (cl.name === targetName) {
        const currentSched = cl.schedule ? [...cl.schedule] : [];
        return {
          ...cl,
          schedule: [...currentSched, ...newSlots]
        };
      }
      return cl;
    }));

    setIsAddScheduleModalOpen(false);
  };

  // Clicking a class in Timetable jumps to that class and switches to students roster
  const handleOpenClassFromSchedule = (targetClassName: string) => {
    setActiveClassName(targetClassName);
    setActiveTab('students');
  };

  // Parsing helper for file upload inside class creation modal
  const handleModalRosterFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCreatedClassFileError(null);
    setCreatedClassFileSuccess(null);
    const file = e.target.files?.[0];
    if (!file) return;

    processModalRosterFile(file);
  };

  const processModalRosterFile = async (file: File) => {
    try {
      const result = await parseAnyExcelFile(file, newClassNameInput.trim() || 'New Class');
      const targetSec = result.sections[0];
      if (!targetSec || targetSec.students.length === 0) {
        setCreatedClassFileError("No student records found in file. Ensure file contains student IDs and names.");
        return;
      }

      setCreatedClassStudents(targetSec.students);
      setCreatedClassFileSuccess(`Successfully recognized ${targetSec.students.length} student records from spreadsheet.`);
    } catch (err: any) {
      setCreatedClassFileError(`Error parsing file: ${err?.message || 'Unsupported format'}`);
    }
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
  const handleScoreChange = (studentId: string, assessmentName: string, val: number | undefined) => {
    setClasses(prev => prev.map(cl => {
      if (cl.name === activeClass.name) {
        const nextStudentScores = cl.grades[studentId] ? { ...cl.grades[studentId] } : {};
        if (val === undefined || isNaN(val)) {
          delete nextStudentScores[assessmentName];
        } else {
          const currentAss = cl.assessments.find(a => a.name === assessmentName);
          const maxAllowed = currentAss?.maxMark && currentAss.maxMark > 0 ? currentAss.maxMark : 100;
          let clamped = val;
          if (clamped < 0) clamped = 0;
          if (clamped > maxAllowed) clamped = maxAllowed;
          nextStudentScores[assessmentName] = clamped;
        }

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

  const processRosterFile = async (file: File) => {
    try {
      const result = await parseAnyExcelFile(file, activeClass.name);

      if (result.isMultiSection) {
        // Multi-section file (like Merged Attendance or Merged Grades for All Classes)
        setPendingImport({
          className: activeClass.name,
          normal: result.sections.flatMap(s => s.students.filter(st => st.id.length > 6)),
          short: result.sections.flatMap(s => s.students.filter(st => st.id.length <= 6)),
          isMultiSection: true,
          allSections: result.sections
        });

        setSelectedShortIds(result.sections.flatMap(s => s.students.map(st => st.id)));
        setFileSuccess(`Detected multi-section workbook with ${result.sections.length} class sections and ${result.totalStudents} total students! Review below to import.`);
      } else {
        const sec = result.sections[0];
        const shortIds = sec.students.filter(st => st.id.length <= 6);
        const normalIds = sec.students.filter(st => st.id.length > 6);

        const matchingClass = classes.find(c => c.name.toLowerCase() === sec.className.toLowerCase());
        const targetClassName = matchingClass ? matchingClass.name : (classes.length > 0 ? activeClass.name : (sec.className || 'New Class Section'));

        setPendingImport({
          className: targetClassName,
          normal: normalIds,
          short: shortIds,
          attendanceDates: sec.attendanceDates,
          attendance: sec.attendance,
          assessments: sec.assessments,
          grades: sec.grades
        });

        setSelectedShortIds(sec.students.map(s => s.id));

        let successMsg = `Previewing imported spreadsheet containing ${sec.students.length} student records.`;
        if (sec.attendanceDates && sec.attendanceDates.length > 0) {
          successMsg += ` Includes ${sec.attendanceDates.length} attendance date logs.`;
        }
        if (sec.assessments && sec.assessments.length > 0) {
          successMsg += ` Includes ${sec.assessments.length} assessment columns.`;
        }
        setFileSuccess(successMsg);
      }
    } catch (err: any) {
      setFileError(`Error parsing spreadsheet: ${err?.message || 'Unsupported format'}`);
    }
  };

  // Commit Staging Roster with Checkbox Selection
  const handleConfirmPendingImport = () => {
    if (!pendingImport) return;

    if (pendingImport.isMultiSection && pendingImport.allSections) {
      // Multi-section merge into state
      setClasses(prev => {
        let updated = [...prev];
        pendingImport.allSections!.forEach(sec => {
          const existingIdx = updated.findIndex(c => c.name.toLowerCase() === sec.className.toLowerCase());
          if (existingIdx >= 0) {
            const existing = updated[existingIdx];
            // Merge students (preserve existing records / colors)
            const studentMap = new Map<string, Student>();
            existing.students.forEach(s => studentMap.set(s.id, s));
            sec.students.forEach(s => {
              if (studentMap.has(s.id)) {
                studentMap.set(s.id, { ...studentMap.get(s.id)!, ...s });
              } else {
                studentMap.set(s.id, s);
              }
            });
            const mergedStudents = Array.from(studentMap.values()).sort((a, b) => a.firstName.localeCompare(b.firstName));

            // Merge attendance dates and records
            const existingDates = existing.attendanceDates || [];
            const newDates = sec.attendanceDates || [];
            const mergedDates = Array.from(new Set([...newDates, ...existingDates]));

            const mergedAtt = { ...existing.attendance };
            if (sec.attendance) {
              Object.keys(sec.attendance).forEach(stId => {
                mergedAtt[stId] = { ...(mergedAtt[stId] || {}), ...(sec.attendance![stId] || {}) };
              });
            }

            // Merge assessments and grades
            const existingAssessments = existing.assessments || [];
            const newAssessments = sec.assessments || [];
            const assessMap = new Map<string, Assessment>();
            existingAssessments.forEach(a => assessMap.set(a.name.toLowerCase(), a));
            newAssessments.forEach(a => assessMap.set(a.name.toLowerCase(), a));
            const mergedAssessments = Array.from(assessMap.values());

            const mergedGrades = { ...existing.grades };
            if (sec.grades) {
              Object.keys(sec.grades).forEach(stId => {
                mergedGrades[stId] = { ...(mergedGrades[stId] || {}), ...(sec.grades![stId] || {}) };
              });
            }

            updated[existingIdx] = {
              ...existing,
              students: mergedStudents,
              attendanceDates: mergedDates.length > 0 ? mergedDates : existing.attendanceDates,
              attendance: mergedAtt,
              assessments: mergedAssessments.length > 0 ? mergedAssessments : existing.assessments,
              grades: mergedGrades
            };
          } else {
            // New class section from file
            updated.push({
              name: sec.className,
              students: sec.students,
              attendanceDates: sec.attendanceDates,
              attendance: sec.attendance,
              assessments: sec.assessments,
              grades: sec.grades,
              schedule: []
            });
          }
        });
        return updated;
      });

      setPendingImport(null);
      setFileSuccess(`Successfully imported ${pendingImport.allSections.length} class sections from spreadsheet!`);
      return;
    }

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

    setClasses(prev => {
      const matchIdx = prev.findIndex(c => c.name.toLowerCase() === pendingImport.className.toLowerCase());
      if (matchIdx >= 0) {
        return prev.map((cl, idx) => {
          if (idx !== matchIdx) return cl;
          // Attendance records
          const nextAtt: any = { ...cl.attendance };
          consolidatedList.forEach(st => {
            if (!nextAtt[st.id]) nextAtt[st.id] = cl.attendance[st.id] || {};
            if (pendingImport.attendance && pendingImport.attendance[st.id]) {
              nextAtt[st.id] = { ...nextAtt[st.id], ...pendingImport.attendance[st.id] };
            }
          });

          // Attendance Dates
          const existingDates = cl.attendanceDates || [];
          const importedDates = pendingImport.attendanceDates || [];
          const mergedDates = Array.from(new Set([...importedDates, ...existingDates]));

          // Grades
          const nextGrades: any = { ...cl.grades };
          consolidatedList.forEach(st => {
            if (!nextGrades[st.id]) nextGrades[st.id] = cl.grades[st.id] || {};
            if (pendingImport.grades && pendingImport.grades[st.id]) {
              nextGrades[st.id] = { ...nextGrades[st.id], ...pendingImport.grades[st.id] };
            }
          });

          // Assessments
          const existingAssess = cl.assessments || [];
          const importedAssess = pendingImport.assessments || [];
          const assessMap = new Map<string, Assessment>();
          existingAssess.forEach(a => assessMap.set(a.name.toLowerCase(), a));
          importedAssess.forEach(a => assessMap.set(a.name.toLowerCase(), a));
          const mergedAssess = Array.from(assessMap.values());

          return {
            ...cl,
            students: consolidatedList,
            attendanceDates: mergedDates.length > 0 ? mergedDates : cl.attendanceDates,
            attendance: nextAtt,
            assessments: mergedAssess.length > 0 ? mergedAssess : cl.assessments,
            grades: nextGrades
          };
        });
      } else {
        return [
          ...prev,
          {
            name: pendingImport.className,
            students: consolidatedList,
            attendanceDates: pendingImport.attendanceDates || [],
            attendance: pendingImport.attendance || {},
            assessments: pendingImport.assessments || [],
            grades: pendingImport.grades || {},
            schedule: []
          }
        ];
      }
    });

    if (classes.length === 0 || !activeClassName) {
      setActiveClassName(pendingImport.className);
    }

    setPendingImport(null);
    let msg = `Successfully imported roster containing ${consolidatedList.length} records.`;
    if (pendingImport.attendanceDates && pendingImport.attendanceDates.length > 0) {
      msg += ` Restored ${pendingImport.attendanceDates.length} attendance date logs.`;
    }
    if (pendingImport.assessments && pendingImport.assessments.length > 0) {
      msg += ` Restored ${pendingImport.assessments.length} assessment columns.`;
    }
    setFileSuccess(msg);
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

  // Add a new student to the class currently being edited in the Edit Class Section modal
  const handleAddStudentToClass = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!editingClassTarget) return;

    const id = addStudentIdInput.trim();
    const name = addStudentNameInput.trim();

    if (!id) {
      setAddStudentModalError("Please enter a Student ID.");
      return;
    }
    if (!name) {
      setAddStudentModalError("Please enter the student's name.");
      return;
    }

    const targetClass = classes.find(c => c.name === editingClassTarget.originalName);
    if (!targetClass) {
      setAddStudentModalError("Class section not found.");
      return;
    }

    if (targetClass.students.some(s => s.id.toLowerCase() === id.toLowerCase())) {
      setAddStudentModalError(`A student with ID '${id}' is already enrolled in this class.`);
      return;
    }

    const parts = name.split(/\s+/);
    const firstName = parts[0] || name;
    const lastName = parts.slice(1).join(' ') || '';

    const newStudent: Student = {
      id,
      firstName,
      lastName,
      fullName: name
    };

    const updatedStudents = [...targetClass.students, newStudent].sort((a, b) => 
      a.firstName.localeCompare(b.firstName)
    );

    const dates = getAttendanceDates(targetClass);
    const studentAtt: Record<string, 'Not Taken' | 'Attended' | 'Absent'> = {};
    dates.forEach(d => {
      studentAtt[d] = 'Not Taken';
    });

    const studentGrades: Record<string, number> = {};
    (targetClass.assessments || []).forEach(ass => {
      studentGrades[ass.name] = 0;
    });

    setClasses(prev => prev.map(cl => {
      if (cl.name === editingClassTarget.originalName) {
        return {
          ...cl,
          students: updatedStudents,
          attendance: {
            ...cl.attendance,
            [id]: studentAtt
          },
          grades: {
            ...cl.grades,
            [id]: studentGrades
          }
        };
      }
      return cl;
    }));

    setAddStudentIdInput('');
    setAddStudentNameInput('');
    setAddStudentModalError(null);
    setIsAddStudentModalOpen(false);
    setAddStudentSuccessMsg(`Student '${name}' (${id}) was successfully added!`);
    setTimeout(() => setAddStudentSuccessMsg(null), 3500);
  };

  // Single Class Exporters with Auto Column Width Formatting
  const handleExportFullClass = (targetClass: ClassSection = activeClass) => {
    if (!targetClass || targetClass.students.length === 0) {
      alert("No student records found in this class section.");
      return;
    }

    const dates = getAttendanceDates(targetClass);
    const assessCols = (targetClass.assessments || []).map(a => `${a.name} (Max ${a.maxMark || 100}, ${a.weight}%)`);
    const headers = [
      'ID Number',
      'Full Name',
      ...assessCols,
      'Final Grade',
      '', // Distinction separator column between grades and attendance
      'Attendance Summary',
      ...dates.map(d => formatDisplayDate(d))
    ];

    const aoa: any[][] = [];
    // 1. Add Class Section banner so importer immediately recognizes the exact class
    aoa.push([`🏫 ${targetClass.name}`]);
    // 2. Add header row
    aoa.push(headers);

    // 3. Add student data rows
    targetClass.students.forEach(st => {
      const row: any[] = [st.id, st.fullName];
      
      // Grade values closer to student names
      (targetClass.assessments || []).forEach(ass => {
        row.push(targetClass.grades[st.id]?.[ass.name] ?? 0);
      });

      // Final Grade
      row.push(calculateFinalGrade(st.id, targetClass));

      // Empty separator column
      row.push('');

      // Attendance summary showing how many lectures attended, e.g. (2/6 lectures 33%)
      const totalDates = dates.length;
      const attendedCount = dates.filter(d => targetClass.attendance[st.id]?.[d] === 'Attended').length;
      const pct = totalDates > 0 ? Math.round((attendedCount / totalDates) * 100) : 0;
      row.push(`(${attendedCount}/${totalDates} lectures ${pct}%)`);

      // Attendance date values
      dates.forEach(d => {
        row.push(targetClass.attendance[st.id]?.[d] ?? 'Not Taken');
      });

      aoa.push(row);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    applyAutoColumnWidths(worksheet, aoa);
    const workbook = XLSX.utils.book_new();
    const safeSheetName = (targetClass.name || 'Full Class Info').replace(/[:\\/?*\[\]]/g, '').slice(0, 31);
    XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName || 'Full Class Info');

    XLSX.writeFile(workbook, `Full_Info_${targetClass.name.replace(/\s+/g, '_')}.xlsx`);
  };

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
    applyAutoColumnWidthsToTable(worksheet, tableData);
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
    applyAutoColumnWidthsToTable(worksheet, tableData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Attendance Ledger');

    XLSX.writeFile(workbook, `Attendance_Log_${activeClass.name.replace(/\s+/g, '_')}.xlsx`);
  };

  // GLOBAL MULTI-CLASS MERGED DOWNLOADING SYSTEM
  // Downloads structured sheets with auto-formatted column widths and metadata
  const handleExportMerged = (type: 'grades' | 'attendance' | 'full') => {
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
      } else if (type === 'grades') {
        // Grades Roster with max mark and weight metadata
        const assessCols = cl.assessments.map(a => `${a.name} (Max ${a.maxMark || 100}, ${a.weight}%)`);
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
      } else {
        // FULL: Grades closer to names + empty separator column + Attendance Summary (e.g. (2/6 lectures 33%)) + Attendance dates
        const dates = getAttendanceDates(cl);
        const assessCols = (cl.assessments || []).map(a => `${a.name} (Max ${a.maxMark || 100}, ${a.weight}%)`);
        const headers = [
          'ID Number',
          'Full Name',
          ...assessCols,
          'Final Grade',
          '', // Distinction separator column between grades and attendance
          'Attendance Summary',
          ...dates.map(d => formatDisplayDate(d))
        ];
        aoa.push(headers);

        cl.students.forEach(st => {
          const row: any[] = [st.id, st.fullName];

          // Grades closer to names
          (cl.assessments || []).forEach(ass => {
            row.push(cl.grades[st.id]?.[ass.name] ?? 0);
          });

          // Final Grade
          row.push(calculateFinalGrade(st.id, cl));

          // Empty separator column
          row.push('');

          // Attendance summary showing how many lectures attended, e.g. (2/6 lectures 33%)
          const totalDates = dates.length;
          const attendedCount = dates.filter(d => cl.attendance[st.id]?.[d] === 'Attended').length;
          const pct = totalDates > 0 ? Math.round((attendedCount / totalDates) * 100) : 0;
          row.push(`(${attendedCount}/${totalDates} lectures ${pct}%)`);

          // Individual attendance date columns
          dates.forEach(d => {
            row.push(cl.attendance[st.id]?.[d] ?? 'Not Taken');
          });

          aoa.push(row);
        });
      }

      // 3. Write exactly 2 lines of empty cell structures between sections
      aoa.push([]);
      aoa.push([]);
    });

    if (aoa.length === 0) {
      alert("No rosters containing student listings found across sections.");
      return;
    }

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    applyAutoColumnWidths(worksheet, aoa);
    const workbook = XLSX.utils.book_new();
    const sheetName = type === 'full' ? 'Merged_Full_Info' : `Merged_${type.charAt(0).toUpperCase() + type.slice(1)}`;
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    const fileName = type === 'full'
      ? 'Merged_Full_Info_All_Classes.xlsx'
      : `Merged_${type.charAt(0).toUpperCase() + type.slice(1)}_All_Classes.xlsx`;

    XLSX.writeFile(workbook, fileName);
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

  // Teacher Weekly Timetable Component Renderer
  const renderTeacherWeeklyTimetable = () => {
    // Earliest possible lecture is 8:00 AM (480 minutes)
    const TIMELINE_START_HOUR = 8;
    const TIMELINE_START_MIN = TIMELINE_START_HOUR * 60;

    // Dynamically calculate the latest end time (at least 6:00 PM / 18:00)
    let maxEndMin = 18 * 60;
    classes.forEach(cl => {
      (cl.schedule || []).forEach(s => {
        const e = parseTimeToMinutes(s.endTime);
        if (e > maxEndMin) maxEndMin = e;
      });
    });
    const TIMELINE_END_HOUR = Math.max(18, Math.ceil(maxEndMin / 60));
    const TOTAL_HOURS = TIMELINE_END_HOUR - TIMELINE_START_HOUR;
    const TOTAL_MINUTES = TOTAL_HOURS * 60;

    // Generate hour markers starting at 8:00 AM
    const hourMarkers: { hour: number; label: string; period: string }[] = [];
    for (let h = TIMELINE_START_HOUR; h <= TIMELINE_END_HOUR; h++) {
      const period = h >= 12 ? 'PM' : 'AM';
      const displayH = h % 12 === 0 ? 12 : h % 12;
      const label = `${String(displayH).padStart(2, '0')}:00`;
      hourMarkers.push({ hour: h, label, period });
    }

    // Check current day and time
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const now = new Date();
    const todayName = dayNames[now.getDay()];
    const nowMinutes = now.getHours() * 60 + now.getMinutes();
    const isNowInTimeline = nowMinutes >= TIMELINE_START_MIN && nowMinutes <= TIMELINE_END_HOUR * 60;
    const nowPercent = isNowInTimeline ? ((nowMinutes - TIMELINE_START_MIN) / TOTAL_MINUTES) * 100 : -1;

    let totalLecturesAcrossWeek = 0;
    classes.forEach(cl => {
      totalLecturesAcrossWeek += (cl.schedule || []).length;
    });

    return (
      <div className="bg-white rounded-2xl border border-[#E8ECF5] p-6 shadow-xs flex flex-col gap-6" id="teacher-timetable-viewport">
        {/* Header */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-3 border-b border-[#E8ECF5]">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-[#E3EAFA] text-[#4D5A8F] rounded-2xl border border-[#B3C0E6]/60 shadow-3xs">
              <Clock className="h-6 w-6 text-[#4D5A8F]" />
            </div>
            <div>
              <h3 className="font-extrabold text-[#252B3B] text-lg tracking-tight">
                Teacher Weekly Lecture Timetable
              </h3>
              <p className="text-xs text-[#747D91] mt-0.5">
                Weekly distribution of lectures across all class sections. Click on any class to jump straight to its roster &amp; sneak peek.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => {
                setScheduleModalClass(activeClass.name || classes[0]?.name || '');
                setScheduleModalDays(['Monday']);
                setScheduleModalStartHour('09:00');
                setScheduleModalStartPeriod('AM');
                setScheduleModalEndHour('10:30');
                setScheduleModalEndPeriod('AM');
                setIsAddScheduleModalOpen(true);
              }}
              disabled={classes.length === 0}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-bold text-white bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-95 rounded-xl shadow-xs transition-all cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed border border-[#4D5A8F]"
              id="add-timetable-slot-btn"
            >
              <Plus className="h-4 w-4" />
              <span>Add Lecture Slot</span>
            </button>
          </div>
        </div>

        {/* Class Section Color Legend & Quick Navigation */}
        <div className="flex items-center gap-2 flex-wrap text-xs bg-[#F5F8FF] p-3.5 rounded-2xl border border-[#E8ECF5]">
          <span className="text-[10px] font-black text-[#747D91] uppercase tracking-wider mr-1">
            Active Sections:
          </span>
          {classes.length === 0 ? (
            <span className="text-[#747D91] italic text-xs">No active class sections.</span>
          ) : (
            classes.map((cl, cIdx) => {
              const palette = CLASS_COLOR_PALETTES[cIdx % CLASS_COLOR_PALETTES.length];
              const slotCount = (cl.schedule || []).length;
              return (
                <button
                  key={cl.name}
                  type="button"
                  onClick={() => handleOpenClassFromSchedule(cl.name)}
                  className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-xl border text-xs font-bold cursor-pointer transition-all hover:scale-105 active:scale-95 shadow-3xs ${palette.bg} ${palette.border} ${palette.text}`}
                  title={`Click to open ${cl.name} roster and sneak peek`}
                >
                  <span>{cl.name}</span>
                  <span className={`text-[10px] px-1.5 py-0.2 rounded-md font-mono ${palette.badge}`}>
                    {slotCount} {slotCount === 1 ? 'slot' : 'slots'}
                  </span>
                </button>
              );
            })
          )}
        </div>

        {/* Weekly Timetable Table Matrix with Proportionate Duration Bars */}
        {classes.length === 0 ? (
          <div className="py-16 text-center text-[#747D91] italic bg-[#F5F8FF] border border-dashed border-[#E8ECF5] rounded-2xl">
            No active class sections created yet. Create a class section to view its weekly schedule.
          </div>
        ) : (
          <div className="flex flex-col gap-3">
            <div className="overflow-x-auto border border-[#E8ECF5] rounded-2xl shadow-xs bg-white" id="timetable-proportional-timeline">
              <div className="min-w-[820px] sm:min-w-[1080px] flex flex-col">
                
                {/* Timeline Header: First Column = Day of Week, Rest = Hours from 8:00 AM */}
                <div className="flex items-stretch border-b border-[#E8ECF5] bg-[#E3EAFA] text-[#252B3B] select-none sticky top-0 z-20">
                  {/* First Column Header: Days */}
                  <div className="w-28 sm:w-52 shrink-0 px-2.5 sm:px-4 py-3 sm:py-3.5 border-r border-[#E8ECF5] sticky left-0 z-30 bg-[#E3EAFA] flex items-center justify-between shadow-[2px_0_4px_-2px_rgba(0,0,0,0.05)]">
                    <div className="flex flex-col min-w-0">
                      <span className="text-[11px] font-black uppercase tracking-wider text-[#252B3B] truncate">
                        Day
                      </span>
                      <span className="text-[9.5px] text-[#747D91] font-medium hidden sm:inline">
                        Sun → Sat
                      </span>
                    </div>
                    <span className="text-[9.5px] sm:text-[10px] font-mono font-bold text-white bg-[#4D5A8F] px-1.5 sm:px-2 py-0.5 rounded-md shadow-3xs">
                      {totalLecturesAcrossWeek} {totalLecturesAcrossWeek === 1 ? 'Slot' : 'Slots'}
                    </span>
                  </div>

                  {/* Timeline Hour Markers starting at 8:00 AM */}
                  <div className="flex-1 relative h-13 bg-[#E3EAFA]/95">
                    {hourMarkers.map((hm, idx) => {
                      const leftPct = ((hm.hour - TIMELINE_START_HOUR) / TOTAL_HOURS) * 100;
                      const isLast = idx === hourMarkers.length - 1;
                      return (
                        <div
                          key={hm.hour}
                          className="absolute top-0 bottom-0 flex flex-col justify-between"
                          style={{
                            left: `${leftPct}%`,
                            transform: isLast ? 'translateX(-100%)' : idx === 0 ? 'translateX(0)' : 'translateX(-50%)'
                          }}
                        >
                          <div className={`pt-2 flex flex-col ${isLast ? 'items-end pr-2' : idx === 0 ? 'items-start pl-2' : 'items-center'}`}>
                            <span className="font-mono text-[11px] font-black text-[#252B3B] whitespace-nowrap">
                              {hm.label} <span className="text-[9.5px] font-extrabold text-[#4D5A8F]">{hm.period}</span>
                            </span>
                            <span className="text-[9px] font-mono text-[#747D91] -mt-0.5">
                              {hm.hour >= 12 ? 'Afternoon' : 'Morning'}
                            </span>
                          </div>
                          <div className={`h-2.5 w-px bg-[#8F9DCC] ${isLast ? 'self-end mr-0' : idx === 0 ? 'self-start ml-0' : 'self-center'}`} />
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Timeline Body: 7 Weekday Rows starting with Sunday and ending with Saturday */}
                <div className="divide-y divide-[#E8ECF5]">
                  {DAYS_OF_WEEK.map(day => {
                    // Gather all lectures scheduled on this day
                    const dayLectures: Array<{
                      className: string;
                      classSection: ClassSection;
                      slot: ClassScheduleSlot;
                      palette: typeof CLASS_COLOR_PALETTES[0];
                      startMin: number;
                      endMin: number;
                      durationMin: number;
                    }> = [];

                    classes.forEach((cl, cIdx) => {
                      const palette = CLASS_COLOR_PALETTES[cIdx % CLASS_COLOR_PALETTES.length];
                      (cl.schedule || []).forEach(s => {
                        if (s.day === day) {
                          const startMin = parseTimeToMinutes(s.startTime);
                          let endMin = parseTimeToMinutes(s.endTime);
                          if (endMin <= startMin) endMin = startMin + 60; // fallback if invalid
                          const durationMin = Math.max(15, endMin - startMin);
                          dayLectures.push({
                            className: cl.name,
                            classSection: cl,
                            slot: s,
                            palette,
                            startMin,
                            endMin,
                            durationMin
                          });
                        }
                      });
                    });

                    // Sort lectures chronologically by start time
                    dayLectures.sort((a, b) => a.startMin - b.startMin);

                    // Non-colliding lane allocation so overlapping classes stack neatly
                    const lanesEnd: number[] = [];
                    const positionedBars = dayLectures.map(lec => {
                      let assignedLane = -1;
                      for (let l = 0; l < lanesEnd.length; l++) {
                        if (lanesEnd[l] <= lec.startMin) {
                          assignedLane = l;
                          lanesEnd[l] = lec.endMin;
                          break;
                        }
                      }
                      if (assignedLane === -1) {
                        assignedLane = lanesEnd.length;
                        lanesEnd.push(lec.endMin);
                      }

                      // Proportional position and width relative to 8:00 AM timeline
                      const leftPct = Math.max(0, Math.min(100, ((lec.startMin - TIMELINE_START_MIN) / TOTAL_MINUTES) * 100));
                      const widthPct = Math.max(3.5, Math.min(100 - leftPct, (lec.durationMin / TOTAL_MINUTES) * 100));

                      return {
                        ...lec,
                        lane: assignedLane,
                        leftPercent: leftPct,
                        widthPercent: widthPct
                      };
                    });

                    const totalLanes = Math.max(1, lanesEnd.length);
                    const rowHeight = totalLanes === 1 ? 84 : totalLanes * 66 + 18;
                    const isToday = day === todayName;

                    return (
                      <div key={day} className={`flex items-stretch transition-colors ${isToday ? 'bg-[#E3EAFA]/30' : 'hover:bg-[#F5F8FF]/80'}`}>
                        {/* FIRST COLUMN: Weekday Name, Lecture Count, and Quick Add Action */}
                        <div 
                          className={`w-28 sm:w-52 shrink-0 p-2 sm:p-3.5 border-r border-[#E8ECF5] sticky left-0 z-10 flex flex-col justify-between transition-colors shadow-[2px_0_4px_-2px_rgba(0,0,0,0.05)] ${
                            isToday ? 'bg-[#E3EAFA] border-r-2 border-r-[#4D5A8F]' : 'bg-[#F5F8FF]'
                          }`}
                          style={{ height: `${rowHeight}px` }}
                        >
                          <div className="flex items-start justify-between gap-1">
                            <div className="flex items-center gap-1 min-w-0">
                              <span className={`text-xs sm:text-sm font-black tracking-tight truncate ${isToday ? 'text-[#4D5A8F]' : 'text-[#252B3B]'}`}>
                                <span className="sm:hidden">{day.slice(0, 3)}</span>
                                <span className="hidden sm:inline">{day}</span>
                              </span>
                              {isToday && (
                                <span className="px-1 py-0.2 rounded-full text-[8.5px] sm:text-[9px] font-black uppercase tracking-wider bg-[#4D5A8F] text-white shadow-3xs">
                                  Today
                                </span>
                              )}
                            </div>
                            
                            <button
                              type="button"
                              onClick={() => {
                                setScheduleModalClass(activeClass.name || classes[0]?.name || '');
                                setScheduleModalDays([day]);
                                setScheduleModalStartHour('09:00');
                                setScheduleModalStartPeriod('AM');
                                setScheduleModalEndHour('10:30');
                                setScheduleModalEndPeriod('AM');
                                setIsAddScheduleModalOpen(true);
                              }}
                              className="p-1 rounded-lg text-[#747D91] hover:text-[#4D5A8F] hover:bg-[#E3EAFA] border border-transparent hover:border-[#B3C0E6] transition-all cursor-pointer shadow-none hover:shadow-3xs active:scale-95 shrink-0"
                              title={`Add lecture slot on ${day}`}
                            >
                              <Plus className="h-3.5 w-3.5" />
                            </button>
                          </div>

                          <div className="flex items-center justify-between text-[10.5px]">
                            {dayLectures.length > 0 ? (
                              <span className="font-mono font-bold text-[#252B3B] bg-white border border-[#E8ECF5] px-2 py-0.5 rounded-md shadow-3xs">
                                {dayLectures.length} {dayLectures.length === 1 ? 'lecture' : 'lectures'}
                              </span>
                            ) : (
                              <span className="text-[#747D91] font-medium italic">
                                No lectures
                              </span>
                            )}
                            {dayLectures.length > 0 && (
                              <span className="text-[10px] font-mono text-[#747D91] font-semibold" title="Total lecture hours today">
                                {formatDurationMinutes(dayLectures.reduce((s, l) => s + l.durationMin, 0))}
                              </span>
                            )}
                          </div>
                        </div>

                        {/* SECOND COLUMN: Continuous Horizontal Timeline Track with Proportionate Duration Bars */}
                        <div 
                          className="flex-1 relative bg-white group/track"
                          style={{ height: `${rowHeight}px` }}
                        >
                          {/* Vertical Guideline Lines for Every Hour */}
                          {hourMarkers.map((hm) => {
                            const lineLeft = ((hm.hour - TIMELINE_START_HOUR) / TOTAL_HOURS) * 100;
                            return (
                              <div
                                key={hm.hour}
                                className="absolute top-0 bottom-0 pointer-events-none border-r border-dashed border-[#E8ECF5]"
                                style={{ left: `${lineLeft}%` }}
                              />
                            );
                          })}

                          {/* Real-time Indicator Line if today and within timeline */}
                          {isToday && nowPercent >= 0 && (
                            <div 
                              className="absolute top-0 bottom-0 z-20 pointer-events-none flex flex-col items-center"
                              style={{ left: `${nowPercent}%` }}
                              title={`Current Time: ${now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`}
                            >
                              <div className="w-2.5 h-2.5 rounded-full bg-rose-500 shadow-xs -mt-1 ring-2 ring-white" />
                              <div className="w-0.5 h-full bg-rose-500/80 shadow-xs" />
                            </div>
                          )}

                          {/* Proportionate Duration Class Bars */}
                          {positionedBars.map((bar, bIdx) => {
                            const topPx = bar.lane * 64 + 10;
                            return (
                              <div
                                key={bIdx}
                                onClick={() => handleOpenClassFromSchedule(bar.className)}
                                className={`absolute rounded-xl border ${bar.palette.bg} ${bar.palette.border} ${bar.palette.hover} cursor-pointer transition-all shadow-xs hover:shadow-md hover:scale-[1.01] hover:z-30 flex flex-col justify-between p-2 group/bar overflow-hidden select-none`}
                                style={{
                                  left: `${bar.leftPercent}%`,
                                  width: `${bar.widthPercent}%`,
                                  top: `${topPx}px`,
                                  height: '60px'
                                }}
                                title={`Click to open ${bar.className} roster and sneak peek (${bar.slot.startTime} – ${bar.slot.endTime}, duration ${formatDurationMinutes(bar.durationMin)})`}
                              >
                                {/* Top Line: Class Name, Pen Icon, Duration Badge, and Quick Delete Slot Button */}
                                <div className="flex items-center justify-between gap-1 min-w-0">
                                  <div className="flex items-center gap-1.5 min-w-0">
                                    <span className={`w-2 h-2 rounded-full ${bar.palette.dot} shrink-0`} />
                                    <span className={`text-xs font-black ${bar.palette.text} truncate`}>
                                      {bar.className}
                                    </span>
                                    {/* Pen Icon next to class name to modify slot time */}
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        const partsStart = (bar.slot.startTime || '09:00 AM').trim().split(' ');
                                        const partsEnd = (bar.slot.endTime || '10:30 AM').trim().split(' ');
                                        setEditingSlotTarget({
                                          className: bar.className,
                                          slotId: bar.slot.id || `${bar.slot.day}-${bar.slot.startTime}`,
                                          day: bar.slot.day,
                                          startHour: partsStart[0] || '09:00',
                                          startPeriod: (partsStart[1] as 'AM' | 'PM') || 'AM',
                                          endHour: partsEnd[0] || '10:30',
                                          endPeriod: (partsEnd[1] as 'AM' | 'PM') || 'AM',
                                        });
                                        setEditingSlotError(null);
                                      }}
                                      className="p-1 rounded-md text-[#747D91] hover:text-[#4D5A8F] hover:bg-white/90 transition-all cursor-pointer shrink-0 shadow-none hover:shadow-3xs"
                                      title={`Modify lecture time for ${bar.className} on ${bar.slot.day}`}
                                    >
                                      <Pencil className="h-3 w-3 text-[#747D91] hover:text-[#4D5A8F]" />
                                    </button>
                                    <span className={`text-[9.5px] font-mono font-bold px-1.5 py-0.2 rounded-md ${bar.palette.badge} shrink-0`}>
                                      {formatDurationMinutes(bar.durationMin)}
                                    </span>
                                  </div>

                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleDeleteScheduleSlot(bar.className, bar.slot.id || `${bar.slot.day}-${bar.slot.startTime}`);
                                    }}
                                    className="p-1 text-[#747D91]/60 hover:text-rose-600 hover:bg-white/90 rounded-md transition-colors opacity-0 group-hover/bar:opacity-100 cursor-pointer shrink-0"
                                    title="Delete this lecture slot"
                                  >
                                    <Trash2 className="h-3 w-3" />
                                  </button>
                                </div>

                                {/* Bottom Line: Time Range & Student Count Navigation */}
                                <div className="flex items-center justify-between gap-1 text-[10px] text-[#747D91] pt-0.5">
                                  <div className="inline-flex items-center gap-1 font-mono font-bold text-[#252B3B] bg-white/90 px-1.5 py-0.5 rounded border border-[#E8ECF5] truncate shadow-3xs">
                                    <Clock className="h-2.5 w-2.5 text-[#4D5A8F] shrink-0" />
                                    <span className="truncate">{bar.slot.startTime} – {bar.slot.endTime}</span>
                                  </div>
                                  
                                  <div className="flex items-center gap-1 shrink-0 text-[#747D91] font-medium">
                                    <span className="hidden sm:inline">👥 {bar.classSection.students.length}</span>
                                    <span className="font-bold text-[#4D5A8F] group-hover/bar:underline ml-0.5 flex items-center">
                                      Open →
                                    </span>
                                  </div>
                                </div>
                              </div>
                            );
                          })}

                          {/* Empty Day Prompt: Display friendly add button on hover */}
                          {positionedBars.length === 0 && (
                            <div className="absolute inset-0 flex items-center justify-center opacity-0 group-hover/track:opacity-100 transition-opacity bg-slate-50/40">
                              <button
                                type="button"
                                onClick={() => {
                                  setScheduleModalClass(activeClass.name || classes[0]?.name || '');
                                  setScheduleModalDays([day]);
                                  setScheduleModalStartHour('09:00');
                                  setScheduleModalStartPeriod('AM');
                                  setScheduleModalEndHour('10:30');
                                  setScheduleModalEndPeriod('AM');
                                  setIsAddScheduleModalOpen(true);
                                }}
                                className="text-xs font-bold text-[#52796F] bg-white border border-[#CAD2C5] hover:bg-[#E8EEE7] px-3 py-1.5 rounded-xl shadow-3xs cursor-pointer flex items-center gap-1.5 active:scale-95 transition-all"
                              >
                                <Plus className="h-3.5 w-3.5" />
                                <span>Add lecture slot for {day}</span>
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Timetable Sub-footer / Legend Bar */}
            <div className="flex items-center justify-between flex-wrap gap-2 text-xs text-[#354F52]/80 bg-[#F6F8F5] p-3 rounded-2xl border border-[#CAD2C5]/50">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="font-bold text-[#2F3E46]">⏱️ Proportional Timeline:</span>
                <span className="text-[#354F52]/80">
                  Day begins at <strong>08:00 AM</strong>. Lecture bar widths are sized proportionally to class duration (e.g., 2h 50m bar is visibly wider than 1h bar).
                </span>
              </div>
              <div className="flex items-center gap-3 text-[11px] font-mono text-[#354F52]/60">
                <span>Days: Sun → Sat</span>
                <span>•</span>
                <span>Click any class bar to jump to student sneak peak</span>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  // Calendar & Agenda System Handlers & Renderer
  const handlePrevCalendarMonth = () => {
    setCalendarViewDate(prev => new Date(prev.getFullYear(), prev.getMonth() - 1, 1));
  };

  const handleNextCalendarMonth = () => {
    setCalendarViewDate(prev => new Date(prev.getFullYear(), prev.getMonth() + 1, 1));
  };

  const handleJumpToToday = () => {
    const now = new Date();
    setCalendarViewDate(new Date(now.getFullYear(), now.getMonth(), 1));
    setCalendarSelectedDate(getTodayYMD());
  };

  const handleOpenAddEvent = (date?: string) => {
    const targetDate = date || calendarSelectedDate || getTodayYMD();
    setEditingEventId(null);
    setEventTitleInput('');
    setEventDateInput(targetDate);
    setEventTimeInput('10:00 AM');
    setIsEventAllDay(false);
    setEventTypeInput('exam');
    setEventClassInput(classes.length > 0 ? (activeClass?.name || 'All Classes') : 'All Classes');
    setEventDescriptionInput('');
    setEventModalError(null);
    setIsEventModalOpen(true);
  };

  const handleOpenEditEvent = (evt: CalendarEvent) => {
    setEditingEventId(evt.id);
    setEventTitleInput(evt.title);
    setEventDateInput(evt.date);
    setEventTimeInput(evt.time || '10:00 AM');
    setIsEventAllDay(!evt.time || evt.time === 'All Day');
    setEventTypeInput(evt.type);
    setEventClassInput(evt.className || 'All Classes');
    setEventDescriptionInput(evt.description || '');
    setEventModalError(null);
    setIsEventModalOpen(true);
  };

  const handleSaveEvent = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const trimmedTitle = eventTitleInput.trim();
    if (!trimmedTitle) {
      setEventModalError('Please enter an event or note title.');
      return;
    }
    if (!eventDateInput) {
      setEventModalError('Please choose a valid date.');
      return;
    }

    const payload: CalendarEvent = {
      id: editingEventId || `evt-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      title: trimmedTitle,
      date: eventDateInput,
      time: isEventAllDay ? 'All Day' : (eventTimeInput.trim() || 'All Day'),
      type: eventTypeInput,
      className: eventClassInput === 'all' || !eventClassInput ? 'All Classes' : eventClassInput,
      description: eventDescriptionInput.trim() || undefined,
      createdAt: editingEventId
        ? (calendarEvents.find(ev => ev.id === editingEventId)?.createdAt || Date.now())
        : Date.now()
    };

    if (editingEventId) {
      setCalendarEvents(prev => prev.map(ev => ev.id === editingEventId ? payload : ev));
    } else {
      setCalendarEvents(prev => [...prev, payload]);
    }

    setCalendarSelectedDate(eventDateInput);
    setIsEventModalOpen(false);
  };

  const handleDeleteEvent = (id: string) => {
    setCalendarEvents(prev => prev.filter(ev => ev.id !== id));
    setConfirmDeleteEventTarget(null);
  };

  const getEventTypeBadgeInfo = (type: CalendarEventType) => {
    switch (type) {
      case 'exam':
        return {
          label: 'Exam',
          badgeBg: 'bg-purple-100 text-purple-900 border-purple-200',
          dotBg: 'bg-purple-600',
          chipBg: 'bg-purple-50 text-purple-900 hover:bg-purple-100 border-purple-200'
        };
      case 'meeting':
        return {
          label: 'Meeting',
          badgeBg: 'bg-sky-100 text-sky-900 border-sky-200',
          dotBg: 'bg-sky-600',
          chipBg: 'bg-sky-50 text-sky-900 hover:bg-sky-100 border-sky-200'
        };
      case 'lecture':
        return {
          label: 'Lecture',
          badgeBg: 'bg-emerald-100 text-emerald-900 border-emerald-200',
          dotBg: 'bg-emerald-600',
          chipBg: 'bg-emerald-50 text-emerald-900 hover:bg-emerald-100 border-emerald-200'
        };
      case 'deadline':
        return {
          label: 'Deadline',
          badgeBg: 'bg-amber-100 text-amber-900 border-amber-200',
          dotBg: 'bg-amber-600',
          chipBg: 'bg-amber-50 text-amber-900 hover:bg-amber-100 border-amber-200'
        };
      case 'note':
      default:
        return {
          label: 'Note',
          badgeBg: 'bg-[#E3EAFA] text-[#4D5A8F] border-[#B3C0E6]',
          dotBg: 'bg-[#4D5A8F]',
          chipBg: 'bg-[#F5F8FF] text-[#252B3B] hover:bg-[#E3EAFA] border-[#CED8F2]'
        };
    }
  };

  const renderCalendarAgendaView = () => {
    const year = calendarViewDate.getFullYear();
    const month = calendarViewDate.getMonth();
    const monthName = calendarViewDate.toLocaleString('default', { month: 'long' });
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    // Week starts on Monday (0=Mon, 1=Tue, ..., 6=Sun)
    const firstDayIndex = (new Date(year, month, 1).getDay() + 6) % 7;
    const prevMonthDays = new Date(year, month, 0).getDate();
    const todayYMD = getTodayYMD();

    // Prepare grid cells
    interface GridCell {
      dayNumber: number;
      dateStr: string;
      isCurrentMonth: boolean;
      isToday: boolean;
      isSelected: boolean;
    }

    const gridCells: GridCell[] = [];

    // 1. Previous month trailing days
    for (let i = firstDayIndex - 1; i >= 0; i--) {
      const dayNum = prevMonthDays - i;
      const prevDate = new Date(year, month - 1, dayNum);
      const dateStr = formatDateToYMD(prevDate.getFullYear(), prevDate.getMonth() + 1, dayNum);
      gridCells.push({
        dayNumber: dayNum,
        dateStr,
        isCurrentMonth: false,
        isToday: dateStr === todayYMD,
        isSelected: dateStr === calendarSelectedDate
      });
    }

    // 2. Current month days
    for (let dayNum = 1; dayNum <= daysInMonth; dayNum++) {
      const dateStr = formatDateToYMD(year, month + 1, dayNum);
      gridCells.push({
        dayNumber: dayNum,
        dateStr,
        isCurrentMonth: true,
        isToday: dateStr === todayYMD,
        isSelected: dateStr === calendarSelectedDate
      });
    }

    // 3. Next month leading days to complete grid
    const remainingCells = 7 - (gridCells.length % 7);
    if (remainingCells < 7) {
      for (let dayNum = 1; dayNum <= remainingCells; dayNum++) {
        const nextDate = new Date(year, month + 1, dayNum);
        const dateStr = formatDateToYMD(nextDate.getFullYear(), nextDate.getMonth() + 1, dayNum);
        gridCells.push({
          dayNumber: dayNum,
          dateStr,
          isCurrentMonth: false,
          isToday: dateStr === todayYMD,
          isSelected: dateStr === calendarSelectedDate
        });
      }
    }

    // Filtered calendar events based on class and type filters
    const filteredEvents = calendarEvents.filter(evt => {
      if (calendarClassFilter !== 'all' && evt.className && evt.className !== 'All Classes' && evt.className !== calendarClassFilter) {
        return false;
      }
      if (calendarTypeFilter !== 'all' && evt.type !== calendarTypeFilter) {
        return false;
      }
      return true;
    });

    // Events for the selected date
    const selectedDateEvents = filteredEvents.filter(evt => evt.date === calendarSelectedDate);

    // Selected date details
    const selectedDateParts = calendarSelectedDate ? calendarSelectedDate.split('-').map(Number) : [year, month + 1, 1];
    const selectedDateObj = new Date(selectedDateParts[0], selectedDateParts[1] - 1, selectedDateParts[2]);
    const selectedWeekdayName: DayOfWeek = DAY_OF_WEEK_NAMES[selectedDateObj.getDay()];
    const relativeLabel = getRelativeDateLabel(calendarSelectedDate);

    // Recurring timetable classes for this selected weekday
    const recurringSlotsForDate = classes.flatMap(cls => {
      if (!cls.schedule) return [];
      if (calendarClassFilter !== 'all' && cls.name !== calendarClassFilter) return [];
      return cls.schedule
        .filter(slot => slot.day === selectedWeekdayName)
        .map(slot => ({
          ...slot,
          className: cls.name
        }));
    });

    return (
      <div className="flex flex-col gap-6 animate-fade-in" id="calendar-agenda-view-container">
        {/* Top Calendar Toolbar */}
        <div className="bg-white rounded-2xl border border-[#E8ECF5] p-5 shadow-xs flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3.5">
            <div className="p-3 bg-[#E3EAFA] text-[#4D5A8F] rounded-2xl shadow-3xs border border-[#CED8F2]">
              <CalendarDays className="h-6 w-6" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h2 className="text-xl font-black text-[#252B3B] tracking-tight">Calendar &amp; Agenda Book</h2>
                <span className="text-[11px] bg-[#E3EAFA] text-[#4D5A8F] font-bold px-2.5 py-0.5 rounded-full border border-[#B3C0E6]/60">
                  {filteredEvents.length} {filteredEvents.length === 1 ? 'Record' : 'Records'}
                </span>
              </div>
              <p className="text-xs text-[#747D91] mt-0.5">
                Assign exam dates, organize meeting hours, set assignment deadlines, and track daily class notes.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5 flex-wrap">
            {/* View Switch: Month Grid vs Agenda List */}
            <div className="inline-flex bg-[#F5F8FF] p-1 rounded-xl border border-[#E8ECF5] shadow-3xs">
              <button
                type="button"
                onClick={() => setCalendarViewMode('month')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                  calendarViewMode === 'month'
                    ? 'bg-[#4D5A8F] text-white shadow-3xs'
                    : 'text-[#747D91] hover:text-[#252B3B]'
                }`}
                title="Month Grid View"
              >
                <LayoutGrid className="h-3.5 w-3.5" />
                <span>Month</span>
              </button>
              <button
                type="button"
                onClick={() => setCalendarViewMode('agenda')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                  calendarViewMode === 'agenda'
                    ? 'bg-[#4D5A8F] text-white shadow-3xs'
                    : 'text-[#747D91] hover:text-[#252B3B]'
                }`}
                title="Agenda List View"
              >
                <List className="h-3.5 w-3.5" />
                <span>Agenda</span>
              </button>
            </div>

            {/* Quick Add Event / Note Button */}
            <button
              type="button"
              onClick={() => handleOpenAddEvent(calendarSelectedDate)}
              className="inline-flex items-center justify-center gap-2 px-4 py-2 bg-[#4D5A8F] hover:bg-[#3D4873] text-white font-bold text-xs rounded-xl shadow-xs transition-all cursor-pointer active:scale-95 min-h-[40px] w-full sm:w-auto"
              id="calendar-add-event-btn"
            >
              <CalendarPlus className="h-4 w-4" />
              <span>Add Event / Note</span>
            </button>
          </div>
        </div>

        {/* Filter Bar & Month Navigation */}
        <div className="bg-white rounded-2xl border border-[#E8ECF5] p-3.5 sm:p-4 shadow-3xs flex flex-col md:flex-row md:items-center md:justify-between gap-3 sm:gap-4">
          {/* Month / Year Navigator */}
          <div className="flex items-center justify-between sm:justify-start gap-2">
            <button
              type="button"
              onClick={handlePrevCalendarMonth}
              className="p-2 bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] rounded-xl border border-[#E8ECF5] transition-all cursor-pointer active:scale-95 shadow-3xs min-h-[38px] min-w-[38px] flex items-center justify-center"
              title="Previous Month"
              aria-label="Previous Month"
            >
              <ChevronLeft className="h-4 w-4 text-[#4D5A8F]" />
            </button>

            <span className="text-sm sm:text-base font-extrabold text-[#252B3B] min-w-[140px] sm:min-w-[160px] text-center font-mono">
              {monthName} {year}
            </span>

            <button
              type="button"
              onClick={handleNextCalendarMonth}
              className="p-2 bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] rounded-xl border border-[#E8ECF5] transition-all cursor-pointer active:scale-95 shadow-3xs min-h-[38px] min-w-[38px] flex items-center justify-center"
              title="Next Month"
              aria-label="Next Month"
            >
              <ChevronRight className="h-4 w-4 text-[#4D5A8F]" />
            </button>

            <button
              type="button"
              onClick={handleJumpToToday}
              className="px-3 py-1.5 bg-[#E3EAFA] hover:bg-[#CED8F2] text-[#4D5A8F] font-bold text-xs rounded-xl transition-all cursor-pointer border border-[#B3C0E6]/50 shadow-3xs min-h-[38px] flex items-center justify-center"
            >
              Today
            </button>
          </div>

          {/* Filters: Class Section & Event Type (2 columns on mobile) */}
          <div className="grid grid-cols-2 sm:flex sm:items-center gap-2 w-full sm:w-auto">
            {/* Class Section Filter */}
            <div className="flex items-center gap-1.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl px-2.5 py-1.5 shadow-3xs min-h-[38px]">
              <Tag className="h-3.5 w-3.5 text-[#6D7CB0] shrink-0" />
              <label htmlFor="calendar-class-filter-select" className="sr-only">Filter by class section</label>
              <select
                id="calendar-class-filter-select"
                value={calendarClassFilter}
                onChange={(e) => setCalendarClassFilter(e.target.value)}
                className="bg-transparent text-xs font-semibold text-[#252B3B] outline-none cursor-pointer w-full truncate"
              >
                <option value="all">All Classes</option>
                {classes.map(c => (
                  <option key={c.name} value={c.name}>{c.name}</option>
                ))}
              </select>
            </div>

            {/* Event Type Filter */}
            <div className="flex items-center gap-1.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl px-2.5 py-1.5 shadow-3xs min-h-[38px]">
              <Filter className="h-3.5 w-3.5 text-[#6D7CB0] shrink-0" />
              <label htmlFor="calendar-type-filter-select" className="sr-only">Filter by event type</label>
              <select
                id="calendar-type-filter-select"
                value={calendarTypeFilter}
                onChange={(e) => setCalendarTypeFilter(e.target.value)}
                className="bg-transparent text-xs font-semibold text-[#252B3B] outline-none cursor-pointer w-full truncate"
              >
                <option value="all">All Types</option>
                <option value="exam">Exams</option>
                <option value="meeting">Meetings</option>
                <option value="lecture">Lectures</option>
                <option value="deadline">Deadlines</option>
                <option value="note">Notes &amp; Reminders</option>
              </select>
            </div>
          </div>
        </div>

        {/* Calendar Body: Month Grid or Agenda List */}
        {calendarViewMode === 'month' ? (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 sm:gap-6 items-start">
            {/* Left 8 Cols: Month Calendar Grid */}
            <div className="lg:col-span-8 bg-white rounded-2xl border border-[#E8ECF5] p-2.5 sm:p-5 shadow-xs flex flex-col gap-3">
              {/* Day-of-week Headers */}
              <div className="grid grid-cols-7 gap-1 sm:gap-2 text-center pb-2 border-b border-[#E8ECF5]">
                {['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].map((dayStr, idx) => (
                  <div
                    key={dayStr}
                    className={`text-[10px] sm:text-xs font-extrabold uppercase tracking-wider py-1 ${
                      idx >= 5 ? 'text-[#8F9DCC]' : 'text-[#4D5A8F]'
                    }`}
                  >
                    {dayStr}
                  </div>
                ))}
              </div>

              {/* Day Cells Grid */}
              <div className="grid grid-cols-7 gap-1 sm:gap-2">
                {gridCells.map((cell) => {
                  const cellEvents = filteredEvents.filter(e => e.date === cell.dateStr);
                  const isSelected = cell.dateStr === calendarSelectedDate;

                  return (
                    <div
                      key={cell.dateStr}
                      onClick={() => setCalendarSelectedDate(cell.dateStr)}
                      className={`min-h-[58px] sm:min-h-[105px] p-1 sm:p-2 rounded-xl border transition-all cursor-pointer flex flex-col justify-between group relative ${
                        isSelected
                          ? 'border-[#4D5A8F] ring-2 ring-[#4D5A8F]/30 bg-[#F5F8FF] shadow-3xs'
                          : cell.isCurrentMonth
                            ? 'bg-white hover:bg-[#F5F8FF] border-[#E8ECF5] hover:border-[#B3C0E6]'
                            : 'bg-[#F5F8FF]/40 border-transparent text-[#747D91]/50 opacity-60 hover:opacity-100 hover:bg-[#F5F8FF]'
                      }`}
                    >
                      {/* Cell Header: Day Number & Indicator */}
                      <div className="flex items-center justify-between">
                        <span
                          className={`inline-flex items-center justify-center h-5 w-5 sm:h-6 sm:w-6 rounded-full text-[10px] sm:text-xs font-bold ${
                            cell.isToday
                              ? 'bg-[#4D5A8F] text-white shadow-3xs font-extrabold'
                              : isSelected
                                ? 'bg-[#CED8F2] text-[#252B3B]'
                                : cell.isCurrentMonth
                                  ? 'text-[#252B3B]'
                                  : 'text-[#747D91]'
                          }`}
                        >
                          {cell.dayNumber}
                        </span>

                        {/* Quick '+' Hover Add Button on tablet/desktop */}
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleOpenAddEvent(cell.dateStr);
                          }}
                          className="hidden sm:block opacity-0 group-hover:opacity-100 p-1 hover:bg-[#E3EAFA] text-[#4D5A8F] rounded-md transition-opacity cursor-pointer"
                          title={`Add note or event on ${cell.dateStr}`}
                        >
                          <Plus className="h-3.5 w-3.5" />
                        </button>
                      </div>

                      {/* Mobile Indicator Dots (clean touch screen pattern) */}
                      <div className="flex sm:hidden items-center justify-center gap-0.5 mt-1 flex-wrap">
                        {cellEvents.slice(0, 3).map((evt) => {
                          const badge = getEventTypeBadgeInfo(evt.type);
                          return (
                            <span 
                              key={evt.id} 
                              className={`w-1.5 h-1.5 rounded-full ${badge.dotBg}`} 
                              title={`${evt.title} (${badge.label})`}
                            />
                          );
                        })}
                        {cellEvents.length > 3 && (
                          <span className="w-1 h-1 rounded-full bg-[#747D91]" />
                        )}
                      </div>

                      {/* Tablet/Desktop Cell Event Chips (max 2 visible chips) */}
                      <div className="hidden sm:flex flex-col gap-1 mt-1 flex-1 overflow-hidden">
                        {cellEvents.slice(0, 2).map((evt) => {
                          const badge = getEventTypeBadgeInfo(evt.type);
                          return (
                            <div
                              key={evt.id}
                              className={`text-[10px] font-semibold px-1.5 py-0.5 rounded truncate border flex items-center gap-1 ${badge.chipBg}`}
                              title={`${evt.title} (${badge.label}${evt.time ? ` • ${evt.time}` : ''})`}
                            >
                              <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${badge.dotBg}`} />
                              <span className="truncate">{evt.title}</span>
                            </div>
                          );
                        })}

                        {cellEvents.length > 2 && (
                          <span className="text-[9px] font-bold text-[#6D7CB0] pl-1">
                            +{cellEvents.length - 2} more
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Grid Footer Legend */}
              <div className="flex items-center justify-between pt-3 border-t border-[#E8ECF5] text-[11px] text-[#747D91] flex-wrap gap-2">
                <div className="flex items-center gap-3 flex-wrap">
                  <span className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-purple-600" />
                    <span>Exam</span>
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-sky-600" />
                    <span>Meeting</span>
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-600" />
                    <span>Lecture</span>
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-600" />
                    <span>Deadline</span>
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-[#4D5A8F]" />
                    <span>Note</span>
                  </span>
                </div>
                <span className="text-[10px] font-mono text-[#747D91]">
                  Click any date to inspect full agenda and notes
                </span>
              </div>
            </div>

            {/* Right 4 Cols: Selected Day Agenda & Note Inspector */}
            <div className="lg:col-span-4 bg-white rounded-2xl border border-[#E8ECF5] p-5 shadow-xs flex flex-col gap-4 sticky top-6" id="calendar-day-agenda-inspector">
              <div className="flex items-start justify-between border-b border-[#E8ECF5] pb-3.5">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-[#6D7CB0]">
                      Day Agenda &amp; Notes
                    </span>
                    {relativeLabel && (
                      <span className="text-[10px] bg-[#E3EAFA] text-[#4D5A8F] font-bold px-2 py-0.5 rounded-full border border-[#B3C0E6]/50">
                        {relativeLabel}
                      </span>
                    )}
                  </div>
                  <h3 className="text-base font-extrabold text-[#252B3B] mt-1 leading-snug">
                    {formatFullDateString(calendarSelectedDate)}
                  </h3>
                </div>

                <button
                  type="button"
                  onClick={() => handleOpenAddEvent(calendarSelectedDate)}
                  className="p-2 bg-[#E3EAFA] hover:bg-[#CED8F2] text-[#4D5A8F] rounded-xl transition-all cursor-pointer shadow-3xs active:scale-95 shrink-0"
                  title="Add note or event for this date"
                >
                  <Plus className="h-4 w-4" />
                </button>
              </div>

              {/* Recurring Timetable Classes for this Weekday */}
              {recurringSlotsForDate.length > 0 && (
                <div className="bg-[#F5F8FF] rounded-xl border border-[#E8ECF5] p-3 flex flex-col gap-2">
                  <div className="flex items-center gap-1.5 text-[11px] font-bold text-[#4D5A8F]">
                    <Clock className="h-3.5 w-3.5" />
                    <span>Weekly Class Schedule ({selectedWeekdayName})</span>
                  </div>
                  <div className="flex flex-col gap-1.5">
                    {recurringSlotsForDate.map(slot => (
                      <div key={slot.id} className="bg-white rounded-lg border border-[#E8ECF5] px-2.5 py-1.5 flex items-center justify-between text-xs shadow-3xs">
                        <span className="font-bold text-[#252B3B] truncate">{slot.className}</span>
                        <span className="font-mono text-[11px] text-[#6D7CB0] font-semibold shrink-0">
                          {slot.startTime} - {slot.endTime}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Day Events & Notes Listing */}
              <div className="flex flex-col gap-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-[#252B3B] uppercase tracking-wider">
                    Scheduled Notes &amp; Events ({selectedDateEvents.length})
                  </span>
                  <button
                    type="button"
                    onClick={() => handleOpenAddEvent(calendarSelectedDate)}
                    className="text-[11px] font-bold text-[#4D5A8F] hover:underline cursor-pointer"
                  >
                    + Add New
                  </button>
                </div>

                {selectedDateEvents.length === 0 ? (
                  <div className="py-8 flex flex-col items-center justify-center text-center bg-[#F5F8FF]/60 rounded-xl border border-dashed border-[#CED8F2] p-4">
                    <CalendarDays className="h-8 w-8 text-[#8F9DCC] mb-2 opacity-60" />
                    <span className="text-xs font-bold text-[#252B3B]">No entries for this date</span>
                    <p className="text-[11px] text-[#747D91] mt-1 max-w-[220px]">
                      Add an exam reminder, meeting hour, homework deadline, or general academic note.
                    </p>
                    <button
                      type="button"
                      onClick={() => handleOpenAddEvent(calendarSelectedDate)}
                      className="mt-3 px-3 py-1.5 bg-[#4D5A8F] text-white font-bold text-xs rounded-xl shadow-xs hover:bg-[#3D4873] transition-all cursor-pointer"
                    >
                      + Add Note / Event
                    </button>
                  </div>
                ) : (
                  <div className="flex flex-col gap-2.5 max-h-[460px] overflow-y-auto pr-1">
                    {selectedDateEvents.map(evt => {
                      const badge = getEventTypeBadgeInfo(evt.type);
                      return (
                        <div
                          key={evt.id}
                          className="bg-white rounded-xl border border-[#E8ECF5] p-3.5 shadow-3xs flex flex-col gap-2 hover:border-[#B3C0E6] transition-all group"
                        >
                          <div className="flex items-start justify-between gap-2">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${badge.badgeBg}`}>
                                {badge.label}
                              </span>
                              {evt.className && (
                                <span className="text-[10px] font-semibold bg-[#F5F8FF] text-[#252B3B] px-2 py-0.5 rounded-md border border-[#E8ECF5]">
                                  {evt.className}
                                </span>
                              )}
                              <span className="text-[10px] font-mono text-[#6D7CB0] font-bold">
                                🕒 {evt.time || 'All Day'}
                              </span>
                            </div>

                            <div className="flex items-center gap-1 opacity-80 group-hover:opacity-100 transition-opacity">
                              <button
                                type="button"
                                onClick={() => handleOpenEditEvent(evt)}
                                className="p-1 hover:bg-[#E3EAFA] text-[#4D5A8F] rounded-lg transition-colors cursor-pointer"
                                title="Edit this entry"
                              >
                                <Pencil className="h-3.5 w-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => setConfirmDeleteEventTarget(evt)}
                                className="p-1 hover:bg-rose-50 text-rose-600 rounded-lg transition-colors cursor-pointer"
                                title="Delete this entry"
                              >
                                <Trash2 className="h-3.5 w-3.5" />
                              </button>
                            </div>
                          </div>

                          <h4 className="text-sm font-bold text-[#252B3B] tracking-tight">
                            {evt.title}
                          </h4>

                          {evt.description && (
                            <p className="text-xs text-[#747D91] whitespace-pre-line leading-relaxed bg-[#F5F8FF] p-2.5 rounded-lg border border-[#E8ECF5]">
                              {evt.description}
                            </p>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>
          </div>
        ) : (
          /* Full Agenda List View Mode */
          <div className="bg-white rounded-2xl border border-[#E8ECF5] p-6 shadow-xs flex flex-col gap-5">
            <div className="flex items-center justify-between pb-3 border-b border-[#E8ECF5]">
              <div>
                <h3 className="text-base font-extrabold text-[#252B3B]">Chronological Agenda Timeline</h3>
                <p className="text-xs text-[#747D91]">
                  Sorted agenda listing for all academic milestones, exams, and notes.
                </p>
              </div>
              <button
                type="button"
                onClick={() => handleOpenAddEvent()}
                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-[#4D5A8F] hover:bg-[#3D4873] text-white font-bold text-xs rounded-xl shadow-xs transition-all cursor-pointer"
              >
                <Plus className="h-3.5 w-3.5" />
                <span>Create New Entry</span>
              </button>
            </div>

            {filteredEvents.length === 0 ? (
              <div className="py-12 flex flex-col items-center justify-center text-center">
                <CalendarDays className="h-10 w-10 text-[#8F9DCC] mb-2" />
                <span className="text-sm font-bold text-[#252B3B]">No calendar records match current filters</span>
                <p className="text-xs text-[#747D91] mt-1 max-w-sm">
                  Try clearing your filters or create a new event or note to populate your agenda.
                </p>
              </div>
            ) : (
              <div className="flex flex-col gap-4">
                {/* Group events by date */}
                {Array.from(new Set(filteredEvents.map(e => e.date)))
                  .sort()
                  .map(dateStr => {
                    const groupEvents = filteredEvents.filter(e => e.date === dateStr);
                    const isToday = dateStr === todayYMD;
                    const relative = getRelativeDateLabel(dateStr);

                    return (
                      <div key={dateStr} className="flex flex-col gap-2.5">
                        <div className="flex items-center gap-2.5 pt-2">
                          <span className={`text-xs font-bold px-2.5 py-1 rounded-lg border ${
                            isToday
                              ? 'bg-[#4D5A8F] text-white border-[#4D5A8F]'
                              : 'bg-[#E3EAFA] text-[#4D5A8F] border-[#B3C0E6]/60'
                          }`}>
                            {formatFullDateString(dateStr)}
                          </span>
                          {relative && (
                            <span className="text-[10px] font-bold text-[#6D7CB0] bg-[#F5F8FF] px-2 py-0.5 rounded-full border border-[#E8ECF5]">
                              {relative}
                            </span>
                          )}
                          <div className="h-px bg-[#E8ECF5] flex-1" />
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                          {groupEvents.map(evt => {
                            const badge = getEventTypeBadgeInfo(evt.type);
                            return (
                              <div
                                key={evt.id}
                                className="bg-white rounded-xl border border-[#E8ECF5] p-3.5 shadow-3xs flex flex-col justify-between gap-3 hover:border-[#B3C0E6] transition-all group"
                              >
                                <div className="flex flex-col gap-2">
                                  <div className="flex items-start justify-between gap-2">
                                    <div className="flex items-center gap-1.5 flex-wrap">
                                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${badge.badgeBg}`}>
                                        {badge.label}
                                      </span>
                                      {evt.className && (
                                        <span className="text-[10px] font-semibold bg-[#F5F8FF] text-[#252B3B] px-2 py-0.5 rounded-md border border-[#E8ECF5]">
                                          {evt.className}
                                        </span>
                                      )}
                                    </div>
                                    <span className="text-[10px] font-mono text-[#6D7CB0] font-bold">
                                      {evt.time || 'All Day'}
                                    </span>
                                  </div>

                                  <h4 className="text-sm font-bold text-[#252B3B] tracking-tight">
                                    {evt.title}
                                  </h4>

                                  {evt.description && (
                                    <p className="text-xs text-[#747D91] whitespace-pre-line leading-relaxed bg-[#F5F8FF] p-2.5 rounded-lg border border-[#E8ECF5]">
                                      {evt.description}
                                    </p>
                                  )}
                                </div>

                                <div className="flex items-center justify-between pt-2 border-t border-[#E8ECF5] text-xs">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setCalendarSelectedDate(evt.date);
                                      setCalendarViewMode('month');
                                    }}
                                    className="text-[11px] font-semibold text-[#4D5A8F] hover:underline cursor-pointer"
                                  >
                                    View in Month Grid →
                                  </button>
                                  <div className="flex items-center gap-1">
                                    <button
                                      type="button"
                                      onClick={() => handleOpenEditEvent(evt)}
                                      className="p-1 hover:bg-[#E3EAFA] text-[#4D5A8F] rounded-lg transition-colors cursor-pointer"
                                      title="Edit this entry"
                                    >
                                      <Pencil className="h-3.5 w-3.5" />
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => setConfirmDeleteEventTarget(evt)}
                                      className="p-1 hover:bg-rose-50 text-rose-600 rounded-lg transition-colors cursor-pointer"
                                      title="Delete this entry"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#F5F8FF] flex flex-col text-[#252B3B] font-sans antialiased" id="omni-class-app-container">
      {/* Omni Class Brand Accent Strip: Periwinkle SaaS Theme */}
      <div className="h-1 bg-gradient-to-r from-[#B3C0E6] via-[#8F9DCC] via-[#6D7CB0] to-[#4D5A8F] w-full" id="omni-accent-line"></div>

      {/* Header Banner: Clean Single-Row Responsive Layout */}
      <header className="bg-white/95 backdrop-blur-md border-b border-[#E8ECF5] px-3.5 sm:px-6 py-2.5 sm:py-3.5 flex items-center justify-between gap-2.5 shadow-3xs sticky top-0 z-30" id="omni-header">
        <div className="flex items-center gap-2.5 sm:gap-3.5 min-w-0">
          <button 
            type="button"
            onClick={() => setIsSidebarOpen(prev => !prev)}
            className={`p-2 sm:p-2.5 rounded-xl transition-all cursor-pointer flex items-center justify-center border shadow-3xs active:scale-95 shrink-0 ${
              isSidebarOpen 
                ? 'bg-[#4D5A8F] text-white border-[#4D5A8F]' 
                : 'bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] border-[#E8ECF5]'
            }`}
            id="toggle-sidebar-hamburger"
            title={isSidebarOpen ? "Hide sidebar menu" : "Show sidebar menu"}
            aria-label={isSidebarOpen ? "Hide sidebar menu" : "Show sidebar menu"}
          >
            <Menu className="h-5 w-5" />
          </button>
          
          <div className="p-2 sm:p-2.5 bg-[#4D5A8F] text-white rounded-xl sm:rounded-2xl shadow-3xs flex items-center justify-center shrink-0">
            <GraduationCap className="h-5 w-5 sm:h-6 sm:w-6 text-white" />
          </div>

          <div className="min-w-0">
            <h1 className="text-base sm:text-lg md:text-xl font-extrabold text-[#252B3B] tracking-tight truncate">
              Omni Class
            </h1>
            <p className="text-[10px] sm:text-[11px] font-medium text-[#6D7CB0] leading-none mt-0.5 hidden xs:block truncate">Academic Portal</p>
          </div>
        </div>
        
        {/* Quick Class Section Switcher on Mobile/Tablet & Sync Indicator */}
        <div className="flex items-center gap-2 shrink-0">
          {classes.length > 1 && (
            <div className="relative flex items-center">
              <label htmlFor="quick-header-class-select" className="sr-only">Choose active class</label>
              <select
                id="quick-header-class-select"
                value={activeClassName}
                onChange={(e) => setActiveClassName(e.target.value)}
                className="text-xs font-bold bg-[#F5F8FF] text-[#252B3B] border border-[#E8ECF5] rounded-xl px-2.5 py-1.5 pr-7 focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20 outline-none cursor-pointer shadow-3xs truncate max-w-[125px] sm:max-w-[200px]"
                title="Switch active class section"
              >
                {classes.map(c => (
                  <option key={c.name} value={c.name}>{c.name}</option>
                ))}
              </select>
              <ChevronDown className="h-3 w-3 text-[#747D91] pointer-events-none absolute right-2" />
            </div>
          )}

          {user ? (
            <span className="inline-flex bg-[#E3EAFA] text-[#4D5A8F] font-mono text-[10px] uppercase font-bold items-center px-2 sm:px-3 py-1 sm:py-1.5 rounded-full border border-[#B3C0E6]/70 gap-1.5 shadow-3xs select-none" id="data-sync-badge" title={`Connected as ${user.email}`}>
              <span className="w-1.5 h-1.5 bg-[#4D5A8F] rounded-full animate-pulse"></span>
              <span className="hidden sm:inline">Sync ON</span>
              <span className="sm:hidden">Live</span>
            </span>
          ) : (
            <span className="inline-flex bg-[#F5F8FF] text-[#747D91] font-mono text-[10px] uppercase font-bold items-center px-2 sm:px-3 py-1 sm:py-1.5 rounded-full border border-[#E8ECF5] gap-1.5 shadow-3xs select-none" id="data-sync-badge" title="Offline mode - changes saved locally">
              <span className="w-1.5 h-1.5 bg-[#8F9DCC] rounded-full"></span>
              <span className="hidden sm:inline">Sync OFF</span>
              <span className="sm:hidden">Local</span>
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
            className="fixed inset-0 bg-[#252B3B]/40 z-40 lg:hidden backdrop-blur-3xs transition-opacity duration-300"
            id="sidebar-mobile-backdrop"
          />
        )}

        {/* SIDEBAR WIDGET PANELS: White surface with deep periwinkle active navigation */}
        <aside className={`
          fixed inset-y-0 left-0 z-50 w-80 bg-white text-[#252B3B] p-5 flex flex-col shrink-0 gap-4 border-r border-[#E8ECF5] shadow-lg lg:shadow-none overflow-y-auto transition-all duration-300
          lg:self-start lg:sticky lg:top-0 lg:max-h-screen
          ${isSidebarOpen ? 'translate-x-0 flex lg:flex lg:w-80' : '-translate-x-full hidden lg:hidden'}
        `} id="omni-sidebar">
          
          {/* Logo & Close Drawer Header */}
          <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-3" id="sidebar-drawer-header">
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-[#4D5A8F] text-white rounded-2xl shadow-3xs flex items-center justify-center">
                <GraduationCap className="h-5 w-5 text-white" />
              </div>
              <div>
                <span className="text-base font-bold text-[#252B3B] tracking-tight block">Omni Class</span>
                <span className="text-[11px] text-[#6D7CB0] block font-medium">Academic Portal</span>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setIsSidebarOpen(false)}
              className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
              title="Close sidebar"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Cloud Database (Firebase Sync Widget) - First in Sidebar */}
          <div className="bg-[#F5F8FF] p-4 rounded-2xl border border-[#E8ECF5] shadow-3xs flex flex-col gap-3" id="sidebar-cloud-database">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-[#252B3B] uppercase tracking-wider flex items-center gap-1.5">
                <Cloud className={`h-4 w-4 ${user ? 'text-[#4D5A8F]' : 'text-[#8F9DCC]'}`} />
                Cloud Database
              </span>
              
              {user && (
                <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold ${
                  isSyncing ? 'bg-[#E3EAFA] text-[#4D5A8F] animate-pulse border border-[#8F9DCC]' : 'bg-[#E3EAFA] text-[#4D5A8F] border border-[#B3C0E6]/60'
                }`}>
                  <span className={`w-1 h-1 rounded-full mr-1 ${isSyncing ? 'bg-[#4D5A8F] animate-ping' : 'bg-[#4D5A8F]'}`} />
                  {isSyncing ? "Syncing..." : "Live"}
                </span>
              )}
            </div>

            {authLoading ? (
              <div className="flex items-center gap-2 text-xs text-[#747D91] font-medium py-1">
                <RefreshCw className="h-3 w-3 animate-spin text-[#4D5A8F]" />
                <span>Checking credentials...</span>
              </div>
            ) : user ? (
              <div className="flex flex-col gap-2">
                <div className="text-[11px] text-[#252B3B] bg-white p-2.5 rounded-xl border border-[#E8ECF5] flex flex-col gap-1 shadow-3xs">
                  <div className="font-semibold text-[#252B3B] truncate flex items-center gap-1.5" title={user.email || ""}>
                    <span className="w-1.5 h-1.5 rounded-full bg-[#4D5A8F]"></span>
                    <span className="truncate">{user.displayName || user.email?.split('@')[0] || "Professor"}</span>
                  </div>
                  <div className="text-[10px] text-[#747D91] truncate font-mono">{user.email}</div>
                </div>

                <button
                  type="button"
                  onClick={handleSignOut}
                  className="w-full flex items-center justify-center gap-1.5 px-3 py-2 bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-800 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-3xs"
                >
                  <LogOut className="h-3.5 w-3.5" />
                  Sign Out
                </button>
              </div>
            ) : (
              <div className="flex flex-col gap-2.5">
                <p className="text-[11px] text-[#747D91] leading-normal font-medium">
                  Connect cloud database to access your class data from any device securely.
                </p>

                <button
                  type="button"
                  onClick={handleGoogleSignIn}
                  disabled={isSyncing}
                  className="w-full flex items-center justify-center gap-2.5 px-3 py-2.5 bg-white hover:bg-[#E3EAFA] text-[#252B3B] rounded-xl text-xs font-bold transition-all cursor-pointer shadow-3xs border border-[#E8ECF5] hover:border-[#8F9DCC] active:scale-[0.98] disabled:opacity-50"
                  id="google-login-btn"
                >
                  {isSyncing ? (
                    <RefreshCw className="h-4 w-4 animate-spin text-[#4D5A8F]" />
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
              <div className="bg-[#fee2e2] text-[#7f1d1d] border border-[#fca5a5] p-2.5 rounded-xl text-[10px] font-medium leading-normal flex items-start gap-1.5">
                <AlertCircle className="h-3.5 w-3.5 shrink-0 mt-0.5" />
                <span>{syncErrorMessage}</span>
              </div>
            )}
          </div>

          {/* Primary Navigation Menu: Timetable -> Student Roster -> Attendance -> Grades */}
          <div className="flex flex-col gap-1" id="sidebar-primary-nav">
            <span className="text-[10px] font-extrabold text-[#747D91] uppercase tracking-wider px-1 pb-1">
              Main Navigation
            </span>
            
            {/* 1. Weekly Timetable */}
            <button
              type="button"
              onClick={() => handleSelectTab('timetable')}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all cursor-pointer ${
                activeTab === 'timetable'
                  ? 'bg-[#4D5A8F] text-white shadow-3xs font-bold'
                  : 'text-[#252B3B] hover:bg-[#F5F8FF] hover:text-[#4D5A8F] font-medium'
              }`}
              id="sidebar-nav-timetable"
            >
              <Clock className={`h-4 w-4 shrink-0 ${activeTab === 'timetable' ? 'text-white' : 'text-[#6D7CB0]'}`} />
              <span className={activeTab === 'timetable' ? 'text-white font-bold' : 'text-[#252B3B]'}>Weekly Timetable</span>
            </button>

            {/* Calendar & Agenda */}
            <button
              type="button"
              onClick={() => handleSelectTab('calendar')}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all cursor-pointer ${
                activeTab === 'calendar'
                  ? 'bg-[#4D5A8F] text-white shadow-3xs font-bold'
                  : 'text-[#252B3B] hover:bg-[#F5F8FF] hover:text-[#4D5A8F] font-medium'
              }`}
              id="sidebar-nav-calendar"
            >
              <CalendarDays className={`h-4 w-4 shrink-0 ${activeTab === 'calendar' ? 'text-white' : 'text-[#6D7CB0]'}`} />
              <span className={activeTab === 'calendar' ? 'text-white font-bold' : 'text-[#252B3B]'}>Calendar &amp; Agenda</span>
            </button>

            {/* 2. Student Roster */}
            <button
              type="button"
              onClick={() => handleSelectTab('students')}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all cursor-pointer ${
                activeTab === 'students'
                  ? 'bg-[#4D5A8F] text-white shadow-3xs font-bold'
                  : 'text-[#252B3B] hover:bg-[#F5F8FF] hover:text-[#4D5A8F] font-medium'
              }`}
              id="sidebar-nav-students"
            >
              <Users className={`h-4 w-4 shrink-0 ${activeTab === 'students' ? 'text-white' : 'text-[#6D7CB0]'}`} />
              <span className={activeTab === 'students' ? 'text-white font-bold' : 'text-[#252B3B]'}>Student Roster</span>
            </button>

            {/* 3. Attendance Registry */}
            <button
              type="button"
              onClick={() => handleSelectTab('attendance')}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all cursor-pointer ${
                activeTab === 'attendance'
                  ? 'bg-[#4D5A8F] text-white shadow-3xs font-bold'
                  : 'text-[#252B3B] hover:bg-[#F5F8FF] hover:text-[#4D5A8F] font-medium'
              }`}
              id="sidebar-nav-attendance"
            >
              <CalendarCheck className={`h-4 w-4 shrink-0 ${activeTab === 'attendance' ? 'text-white' : 'text-[#6D7CB0]'}`} />
              <span className={activeTab === 'attendance' ? 'text-white font-bold' : 'text-[#252B3B]'}>Attendance Registry</span>
            </button>

            {/* 4. Academic Grades */}
            <button
              type="button"
              onClick={() => handleSelectTab('grades')}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all cursor-pointer ${
                activeTab === 'grades'
                  ? 'bg-[#4D5A8F] text-white shadow-3xs font-bold'
                  : 'text-[#252B3B] hover:bg-[#F5F8FF] hover:text-[#4D5A8F] font-medium'
              }`}
              id="sidebar-nav-grades"
            >
              <GraduationCap className={`h-4 w-4 shrink-0 ${activeTab === 'grades' ? 'text-white' : 'text-[#6D7CB0]'}`} />
              <span className={activeTab === 'grades' ? 'text-white font-bold' : 'text-[#252B3B]'}>Academic Grades</span>
            </button>
          </div>

          {/* Active Section Selector */}
          <div className="flex flex-col gap-2" id="sidebar-class-selector">
            <label className="text-[10px] font-extrabold text-[#747D91] uppercase tracking-widest pl-0.5">
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
                    ? 'bg-[#F5F8FF] text-[#747D91]/60 border border-[#E8ECF5] cursor-not-allowed italic'
                    : 'bg-[#F5F8FF] border border-[#E8ECF5] text-[#252B3B] cursor-pointer focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20'
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
                className="w-10 h-10 bg-[#4D5A8F] hover:bg-[#3D4873] text-white rounded-xl transition-all flex items-center justify-center cursor-pointer shadow-sm active:scale-95 shrink-0"
                title="Create New Class (with optional import)"
                id="create-new-class-sidebar-btn"
                aria-label="Add new class section"
              >
                <Plus className="h-5 w-5 text-white" strokeWidth={2.5} />
              </button>
            </div>
            <p className="text-[10px] text-[#747D91] font-mono pl-1">
              {classes.length === 0 ? 'No section selected' : `Roster size: ${activeClass.students.length} students enrolled`}
            </p>
          </div>

          {/* Delete Active Section Controller */}
          <div className="flex flex-col gap-2 pt-0.5" id="sidebar-delete-class-box">
            {classes.length === 0 ? (
              <div className="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-[11px] font-medium text-[#747D91]/50 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl select-none" title="No active sections to delete">
                <Trash2 className="h-3.5 w-3.5 text-[#747D91]/40" />
                <span>No sections to delete</span>
              </div>
            ) : confirmDeleteClass ? (
              <div className="flex flex-col gap-2.5 p-3 rounded-2xl bg-[#fee2e2] border-2 border-[#fca5a5] text-left w-full shadow-sm text-[#7f1d1d]">
                <div className="flex items-start gap-2">
                  <Trash2 className="h-4 w-4 text-[#991b1b] shrink-0 mt-0.5" />
                  <p className="text-xs text-[#7f1d1d] leading-snug font-medium">
                    Are you sure you want to delete <strong className="font-bold underline text-[#991b1b]">'{activeClass.name}'</strong>? {classes.length === 1 ? 'This is your only class section. All' : 'All'} student records, grades, and attendance will be permanently removed.
                  </p>
                </div>
                <div className="flex items-center gap-2 justify-end pt-1">
                  <button
                    type="button"
                    onClick={() => setConfirmDeleteClass(false)}
                    className="text-xs bg-white border border-[#E8ECF5] rounded-lg text-[#252B3B] px-3 py-1.5 hover:bg-[#F5F8FF] font-semibold transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={handleDeleteActiveSection}
                    className="text-xs bg-[#b91c1c] hover:bg-[#991b1b] text-white rounded-lg px-3.5 py-1.5 font-bold shadow-sm transition-all cursor-pointer active:scale-95 flex items-center gap-1.5"
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
                className="w-full inline-flex items-center justify-center gap-2 px-3 py-2 text-xs font-bold text-rose-700 bg-white hover:bg-rose-50 border border-[#E8ECF5] hover:border-rose-300 rounded-xl transition-all cursor-pointer shadow-3xs"
                title={`Delete '${activeClass.name}'`}
                id="delete-class-btn"
              >
                <Trash2 className="h-3.5 w-3.5 text-rose-600" strokeWidth={2.5} />
                <span className="truncate">Delete Class ({activeClass.name})</span>
              </button>
            )}
          </div>

          {/* GLOBAL STUDENT SEARCH */}
          <div className="flex flex-col gap-2 bg-[#F5F8FF] p-4 rounded-2xl border border-[#E8ECF5] shadow-3xs" id="sidebar-global-student-search">
            <div className="flex items-center justify-between pb-0.5">
              <span className="text-[11px] font-black text-[#252B3B] uppercase tracking-wide flex items-center gap-1.5">
                <Search className="h-3.5 w-3.5 text-[#4D5A8F] font-bold" />
                Global Search
              </span>
              {globalSearchQuery && (
                <button 
                  onClick={() => setGlobalSearchQuery('')}
                  className="text-[10px] text-[#6D7CB0] hover:text-[#4D5A8F] font-bold transition-all cursor-pointer"
                >
                  Clear
                </button>
              )}
            </div>

            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-[#747D91]">
                <Search className="h-3.5 w-3.5" />
              </span>
              <input
                type="text"
                value={globalSearchQuery}
                onChange={e => setGlobalSearchQuery(e.target.value)}
                placeholder="Type Student Name or ID..."
                className="w-full pl-9 pr-3 py-2 text-xs bg-white border border-[#E8ECF5] rounded-xl outline-none focus:border-[#4D5A8F] text-[#252B3B] placeholder-[#747D91]/70 font-medium transition-all"
              />
            </div>

            {globalSearchQuery.trim() !== '' && (
              <div className="max-h-80 overflow-y-auto mt-1 flex flex-col gap-2.5 divide-y divide-[#E8ECF5] border-t border-[#E8ECF5] pt-2" id="sidebar-search-results">
                {globalSearchResults.length === 0 ? (
                  <div className="py-5 text-center text-[#747D91] text-[10.5px]">
                    No students found matching <strong className="text-[#252B3B]">"{globalSearchQuery}"</strong>
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
                              <span className="font-mono text-[10px] font-bold text-[#252B3B] bg-[#E3EAFA] px-1.5 py-0.5 rounded border border-[#B3C0E6]/50 shadow-3xs shrink-0">
                                {student.id}
                              </span>
                              <span className="font-bold text-[#252B3B] text-xs truncate leading-tight">
                                {student.fullName}
                              </span>
                              <button
                                onClick={() => {
                                  setActiveClassName(classSection.name);
                                  setIsSidebarOpen(false);
                                }}
                                className="text-[9px] font-bold text-white bg-[#4D5A8F] hover:bg-[#3D4873] px-1.5 py-0.5 rounded transition-all cursor-pointer ml-auto shrink-0 shadow-3xs"
                              >
                                {classSection.name}
                              </button>
                            </div>
                          </div>

                          <div className="text-right shrink-0">
                            <span className="text-[9px] font-bold text-[#747D91] uppercase tracking-tight block">Grade</span>
                            <span className="text-[10.5px] font-black text-[#4D5A8F] bg-[#E3EAFA] px-1.5 py-0.2 rounded-md border border-[#B3C0E6]/60 inline-block">{finalGrade.toFixed(1)}</span>
                          </div>
                        </div>

                        {/* Collapsible Exam Scores */}
                        <div className="flex flex-col gap-1.5 bg-white p-2 rounded-xl border border-[#E8ECF5]">
                          <button
                            type="button"
                            onClick={() => setExpandedSearchStudentId(isExpanded ? null : `${classSection.name}-${student.id}`)}
                            className="text-[9.5px] font-bold text-[#252B3B] hover:text-[#4D5A8F] transition-colors inline-flex items-center gap-0.5 bg-[#F5F8FF] hover:bg-[#E3EAFA] px-2 py-0.5 border border-[#E8ECF5] rounded-md w-full justify-center cursor-pointer shadow-3xs"
                          >
                            {isExpanded ? 'Hide Exam Scores' : 'Edit Exam Scores'}
                          </button>

                          {isExpanded && (
                            <div className="border-t border-[#E8ECF5] pt-2 mt-1 flex flex-col gap-1.5">
                              {assessments.length === 0 ? (
                                <span className="text-[9px] text-[#747D91] italic text-center block py-1">No assessments in {classSection.name}</span>
                              ) : (
                                <div className="grid grid-cols-1 gap-1.5">
                                  {assessments.map(ass => {
                                    const score = classSection.grades[student.id]?.[ass.name];
                                    const maxMark = ass.maxMark || 100;
                                    return (
                                      <div key={ass.name} className="flex items-center justify-between bg-[#F5F8FF] px-2 py-1 rounded-lg border border-[#E8ECF5]">
                                        <span className="text-[10px] text-[#252B3B] truncate font-semibold w-1/2">{ass.name}</span>
                                        <div className="flex items-center gap-1 w-24 justify-end">
                                          <GradeScoreInput
                                            score={score}
                                            maxMark={maxMark}
                                            className="w-14"
                                            onSave={val => handleScoreChangeCrossSection(student.id, ass.name, val, classSection.name)}
                                          />
                                          <span className="text-[9px] text-[#747D91] font-mono">/{maxMark}</span>
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
          <div className="bg-[#F5F8FF] p-4 rounded-2xl border border-[#E8ECF5] shadow-3xs flex flex-col gap-2.5" id="sidebar-exports">
            <h3 className="text-xs font-bold text-[#252B3B] uppercase tracking-wider flex items-center gap-1.5">
              <Download className="h-3.5 w-3.5 text-[#4D5A8F]" />
              Download Ledgers
            </h3>
            
            <div className="flex flex-col gap-2">
              <span className="text-[9px] font-bold text-[#747D91] tracking-wider uppercase pl-0.5">Active Class:</span>
              <button
                onClick={() => handleExportFullClass()}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-bold py-2.5 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-[#4D5A8F] hover:bg-[#3D4873] border-[#4D5A8F] text-white cursor-pointer active:scale-95'
                  : 'bg-[#4D5A8F]/40 text-white/60 border-transparent cursor-not-allowed'
                }`}
                id="export-full-class-sidebar"
                title="Download full class ledger containing both attendance records and assessments/grades"
              >
                <span className="flex items-center gap-2 truncate">
                  <FileSpreadsheet className="h-4 w-4 text-[#CED8F2] shrink-0" />
                  Full Class Info (Excel)
                </span>
                <span className="text-[9px] bg-white/20 text-white font-mono px-1.5 py-0.5 rounded font-bold shrink-0">Both</span>
              </button>

              <button
                onClick={handleExportGrades}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-white hover:bg-[#E3EAFA] border-[#E8ECF5] text-[#252B3B] cursor-pointer'
                  : 'bg-white/50 text-[#747D91]/50 border-[#E8ECF5]/50 cursor-not-allowed'
                }`}
                id="export-grades-excel-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <FileSpreadsheet className="h-4 w-4 text-[#6D7CB0] shrink-0" />
                  Grades Workbook (Excel)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-[#747D91] shrink-0" />
              </button>

              <button
                onClick={handleExportAttendance}
                disabled={activeClass.students.length === 0}
                className={`w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs transition-all ${
                  activeClass.students.length > 0 
                  ? 'bg-white hover:bg-[#E3EAFA] border-[#E8ECF5] text-[#252B3B] cursor-pointer'
                  : 'bg-white/50 text-[#747D91]/50 border-[#E8ECF5]/50 cursor-not-allowed'
                }`}
                id="export-attendance-excel-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <Calendar className="h-4 w-4 text-[#6D7CB0] shrink-0" />
                  Attendance Log (Excel)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-[#747D91] shrink-0" />
              </button>

              {/* Merged Exports System */}
              <span className="text-[9px] font-bold text-[#747D91] tracking-wider uppercase pl-0.5 mt-1">Cross-Section Consolidated:</span>
              <button
                onClick={() => handleExportMerged('full')}
                className="w-full text-xs font-bold py-2.5 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-xs bg-[#4D5A8F] hover:bg-[#3D4873] border-[#4D5A8F] text-white transition-all cursor-pointer active:scale-95"
                id="export-merged-full-sidebar"
                title="Download combined ledger for all classes containing both attendance logs and assessments/grades"
              >
                <span className="flex items-center gap-2 truncate">
                  <FileSpreadsheet className="h-4 w-4 text-[#CED8F2] shrink-0" />
                  Merged Full Info (All)
                </span>
                <span className="text-[9px] bg-white/20 text-white font-mono px-1.5 py-0.5 rounded font-bold shrink-0">Both</span>
              </button>

              <button
                onClick={() => handleExportMerged('grades')}
                className="w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs bg-white hover:bg-[#E3EAFA] border-[#E8ECF5] text-[#252B3B] transition-all cursor-pointer"
                id="export-merged-grades-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <FileSpreadsheet className="h-4 w-4 text-[#6D7CB0] shrink-0" />
                  Merged Grades (All)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-[#747D91] shrink-0" />
              </button>

              <button
                onClick={() => handleExportMerged('attendance')}
                className="w-full text-xs font-semibold py-2 px-3.5 rounded-xl border text-left flex items-center justify-between gap-1 shadow-3xs bg-white hover:bg-[#E3EAFA] border-[#E8ECF5] text-[#252B3B] transition-all cursor-pointer"
                id="export-merged-attendance-sidebar"
              >
                <span className="flex items-center gap-2 truncate">
                  <Calendar className="h-4 w-4 text-[#6D7CB0] shrink-0" />
                  Merged Attendance (All)
                </span>
                <ChevronRight className="h-3.5 w-3.5 text-[#747D91] shrink-0" />
              </button>
            </div>
          </div>

          {/* Import Spreadsheet File System */}
          <div className="bg-[#F5F8FF] p-4 rounded-2xl border border-[#E8ECF5] shadow-3xs flex flex-col gap-2.5" id="sidebar-import">
            <h3 className="text-xs font-bold text-[#252B3B] uppercase tracking-wider flex items-center gap-1.5">
              <Upload className="h-3.5 w-3.5 text-[#4D5A8F]" />
              Import Spreadsheet
            </h3>
            <p className="text-[10px] text-[#747D91] leading-relaxed">
              Upload any roster, single/merged attendance log, or grades workbook (.xlsx or .csv).
            </p>
            <label className="w-full flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl border-2 border-dashed border-[#8F9DCC] bg-white hover:bg-[#E3EAFA]/60 text-[#4D5A8F] font-bold text-xs cursor-pointer transition-all shadow-3xs active:scale-95 text-center">
              <Upload className="h-4 w-4 text-[#4D5A8F]" />
              <span>Upload Excel / CSV</span>
              <input
                type="file"
                accept=".csv, .xlsx, .xls"
                onChange={handleRosterFileUpload}
                className="hidden"
              />
            </label>
          </div>
        </aside>

        {/* MAIN DASHBOARD PANEL */}
        <main className="flex-1 p-3 sm:p-5 md:p-8 flex flex-col gap-4 sm:gap-6 pb-28 lg:pb-8 overflow-x-auto bg-[#F5F8FF] text-[#252B3B]" id="streamlit-main">
          {/* TOP CONTROLLER BAR: Standalone Weekly Timetable & Calendar & Agenda Buttons */}
          <div className="flex items-center justify-between gap-3 flex-wrap" id="main-top-controller">
            <div className="grid grid-cols-2 gap-2 sm:flex sm:items-center sm:gap-2.5 w-full sm:w-auto">
              <button
                type="button"
                onClick={() => {
                  if (activeTab === 'timetable') {
                    setActiveTab('students');
                  } else {
                    setActiveTab('timetable');
                  }
                }}
                className={`inline-flex items-center justify-center gap-2 px-3 sm:px-5 py-2.5 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer shadow-3xs active:scale-95 border min-h-[44px] ${
                  activeTab === 'timetable'
                    ? 'bg-[#4D5A8F] text-white border-[#4D5A8F] shadow-sm'
                    : 'bg-white hover:bg-[#E3EAFA] text-[#252B3B] border-[#E8ECF5]'
                }`}
                id="top-weekly-timetable-btn"
                title={activeTab === 'timetable' ? "Switch to class tables" : "Show Weekly Timetable"}
              >
                <Clock className={`h-4.5 w-4.5 shrink-0 ${activeTab === 'timetable' ? 'text-white' : 'text-[#4D5A8F]'}`} />
                <span className="truncate">Weekly Timetable</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  if (activeTab === 'calendar') {
                    setActiveTab('students');
                  } else {
                    setActiveTab('calendar');
                  }
                }}
                className={`inline-flex items-center justify-center gap-2 px-3 sm:px-5 py-2.5 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer shadow-3xs active:scale-95 border min-h-[44px] ${
                  activeTab === 'calendar'
                    ? 'bg-[#4D5A8F] text-white border-[#4D5A8F] shadow-sm'
                    : 'bg-white hover:bg-[#E3EAFA] text-[#252B3B] border-[#E8ECF5]'
                }`}
                id="top-calendar-agenda-btn"
                title={activeTab === 'calendar' ? "Switch to class tables" : "Show Calendar & Agenda"}
              >
                <CalendarDays className={`h-4.5 w-4.5 shrink-0 ${activeTab === 'calendar' ? 'text-white' : 'text-[#4D5A8F]'}`} />
                <span className="truncate">Calendar &amp; Agenda</span>
              </button>
            </div>
          </div>

          {activeTab === 'timetable' ? (
            renderTeacherWeeklyTimetable()
          ) : activeTab === 'calendar' ? (
            renderCalendarAgendaView()
          ) : (
            <div className="flex flex-col gap-5 sm:gap-6" id="app-working-views">
              {/* CHOSEN CLASS SECTION NAME HEADER */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-1" id="chosen-class-header">
                <div className="flex items-center gap-2.5 min-w-0">
                  {classes.length > 0 && (
                    <button
                      type="button"
                      onClick={() => {
                        setEditingClassTarget({
                          originalName: activeClass.name,
                          newName: activeClass.name,
                          schedule: activeClass.schedule ? [...activeClass.schedule] : []
                        });
                        setEditClassActiveTab('roster');
                        setAddStudentSuccessMsg(null);
                        setEditingClassModalError(null);
                      }}
                      className="p-2 text-[#4D5A8F] hover:text-[#252B3B] hover:bg-[#E3EAFA] rounded-xl transition-colors cursor-pointer shrink-0 min-h-[40px] min-w-[40px] flex items-center justify-center border border-transparent hover:border-[#CED8F2]"
                      title={`Edit name of ${activeClass.name}`}
                      id="edit-class-name-btn"
                      aria-label={`Edit name of ${activeClass.name}`}
                    >
                      <Pencil className="h-4 w-4 text-[#4D5A8F]" />
                    </button>
                  )}
                  <h2 className="text-xl sm:text-2xl font-black text-[#252B3B] tracking-tight truncate">
                    {classes.length === 0 ? "No Active Sections" : activeClass.name}
                  </h2>
                </div>
                {classes.length > 0 && (
                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      type="button"
                      onClick={() => handleExportFullClass()}
                      disabled={activeClass.students.length === 0}
                      className="inline-flex items-center gap-1.5 px-3 py-2 bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-95 text-white font-bold text-xs rounded-xl shadow-xs transition-all cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed min-h-[38px]"
                      title={`Download complete info (both attendance and grades) for ${activeClass.name}`}
                      id="quick-download-full-class"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>Download Full Info</span>
                    </button>
                    <div className="flex items-center gap-2 text-xs text-[#4D5A8F] font-mono font-bold bg-[#E3EAFA] px-3 py-1.5 rounded-xl border border-[#B3C0E6]/50 shadow-3xs min-h-[38px]">
                      <span>{activeClass.students.length} students</span>
                      <span>•</span>
                      <span>{activeClass.assessments.length} assessments</span>
                    </div>
                  </div>
                )}
              </div>
            {/* STAGED ROSTER VERIFICATION & PREVIEW HUB */}
              {pendingImport && (
                <div className={`rounded-2xl border shadow-sm p-6 flex flex-col gap-5 animate-fade-in ${
                  pendingImport.short.length > 0
                    ? 'bg-amber-50/70 border-amber-200'
                    : 'bg-[#E3EAFA]/60 border-[#B3C0E6]'
                }`} id="short-id-compliance-container">
                  
                  <div className="flex items-start gap-3">
                    {pendingImport.short.length > 0 ? (
                      <AlertCircle className="h-6 w-6 text-amber-600 shrink-0 mt-0.5" />
                    ) : (
                      <Check className="h-6 w-6 text-emerald-600 shrink-0 mt-0.5" />
                    )}
                    <div className="flex-1">
                      <div className="flex items-center gap-2 flex-wrap">
                        <h3 className={`text-md font-extrabold tracking-tight ${
                          pendingImport.short.length > 0 ? 'text-amber-900' : 'text-[#252B3B]'
                        }`}>
                          {pendingImport.isMultiSection
                            ? `Multi-Section Import: ${pendingImport.allSections?.length || 0} Class Sections Found`
                            : pendingImport.short.length > 0 
                              ? 'Roster Verification Required: Short Student IDs Detected' 
                              : `Spreadsheet Import Preview: ${pendingImport.className}`
                          }
                        </h3>
                        {pendingImport.attendanceDates && pendingImport.attendanceDates.length > 0 && (
                          <span className="text-[10px] bg-blue-100 text-blue-800 font-bold px-2 py-0.5 rounded-full border border-blue-200">
                            📅 {pendingImport.attendanceDates.length} Attendance Dates
                          </span>
                        )}
                        {pendingImport.assessments && pendingImport.assessments.length > 0 && (
                          <span className="text-[10px] bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded-full border border-emerald-200">
                            📊 {pendingImport.assessments.length} Assessment Columns
                          </span>
                        )}
                      </div>
                      <p className={`text-xs mt-1 leading-normal ${
                        pendingImport.short.length > 0 ? 'text-amber-700' : 'text-[#747D91]'
                      }`}>
                        {pendingImport.isMultiSection
                          ? 'This workbook contains multiple class sections. Review the breakdown below and click import to update all sections simultaneously:'
                          : pendingImport.short.length > 0 
                            ? 'We detected student records with short ID numbers (6 or less digits). Decide which candidates you wish to discard or keep in your class directory below:'
                            : 'Review the student records and metadata parsed from your spreadsheet. Confirm what to keep before saving to the database:'
                        }
                      </p>
                    </div>
                  </div>

                  {/* Multi-Section Breakdown Cards */}
                  {pendingImport.isMultiSection && pendingImport.allSections && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                      {pendingImport.allSections.map(sec => (
                        <div key={sec.className} className="bg-white rounded-xl border border-[#E8ECF5] p-3.5 flex flex-col gap-2 shadow-3xs">
                          <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-2">
                            <span className="font-bold text-xs text-[#252B3B] truncate">{sec.className}</span>
                            <span className="text-[10px] bg-[#E3EAFA] text-[#4D5A8F] font-bold px-2 py-0.5 rounded-md">
                              {sec.students.length} students
                            </span>
                          </div>
                          <div className="flex flex-col gap-1 text-[11px] text-[#747D91]">
                            {sec.attendanceDates.length > 0 && (
                              <span className="flex items-center gap-1 text-blue-700">
                                📅 {sec.attendanceDates.length} attendance dates
                              </span>
                            )}
                            {sec.assessments.length > 0 && (
                              <span className="flex items-center gap-1 text-emerald-700">
                                📊 {sec.assessments.length} assessments: {sec.assessments.map(a => a.name).join(', ')}
                              </span>
                            )}
                            <div className="text-[10px] text-[#747D91] truncate mt-1">
                              e.g. {sec.students.slice(0, 2).map(s => s.fullName).join(', ')}
                              {sec.students.length > 2 ? '...' : ''}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Single Section Student Checklist */}
                  {!pendingImport.isMultiSection && (
                    <div className="bg-white border border-[#E8ECF5] rounded-xl max-h-72 overflow-y-auto divide-y divide-[#E8ECF5] shadow-3xs text-left">
                      
                      {/* PRIORITY SECTION: Short Student IDs */}
                      {pendingImport.short.length > 0 && (
                        <div>
                          <div className="bg-amber-100/65 text-amber-950 font-bold px-3.5 py-2 text-[10px] uppercase tracking-wide flex items-center justify-between sticky top-0 z-10 border-b border-amber-200/55 backdrop-blur-xs">
                            <span className="flex items-center gap-1">
                              ⚠️ Priority Review: Short IDs ({pendingImport.short.length})
                            </span>
                            <span className="text-[9px] bg-amber-200 text-amber-900 font-bold px-1.5 py-0.5 rounded-md">
                              Review Recommended
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
                          <div className="bg-[#E3EAFA] text-[#252B3B] font-bold px-3.5 py-2 text-[10px] uppercase tracking-wide flex items-center justify-between sticky top-0 z-10 border-b border-[#E8ECF5] backdrop-blur-xs">
                            <span>✅ General Roster List ({pendingImport.normal.length})</span>
                            <span className="text-[9px] bg-white text-[#4D5A8F] font-bold px-1.5 py-0.5 rounded-md border border-[#E8ECF5]">
                              Normal ID Formats
                            </span>
                          </div>
                          <div className="divide-y divide-[#E8ECF5]">
                            {pendingImport.normal.map(st => {
                              const isChecked = selectedShortIds.includes(st.id);
                              return (
                                <div key={st.id} className="flex items-center justify-between p-3.5 text-xs hover:bg-[#F5F8FF] transition-colors">
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
                                      className="h-4.5 w-4.5 text-[#4D5A8F] focus:ring-[#4D5A8F] rounded border-[#E8ECF5] cursor-pointer"
                                    />
                                    <label htmlFor={`chk-import-${st.id}`} className="flex flex-col sm:flex-row sm:items-center gap-1.5 cursor-pointer select-none">
                                      <span className="font-mono bg-[#E3EAFA] text-[#4D5A8F] px-1.5 py-0.5 rounded font-bold text-[10.5px]">ID: {st.id}</span>
                                      <span className="font-bold text-[#252B3B]">{st.fullName}</span>
                                    </label>
                                  </div>
                                  <span className="text-[10px] text-[#747D91] font-mono">({st.id.length} digits)</span>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}

                    </div>
                  )}

                  {/* Summary of checked items */}
                  {!pendingImport.isMultiSection && (
                    <div className="text-[11px] font-bold text-[#747D91] flex items-center gap-1 px-1">
                      <span>Selected to Keep:</span>
                      <span className="text-emerald-700 bg-emerald-50 px-2.5 py-0.5 rounded-full border border-emerald-100">
                        {selectedShortIds.length} of {pendingImport.short.length + pendingImport.normal.length} Students Active
                      </span>
                    </div>
                  )}

                  {/* Interactive Buttons Bar */}
                  <div className="flex items-center gap-2.5 flex-wrap border-t border-[#E8ECF5] pt-3">
                    <button
                      onClick={handleConfirmPendingImport}
                      className="bg-[#4D5A8F] hover:bg-[#3D4873] text-white font-bold text-xs px-4.5 py-2.5 rounded-xl shadow-xs transition-all cursor-pointer active:scale-95 inline-flex items-center gap-1"
                    >
                      <span>
                        {pendingImport.isMultiSection 
                          ? `Import All Sections (${pendingImport.allSections?.length || 0})` 
                          : `Save & Complete Import (${selectedShortIds.length})`
                        }
                      </span>
                    </button>
                    
                    {!pendingImport.isMultiSection && pendingImport.short.length > 0 && (
                      <button
                        onClick={handleDeselectShortIDsChecklist}
                        className="bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-250 text-xs font-bold px-4.5 py-2.5 rounded-xl transition-all cursor-pointer"
                        title="Unchecks all student records possessing 6 or less digits"
                      >
                        Deselect All Short IDs
                      </button>
                    )}

                    {!pendingImport.isMultiSection && (
                      <button
                        onClick={handleSelectAllChecklist}
                        className="bg-white hover:bg-[#E3EAFA] border border-[#E8ECF5] text-[#252B3B] text-xs font-bold px-4 py-2.5 rounded-xl transition-all cursor-pointer"
                      >
                        Select All ({pendingImport.short.length + pendingImport.normal.length})
                      </button>
                    )}

                    <button
                      onClick={() => setPendingImport(null)}
                      className="bg-[#E3EAFA] hover:bg-[#CED8F2] text-[#252B3B] text-xs font-bold px-4 py-2.5 rounded-xl transition-colors cursor-pointer ml-auto"
                    >
                      Abort
                    </button>
                  </div>
                </div>
              )}

              {/* MATERIAL DESIGN SEGMENTED TIMELINE TABS */}
              <div className="flex bg-[#E3EAFA] p-1 sm:p-1.5 rounded-2xl border border-[#E8ECF5] shrink-0 gap-1 sm:gap-1.5 overflow-x-auto max-w-full md:w-fit shadow-3xs" id="tabs-ribbon">
                <button
                  onClick={() => setActiveTab('students')}
                  className={`flex items-center justify-center gap-1.5 sm:gap-2 px-3 sm:px-5 py-2 sm:py-2.5 text-xs sm:text-sm font-bold rounded-xl transition-all duration-200 select-none cursor-pointer whitespace-nowrap min-h-[42px] ${
                    activeTab === 'students'
                    ? 'bg-[#4D5A8F] text-white shadow-xs'
                    : 'bg-transparent text-[#252B3B] hover:bg-white/80 hover:text-[#252B3B]'
                  }`}
                  id="tab-btn-students"
                >
                  <Users className="h-4 w-4 shrink-0" />
                  <span>Student Roster</span>
                </button>

                <button
                  onClick={() => setActiveTab('attendance')}
                  className={`flex items-center justify-center gap-1.5 sm:gap-2 px-3 sm:px-5 py-2 sm:py-2.5 text-xs sm:text-sm font-bold rounded-xl transition-all duration-200 select-none cursor-pointer whitespace-nowrap min-h-[42px] ${
                    activeTab === 'attendance'
                    ? 'bg-[#4D5A8F] text-white shadow-xs'
                    : 'bg-transparent text-[#252B3B] hover:bg-white/80 hover:text-[#252B3B]'
                  }`}
                  id="tab-btn-attendance"
                >
                  <CalendarCheck className="h-4 w-4 shrink-0" />
                  <span>Attendance Registry</span>
                </button>

                <button
                  onClick={() => setActiveTab('grades')}
                  className={`flex items-center justify-center gap-1.5 sm:gap-2 px-3 sm:px-5 py-2 sm:py-2.5 text-xs sm:text-sm font-bold rounded-xl transition-all duration-200 select-none cursor-pointer whitespace-nowrap min-h-[42px] ${
                    activeTab === 'grades'
                    ? 'bg-[#4D5A8F] text-white shadow-xs'
                    : 'bg-transparent text-[#252B3B] hover:bg-white/80 hover:text-[#252B3B]'
                  }`}
                  id="tab-btn-grades"
                >
                  <GraduationCap className="h-4 w-4 shrink-0" />
                  <span>Academic Grades</span>
                </button>
              </div>

              {/* ACTIVE WORKING VIEWPORTS */}
              <div className="bg-white rounded-2xl border border-[#E8ECF5] p-3.5 sm:p-6 shadow-xs" id="tab-viewport">
                
                {/* STUDENTS VIEWPORT */}
                {activeTab === 'students' && (
                  <div className="flex flex-col gap-4">
                    <div className="flex justify-between items-center pb-2 border-b border-[#E8ECF5] flex-wrap gap-2">
                      <h3 className="font-extrabold text-[#252B3B] text-md tracking-tight">
                        {classes.length === 0 ? "Class Roster: No Active Sections" : `Class Roster Listing: ${activeClass.name}`}
                      </h3>
                      <div className="flex items-center gap-2">
                        <label className="inline-flex items-center gap-1.5 px-3 py-1 bg-white hover:bg-[#E3EAFA] text-[#4D5A8F] border border-[#E8ECF5] hover:border-[#8F9DCC] rounded-lg text-xs font-bold transition-all shadow-3xs cursor-pointer active:scale-95" title="Import students or entire workbook">
                          <Upload className="h-3.5 w-3.5 text-[#4D5A8F]" />
                          <span>Import Spreadsheet</span>
                          <input
                            type="file"
                            accept=".csv, .xlsx, .xls"
                            onChange={handleRosterFileUpload}
                            className="hidden"
                          />
                        </label>
                        <span className="text-xs bg-[#E3EAFA] px-2.5 py-1 text-[#4D5A8F] font-bold rounded-lg border border-[#B3C0E6]/50">
                          Total: {classes.length === 0 ? 0 : activeClass.students.length} students
                        </span>
                      </div>
                    </div>

                    {classes.length === 0 ? (
                      <div className="py-14 flex flex-col items-center justify-center text-center">
                        <Users className="h-12 w-12 text-[#B3C0E6] mb-3" />
                        <h4 className="font-bold text-[#252B3B] text-sm">No Active Sections</h4>
                        <p className="text-xs text-[#747D91] mt-1 max-w-sm leading-normal">
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
                          className="mt-4 inline-flex items-center gap-1.5 px-4 py-2 bg-[#4D5A8F] hover:bg-[#3D4873] text-white rounded-xl text-xs font-bold transition-all shadow-xs active:scale-95 cursor-pointer"
                        >
                          <Plus className="h-4 w-4 text-white" strokeWidth={2.5} />
                          <span>Create New Class</span>
                        </button>
                      </div>
                    ) : activeClass.students.length === 0 ? (
                      <div className="py-12 flex flex-col items-center justify-center text-center">
                        <Users className="h-12 w-12 text-[#B3C0E6] mb-3" />
                        <h4 className="font-bold text-[#252B3B] text-sm">Roster list is currently empty</h4>
                        <p className="text-xs text-[#747D91] mt-1 max-w-sm leading-normal">
                          Create a class section and import a spreadsheet roster to populate enrolled students.
                        </p>
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-[#E8ECF5] rounded-xl shadow-3xs">
                        <table className="w-full text-left border-collapse" id="students-table-matrix">
                          <thead>
                            <tr className="bg-[#E3EAFA] text-[#252B3B] text-[10px] uppercase font-bold tracking-widest border-b border-[#E8ECF5]">
                              <th className="sticky left-0 z-20 bg-[#E3EAFA] pl-3.5 sm:pl-4 pr-1.5 py-3 w-24 sm:w-28 whitespace-nowrap shadow-[1px_0_0_0_#E8ECF5]">ID Number</th>
                              <th className="sticky left-24 sm:left-28 z-20 bg-[#E3EAFA] pl-1.5 pr-3 sm:pr-4 py-3 w-36 sm:w-44 whitespace-nowrap shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]">Student Name</th>
                              <th className="px-4 py-3 whitespace-nowrap">Academic Sneak Peek (Semester Progress)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-[#E8ECF5] text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || 'bg-white';
                              const totalDates = getAttendanceDates(activeClass);
                              const attendedCount = totalDates.filter(d => activeClass.attendance[s.id]?.[d] === 'Attended').length;
                              const finalGrade = calculateFinalGrade(s.id, activeClass);

                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-[#F5F8FF]'} transition-all`}>
                                  <td className={`sticky left-0 z-10 ${rowBgClass} pl-3.5 sm:pl-4 pr-1.5 py-3 whitespace-nowrap font-medium ${colorInfo.border || ''} shadow-[1px_0_0_0_#E8ECF5]`}>
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
                                      <span className="font-mono text-xs font-bold text-[#252B3B] bg-[#F5F8FF] px-2 py-0.5 rounded border border-[#E8ECF5] shadow-3xs">
                                        {s.id}
                                      </span>
                                    </div>
                                  </td>
                                  <td className={`sticky left-24 sm:left-28 z-10 ${rowBgClass} pl-1.5 pr-3 sm:pr-4 py-3 font-bold text-[#252B3B] whitespace-nowrap shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]`}>
                                    <span>{s.fullName}</span>
                                  </td>
                                  <td className={`px-4 py-3 ${rowBgClass}`}>
                                    <div className="flex items-center gap-2 flex-wrap">
                                      {/* Attendance days attended / total lessons given (e.g. 12/16) */}
                                      <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold bg-[#E3EAFA] text-[#4D5A8F] border border-[#B3C0E6]/50 shadow-3xs">
                                        <CalendarCheck className="h-3.5 w-3.5 text-[#4D5A8F] shrink-0" />
                                        <span>
                                          Attendance: <strong className="font-bold text-[#252B3B] font-mono">({attendedCount}/{totalDates.length})</strong>
                                        </span>
                                        {totalDates.length > 0 && (
                                          <span className="text-[10px] text-[#4D5A8F] font-mono font-bold">
                                            {Math.round((attendedCount / totalDates.length) * 100)}%
                                          </span>
                                        )}
                                      </div>

                                      {/* Assessment names with score next to it */}
                                      {activeClass.assessments.map(ass => {
                                        const score = activeClass.grades[s.id]?.[ass.name];
                                        const hasScore = score !== undefined && score !== null;
                                        return (
                                          <div key={ass.name} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs bg-white border border-[#E8ECF5] shadow-3xs">
                                            <span className="text-[#747D91] font-medium text-[11px] truncate max-w-28">{ass.name}:</span>
                                            <span className={`font-mono text-[11px] ${hasScore ? 'text-[#252B3B] font-extrabold' : 'text-[#747D91] font-normal italic'}`}>
                                              {hasScore ? score : '0'}
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
                                        <span className="text-xs text-[#747D91] italic">
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
                    <div className="pb-3 border-b border-[#E8ECF5]">
                      <h3 className="font-extrabold text-[#252B3B] text-md tracking-tight">Daily Attendance Registry</h3>
                      <p className="text-xs text-[#747D91] mt-1">
                        Define dates and track student presence. Column values represent custom styled markers matching status colors.
                      </p>
                    </div>

                    {/* Choose and configure date columns */}
                    <div className="p-4 bg-[#F5F8FF] rounded-2xl border border-[#E8ECF5] grid md:grid-cols-3 gap-4 items-center">
                      <div className="flex flex-col gap-1">
                        <label htmlFor="attendance-date-picker" className="text-[10px] font-bold text-[#4D5A8F] uppercase tracking-widest">
                          Attendance date
                        </label>
                        <DDMMYYDatePicker
                          id="attendance-date-picker"
                          value={attendanceDate}
                          onChange={setAttendanceDate}
                        />
                      </div>
                      <div className="md:col-span-2 pt-4 md:pt-0 flex items-center gap-2 flex-wrap">
                        <button
                          onClick={handleAddAttendanceColumn}
                          className="bg-[#4D5A8F] hover:bg-[#3D4873] text-white font-bold text-xs px-4 py-2.5 rounded-xl shadow-xs transition-all cursor-pointer shrink-0"
                          id="take-attendance-trigger-btn"
                        >
                          Take Attendance (Add Column)
                        </button>
                        <label className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-white hover:bg-[#E3EAFA] text-[#4D5A8F] border border-[#E8ECF5] hover:border-[#8F9DCC] rounded-xl text-xs font-bold transition-all shadow-3xs cursor-pointer active:scale-95 shrink-0" title="Import attendance log file for this class or all classes">
                          <Upload className="h-3.5 w-3.5 text-[#4D5A8F]" />
                          <span>Import Attendance Log</span>
                          <input
                            type="file"
                            accept=".csv, .xlsx, .xls"
                            onChange={handleRosterFileUpload}
                            className="hidden"
                          />
                        </label>
                      </div>
                    </div>

                    {takeAttendanceMsg && (
                      <div className={`text-xs p-3 rounded-xl border flex items-center gap-2 ${
                        takeAttendanceMsg.type === 'success' 
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
                        : 'bg-amber-50 text-amber-800 border-amber-200'
                      }`}>
                        <Check className="h-4 w-4 shrink-0 text-emerald-600" />
                        <span>{takeAttendanceMsg.text}</span>
                      </div>
                    )}

                    {classes.length === 0 ? (
                      <div className="py-12 text-center text-[#747D91] italic">
                        No active class sections. Please create a section to take attendance.
                      </div>
                    ) : activeClass.students.length === 0 ? (
                      <div className="py-12 text-center text-[#747D91] italic">
                        No registered student lists detected. Setup student information first.
                      </div>
                    ) : getAttendanceDates(activeClass).length === 0 ? (
                      <div className="p-8 text-center bg-[#F5F8FF] border border-dashed border-[#B3C0E6] rounded-2xl text-[#747D91]">
                        No attendance date logs found. Pick an attendance date above and click 'Take Attendance' to initiate your registry.
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-[#E8ECF5] rounded-xl shadow-3xs">
                        <table className="w-full text-left border-collapse" id="attendance-table-matrix">
                          <thead>
                            <tr className="bg-[#E3EAFA] text-[#252B3B] text-[10px] uppercase font-bold tracking-widest border-b border-[#E8ECF5]">
                              <th className="sticky left-0 z-20 bg-[#E3EAFA] pl-3.5 sm:pl-4 pr-1.5 py-3.5 whitespace-nowrap w-24 sm:w-28 shadow-[1px_0_0_0_#E8ECF5]">ID Number</th>
                              <th className="sticky left-24 sm:left-28 z-20 bg-[#E3EAFA] pl-1.5 pr-3 sm:pr-4 py-3.5 whitespace-nowrap w-36 sm:w-44 shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]">Full Name</th>
                              {getAttendanceDates(activeClass).map(d => (
                                <th key={d} className="px-3 sm:px-4 py-3.5 whitespace-nowrap text-center">
                                  <div className="inline-flex items-center gap-1.5 justify-center">
                                    <span className="font-mono text-xs font-bold text-[#252B3B]">
                                      {formatDisplayDate(d)}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={() => setConfirmDeleteDateTarget({ date: d, displayDate: formatDisplayDate(d) })}
                                      title={`Delete attendance column for ${formatDisplayDate(d)}`}
                                      className="p-1 rounded text-[#747D91] hover:text-rose-600 hover:bg-rose-50 transition-colors cursor-pointer"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </div>
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-[#E8ECF5] text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || 'bg-white';
                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-[#F5F8FF]'} transition-all`}>
                                  <td className={`sticky left-0 z-10 ${rowBgClass} pl-3.5 sm:pl-4 pr-1.5 py-3.5 text-[#747D91] font-mono font-medium ${colorInfo.border || ''} shadow-[1px_0_0_0_#E8ECF5]`}>{s.id}</td>
                                  <td className={`sticky left-24 sm:left-28 z-10 ${rowBgClass} pl-1.5 pr-3 sm:pr-4 py-3.5 font-bold text-[#252B3B] whitespace-nowrap shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]`}>{s.fullName}</td>
                                  {getAttendanceDates(activeClass).map(d => {
                                    const status = activeClass.attendance[s.id]?.[d] || 'Not Taken';
                                    
                                    let pillClass = 'bg-[#F5F8FF] text-[#747D91] border-[#E8ECF5]';
                                    if (status === 'Attended') pillClass = 'bg-emerald-50 text-emerald-700 border-emerald-200';
                                    else if (status === 'Absent') pillClass = 'bg-rose-50 text-rose-700 border-rose-200';

                                    return (
                                      <td key={d} className={`px-3 sm:px-4 py-3 text-center whitespace-nowrap ${rowBgClass}`}>
                                        <button
                                          onClick={() => handleToggleAttendanceCell(s.id, d, status)}
                                          className={`inline-flex items-center justify-center min-w-[76px] min-h-[32px] px-3 py-1 rounded-full text-[11px] font-bold border transition-all cursor-pointer active:scale-95 ${pillClass}`}
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
                    <div className="pb-3 border-b border-[#E8ECF5]">
                      <h3 className="font-extrabold text-[#252B3B] text-md tracking-tight">Academic Grading Sheets</h3>
                      <p className="text-xs text-[#747D91] mt-1">
                        Register dynamic assessments and record scores. Custom Excel exporter highlights whitelists and Final Grade calculations automatically.
                      </p>
                    </div>

                    {/* Action Toolbar: Clean button with plus sign to Add Assessment & Total Weight breakdown */}
                    <div className="flex flex-wrap items-center justify-between gap-3 p-3.5 bg-[#F5F8FF] rounded-2xl border border-[#E8ECF5]">
                      <div className="flex items-center gap-3">
                        <button
                          type="button"
                          onClick={handleOpenAddAssessment}
                          className="inline-flex items-center gap-2 bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-95 text-white font-bold text-xs px-4 py-2.5 rounded-xl shadow-xs transition-all cursor-pointer"
                          id="open-add-assessment-btn"
                        >
                          <Plus className="h-4 w-4" />
                          <span>Add Assessment</span>
                        </button>
                        <label className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-white hover:bg-[#E3EAFA] text-[#4D5A8F] border border-[#E8ECF5] hover:border-[#8F9DCC] rounded-xl text-xs font-bold transition-all shadow-3xs cursor-pointer active:scale-95" title="Import grades report file for this class or all classes">
                          <Upload className="h-3.5 w-3.5 text-[#4D5A8F]" />
                          <span>Import Grades Sheet</span>
                          <input
                            type="file"
                            accept=".csv, .xlsx, .xls"
                            onChange={handleRosterFileUpload}
                            className="hidden"
                          />
                        </label>
                        <span className="text-xs text-[#747D91] hidden sm:inline font-medium">
                          Click to specify assessment name, full mark, and weight.
                        </span>
                      </div>

                      <button
                        type="button"
                        onClick={handleOpenFinalGradeWeights}
                        className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-bold border transition-colors cursor-pointer bg-white hover:bg-[#E3EAFA] border-[#E8ECF5] text-[#252B3B] shadow-3xs"
                        title="Click to view and adjust all assessment weights"
                        id="final-grade-weights-breakdown-btn"
                      >
                        <Settings className="h-3.5 w-3.5 text-[#4D5A8F]" />
                        <span>Weight Sum:</span>
                        <span className={`font-mono px-2 py-0.5 rounded text-[11px] font-extrabold ${
                          totalAllocatedWeight === 100 
                            ? 'bg-emerald-50 text-emerald-800' 
                            : totalAllocatedWeight < 100 
                              ? 'bg-amber-50 text-amber-800' 
                              : 'bg-rose-50 text-rose-800'
                        }`}>
                          {totalAllocatedWeight}% / 100%
                        </span>
                      </button>
                    </div>

                    {assessFeedback && (
                      <div className={`text-xs p-3 rounded-xl border ${
                        assessFeedback.type === 'success' 
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
                        : 'bg-rose-50 text-rose-800 border-rose-200'
                      }`}>
                        {assessFeedback.text}
                      </div>
                    )}

                    {classes.length === 0 ? (
                      <div className="py-12 text-center text-[#747D91] italic">
                        No active class sections. Please create a section to record grades.
                      </div>
                    ) : activeClass.students.length === 0 ? (
                      <div className="py-12 text-center text-[#747D91] italic">
                        No active roster exists. Create a class with students or import a roster to begin recording grades.
                      </div>
                    ) : activeClass.assessments.length === 0 ? (
                      <div className="p-8 text-center bg-[#F5F8FF] border border-dashed border-[#B3C0E6] rounded-2xl text-[#747D91] flex flex-col items-center gap-2">
                        <p>No assessments registered yet. Click 'Add Assessment' above to create test columns.</p>
                        <button
                          type="button"
                          onClick={handleOpenAddAssessment}
                          className="mt-1 inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold bg-[#4D5A8F] hover:bg-[#3D4873] text-white rounded-xl shadow-xs transition-colors cursor-pointer"
                        >
                          <Plus className="h-3.5 w-3.5" />
                          <span>Add Assessment</span>
                        </button>
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-[#E8ECF5] rounded-xl shadow-3xs">
                        <table className="w-full text-left border-collapse" id="grades-table-matrix">
                          <thead>
                            <tr className="bg-[#E3EAFA] text-[#252B3B] text-[10px] uppercase font-bold tracking-widest border-b border-[#E8ECF5]">
                              <th className="sticky left-0 z-20 bg-[#E3EAFA] pl-3.5 sm:pl-4 pr-1.5 py-3.5 whitespace-nowrap w-24 sm:w-28 shadow-[1px_0_0_0_#E8ECF5]">ID Number</th>
                              <th className="sticky left-24 sm:left-28 z-20 bg-[#E3EAFA] pl-1.5 pr-3 sm:pr-4 py-3.5 whitespace-nowrap w-36 sm:w-44 shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]">Full Name</th>
                              {activeClass.assessments.map(a => (
                                <th 
                                  key={a.name} 
                                  onClick={() => handleOpenEditAssessment(a)}
                                  className="px-4 py-3 whitespace-nowrap cursor-pointer hover:bg-[#CED8F2] transition-colors group select-none text-[#252B3B]"
                                  title="Click to edit test name, max mark, or weight"
                                >
                                  <div className="flex flex-col items-start gap-0.5">
                                    <div className="flex items-center gap-1.5 font-bold text-[#252B3B] text-xs">
                                      <span>{a.name}</span>
                                      <Settings className="h-3 w-3 text-[#4D5A8F] opacity-60 group-hover:opacity-100 transition-opacity" />
                                    </div>
                                    <div className="flex items-center gap-1.5 text-[10px] text-[#747D91] font-normal">
                                      <span>Full Mark: <strong className="font-semibold text-[#252B3B]">{a.maxMark || 100}</strong></span>
                                      <span>•</span>
                                      <span>Weight: <strong className="font-semibold text-[#252B3B]">{a.weight}%</strong></span>
                                    </div>
                                  </div>
                                </th>
                              ))}
                              <th 
                                onClick={handleOpenFinalGradeWeights}
                                className="px-5 py-3.5 whitespace-nowrap text-right text-emerald-900 font-extrabold bg-emerald-50 hover:bg-emerald-100/70 transition-colors cursor-pointer select-none group border-l border-[#E8ECF5]"
                                title="Click to view and adjust all assessment weights"
                              >
                                <div className="flex flex-col items-end gap-0.5">
                                  <div className="inline-flex items-center gap-1.5 justify-end">
                                    <span>📊 Final Grade</span>
                                    <Settings className="h-3.5 w-3.5 text-emerald-800 group-hover:rotate-45 transition-transform" />
                                  </div>
                                  <span className="text-[10px] text-emerald-800 font-bold uppercase tracking-wider">
                                    Weight: {totalAllocatedWeight}% / 100%
                                  </span>
                                </div>
                              </th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-[#E8ECF5] text-xs">
                            {activeClass.students.map(s => {
                              const colorInfo = getStudentHighlightClasses(s.color);
                              const rowBgClass = colorInfo.bg || 'bg-white';
                              const finalGradeCellClass = s.color && s.color !== 'none' ? rowBgClass : 'bg-[#E3EAFA]/30';
                              return (
                                <tr key={s.id} className={`${rowBgClass} ${colorInfo.hover || 'hover:bg-[#F5F8FF]'} transition-all`}>
                                  <td className={`sticky left-0 z-10 ${rowBgClass} pl-3.5 sm:pl-4 pr-1.5 py-3.5 text-[#747D91] font-mono font-medium ${colorInfo.border || ''} shadow-[1px_0_0_0_#E8ECF5]`}>{s.id}</td>
                                  <td className={`sticky left-24 sm:left-28 z-10 ${rowBgClass} pl-1.5 pr-3 sm:pr-4 py-3.5 font-bold text-[#252B3B] whitespace-nowrap shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]`}>{s.fullName}</td>
                                  {activeClass.assessments.map(a => {
                                    const score = activeClass.grades[s.id]?.[a.name];
                                    const maxMark = a.maxMark || 100;
                                    return (
                                      <td key={a.name} className={`px-4 py-3.5 whitespace-nowrap ${rowBgClass}`}>
                                        <div className="flex items-center gap-1">
                                          <GradeScoreInput 
                                            score={score}
                                            maxMark={maxMark}
                                            highlightColor={s.color}
                                            onSave={val => handleScoreChange(s.id, a.name, val)}
                                          />
                                          <span className="text-[10px] text-[#747D91] font-mono">/{maxMark}</span>
                                        </div>
                                      </td>
                                    );
                                  })}
                                  <td className={`px-5 py-3.5 text-right font-black text-emerald-950 whitespace-nowrap ${finalGradeCellClass}`}>
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

      {/* GOOGLE MATERIAL DESIGN CLASS CREATION MODAL WITH ISOLATED INTEGRATED ROSTER IMPORT */}
      {isCreatingClass && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-fade-in" id="class-creation-modal-overlay">
          <div className="bg-white rounded-3xl p-6 md:p-8 max-w-lg w-full shadow-2xl flex flex-col gap-5 border border-[#E8ECF5] max-h-[90vh] overflow-y-auto transform scale-100 transition-all duration-300">
            
            <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-3.5 animate-none" id="modal-header">
              <span className="text-base font-bold text-[#252B3B] tracking-tight flex items-center gap-1.5 font-sans">
                <Plus className="h-5 w-5 text-[#4D5A8F]" />
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
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] border border-[#E8ECF5] transition-all cursor-pointer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleCreateClass} className="flex flex-col gap-4">
              
              {/* Class Name Input */}
              <div className="flex flex-col gap-1">
                <label className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider">Class Section Name</label>
                <input
                  type="text"
                  value={newClassNameInput}
                  onChange={e => setNewClassNameInput(e.target.value)}
                  placeholder="e.g. Physics 301, Section B"
                  className="w-full text-sm px-3.5 py-2.5 rounded-xl border border-[#E8ECF5] bg-white shadow-3xs focus:outline-none focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] transition-all"
                  autoFocus
                  required
                />
              </div>

              {/* Roster File Import Box */}
              <div className="flex flex-col gap-2 border-t border-[#E8ECF5] pt-4" id="modal-roster-uploader">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider flex items-center gap-1.5">
                    <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                    Student Roster (Optional Spreadsheet)
                  </label>
                </div>
                <p className="text-[11px] text-[#747D91] leading-normal">
                  Upload a spreadsheet roster (.csv or .xlsx) to automatically populate student names and IDs.
                </p>

                <div 
                  onDragOver={handleModalRosterDragOver}
                  onDragLeave={handleModalRosterDragLeave}
                  onDrop={handleModalRosterDrop}
                  className={`border-2 border-dashed rounded-2xl p-5 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${
                    isDragOver 
                      ? 'border-[#4D5A8F] bg-[#E3EAFA]/60 scale-[1.01]' 
                      : 'border-[#E8ECF5] hover:border-[#8F9DCC] bg-[#F5F8FF] hover:bg-[#E3EAFA]/30'
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
                    <Upload className="h-7 w-7 text-[#8F9DCC] mb-2" />
                    <span className="text-xs font-bold text-[#252B3B]">Choose file or drag and drop here</span>
                    <span className="text-[10px] text-[#747D91] mt-1">Accepts CSV or Excel workbooks</span>
                  </label>
                </div>

                {/* Feedback state notifications inside modal */}
                {createdClassFileError && (
                  <div className="text-[10.5px] bg-rose-50 text-rose-700 p-2.5 rounded-xl border border-rose-200 flex items-start gap-1.5 mt-1.5">
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
                  <div className="mt-1.5 p-2.5 bg-[#E3EAFA]/40 rounded-xl border border-[#B3C0E6]/50 text-[11px] text-[#252B3B] overflow-hidden text-left">
                    <div className="flex items-center justify-between border-b border-[#B3C0E6]/40 pb-1 mb-1.5">
                      <span className="font-bold text-[#252B3B] block">Import Preview ({createdClassStudents.length}):</span>
                      <button
                        type="button"
                        onClick={() => setCreatedClassStudents([])}
                        className="text-[10px] text-rose-500 hover:text-rose-700 font-bold hover:underline cursor-pointer"
                      >
                        Clear All
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto p-1 bg-white/60 rounded-lg">
                      {createdClassStudents.map(st => {
                        const isShort = st.id.length <= 6;
                        return (
                          <span 
                            key={st.id} 
                            className={`inline-flex items-center gap-1 bg-white border px-2 py-0.5 rounded-full text-[10px] font-semibold text-[#252B3B] font-mono shadow-3xs hover:bg-[#F5F8FF] transition-colors ${
                              isShort ? 'border-amber-300 bg-amber-50/50 text-amber-900' : 'border-[#E8ECF5]'
                            }`}
                          >
                            <span>{st.firstName} ({st.id})</span>
                            <button
                              type="button"
                              onClick={() => {
                                setCreatedClassStudents(prev => prev.filter(x => x.id !== st.id));
                              }}
                              className="p-0.5 rounded-full hover:bg-[#E3EAFA] text-[#747D91] hover:text-[#252B3B] flex items-center justify-center cursor-pointer transition-colors"
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

              {/* Optional Lecture Schedule Feature */}
              <div className="flex flex-col gap-2.5 border-t border-[#E8ECF5] pt-4" id="modal-lecture-schedule">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider flex items-center gap-1.5">
                    <Clock className="h-4 w-4 text-[#4D5A8F]" />
                    Lecture Schedule / Timetable (Optional)
                  </label>
                  <span className="text-[10px] text-[#747D91] font-medium">Adds to weekly timetable</span>
                </div>
                <p className="text-[11px] text-[#747D91] leading-normal">
                  Choose the day of the week and enter the lecture time period (__am/pm to __am/pm).
                </p>

                <div className="bg-[#F5F8FF] p-3.5 rounded-2xl border border-[#E8ECF5] flex flex-col gap-3">
                  <div className="flex flex-col gap-1.5">
                    <div className="flex items-center justify-between">
                      <label className="text-[10.5px] font-bold text-[#252B3B]">
                        Select Lecture Days (Choose Multiple)
                      </label>
                      <div className="flex items-center gap-1.5 text-[10px]">
                        <button
                          type="button"
                          onClick={() => setNewClassDays(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])}
                          className="text-[#4D5A8F] font-bold hover:underline cursor-pointer"
                        >
                          Mon–Fri
                        </button>
                        <span className="text-[#8F9DCC]">•</span>
                        <button
                          type="button"
                          onClick={() => setNewClassDays([...DAYS_OF_WEEK])}
                          className="text-[#4D5A8F] font-bold hover:underline cursor-pointer"
                        >
                          All 7 Days
                        </button>
                        <span className="text-[#8F9DCC]">•</span>
                        <button
                          type="button"
                          onClick={() => setNewClassDays([])}
                          className="text-[#747D91] hover:text-rose-600 font-medium cursor-pointer"
                        >
                          Clear
                        </button>
                      </div>
                    </div>

                    {/* Day multi-selection buttons */}
                    <div className="grid grid-cols-7 gap-1">
                      {DAYS_OF_WEEK.map(d => {
                        const isSelected = newClassDays.includes(d);
                        const shortName = d.slice(0, 3);
                        return (
                          <button
                            key={d}
                            type="button"
                            onClick={() => {
                              if (isSelected) {
                                setNewClassDays(prev => prev.filter(x => x !== d));
                              } else {
                                setNewClassDays(prev => [...prev, d]);
                              }
                            }}
                            className={`py-2 px-1 rounded-xl text-xs font-bold transition-all border text-center cursor-pointer ${
                              isSelected
                                ? 'bg-[#4D5A8F] text-white border-[#4D5A8F] shadow-3xs scale-100 font-black'
                                : 'bg-white text-[#252B3B] border-[#E8ECF5] hover:bg-[#E3EAFA]'
                            }`}
                            title={d}
                          >
                            <span className="block">{shortName}</span>
                          </button>
                        );
                      })}
                    </div>

                    {newClassDays.length > 0 ? (
                      <span className="text-[10.5px] text-[#4D5A8F] font-semibold">
                        ✓ {newClassDays.length} {newClassDays.length === 1 ? 'day' : 'days'} selected: {newClassDays.join(', ')}
                      </span>
                    ) : (
                      <span className="text-[10px] text-[#747D91] italic">
                        Select one or multiple days to schedule lectures (optional).
                      </span>
                    )}
                  </div>

                  {newClassDays.length > 0 && (
                    <div className="flex flex-col gap-2 pt-2 border-t border-[#E8ECF5] animate-in fade-in duration-150">
                      <label className="text-[10.5px] font-bold text-[#252B3B] flex items-center justify-between">
                        <span>Lecture Time Period for Selected Days</span>
                        <span className="text-[10.5px] text-[#4D5A8F] font-mono font-bold">
                          {newClassStartHour} {newClassStartPeriod} to {newClassEndHour} {newClassEndPeriod}
                        </span>
                      </label>

                      <div className="grid grid-cols-2 gap-2">
                        {/* Start Time */}
                        <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-[#E8ECF5] shadow-3xs">
                          <span className="text-[9.5px] font-bold text-[#747D91] uppercase tracking-wider">Start Time</span>
                          <div className="flex items-center gap-1">
                            <input
                              type="text"
                              value={newClassStartHour}
                              onChange={e => setNewClassStartHour(e.target.value)}
                              placeholder="09:00"
                              className="w-full text-xs font-mono font-bold text-[#252B3B] outline-none"
                            />
                            <button
                              type="button"
                              onClick={() => setNewClassStartPeriod(p => p === 'AM' ? 'PM' : 'AM')}
                              className="px-2 py-0.5 text-[10px] font-black rounded-md bg-[#E3EAFA] text-[#4D5A8F] hover:bg-[#CED8F2] transition-colors cursor-pointer"
                            >
                              {newClassStartPeriod}
                            </button>
                          </div>
                        </div>

                        {/* End Time */}
                        <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-[#E8ECF5] shadow-3xs">
                          <span className="text-[9.5px] font-bold text-[#747D91] uppercase tracking-wider">End Time</span>
                          <div className="flex items-center gap-1">
                            <input
                              type="text"
                              value={newClassEndHour}
                              onChange={e => setNewClassEndHour(e.target.value)}
                              placeholder="10:30"
                              className="w-full text-xs font-mono font-bold text-[#252B3B] outline-none"
                            />
                            <button
                              type="button"
                              onClick={() => setNewClassEndPeriod(p => p === 'AM' ? 'PM' : 'AM')}
                              className="px-2 py-0.5 text-[10px] font-black rounded-md bg-[#E3EAFA] text-[#4D5A8F] hover:bg-[#CED8F2] transition-colors cursor-pointer"
                            >
                              {newClassEndPeriod}
                            </button>
                          </div>
                        </div>
                      </div>

                      {/* Quick presets */}
                      <div className="flex items-center gap-1.5 flex-wrap pt-1">
                        <span className="text-[9.5px] text-[#747D91] font-medium">Quick presets:</span>
                        {[
                          { start: '08:30', sP: 'AM', end: '10:00', eP: 'AM' },
                          { start: '08:30', sP: 'AM', end: '11:20', eP: 'AM' },
                          { start: '09:00', sP: 'AM', end: '10:30', eP: 'AM' },
                          { start: '10:30', sP: 'AM', end: '12:00', eP: 'PM' },
                          { start: '01:00', sP: 'PM', end: '02:00', eP: 'PM' },
                          { start: '01:00', sP: 'PM', end: '02:30', eP: 'PM' },
                          { start: '02:30', sP: 'PM', end: '04:00', eP: 'PM' }
                        ].map((preset, idx) => (
                          <button
                            key={idx}
                            type="button"
                            onClick={() => {
                              setNewClassStartHour(preset.start);
                              setNewClassStartPeriod(preset.sP as 'AM' | 'PM');
                              setNewClassEndHour(preset.end);
                              setNewClassEndPeriod(preset.eP as 'AM' | 'PM');
                            }}
                            className="text-[9.5px] font-mono px-1.5 py-0.5 rounded bg-white hover:bg-[#E3EAFA] hover:text-[#4D5A8F] text-[#252B3B] transition-colors cursor-pointer border border-[#E8ECF5]"
                          >
                            {preset.start}{preset.sP}-{preset.end}{preset.eP}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* Action buttons */}
              <div className="flex flex-col-reverse sm:flex-row items-center gap-2 sm:gap-2.5 justify-end mt-4 border-t border-[#E8ECF5] pt-4" id="modal-actions">
                <button
                  type="button"
                  onClick={() => {
                    setIsCreatingClass(false);
                    setNewClassNameInput('');
                    setCreatedClassStudents([]);
                    setCreatedClassFileError(null);
                    setCreatedClassFileSuccess(null);
                    setNewClassDays([]);
                  }}
                  className="w-full sm:w-auto text-center justify-center bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] border border-[#E8ECF5] text-xs font-bold px-4.5 py-2.5 rounded-xl transition-all cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="w-full sm:w-auto text-center justify-center bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-[0.98] text-white text-xs font-bold px-5 py-2.5 rounded-xl transition-all cursor-pointer shadow-xs inline-flex items-center gap-1.5"
                >
                  <Plus className="h-4 w-4" />
                  Create Class
                </button>
              </div>

            </form>
          </div>
        </div>
      )}

      {/* EDIT CLASS NAME, STUDENT ROSTER & SCHEDULE MODAL */}
      {editingClassTarget && (() => {
        const targetClassData = classes.find(c => c.name === editingClassTarget.originalName);
        const currentEnrolledStudents = targetClassData?.students || [];

        return (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-fade-in" id="edit-class-modal-overlay">
            <div className="bg-white rounded-3xl p-6 md:p-8 max-w-xl w-full shadow-2xl flex flex-col gap-5 border border-[#E8ECF5] max-h-[92vh] overflow-hidden transform scale-100 transition-all duration-300">
              {/* Header */}
              <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-3.5">
                <div className="flex items-center gap-2.5">
                  <div className="w-8 h-8 rounded-xl bg-[#E3EAFA] flex items-center justify-center text-[#4D5A8F]">
                    <Pencil className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-base font-bold text-[#252B3B] tracking-tight">
                      Edit Class Section
                    </h3>
                    <p className="text-[11px] text-[#747D91]">
                      Manage section details, student roster, and timetable schedule
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setEditingClassTarget(null)}
                  className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] border border-[#E8ECF5] transition-all cursor-pointer"
                >
                  <X className="h-4 w-4" />
                </button>
              </div>

              {/* Navigation Tabs (Clear & Uncluttered) */}
              <div className="flex items-center gap-2 bg-[#E3EAFA] p-1 rounded-2xl border border-[#E8ECF5] text-xs font-bold">
                <button
                  type="button"
                  onClick={() => setEditClassActiveTab('roster')}
                  className={`flex-1 py-2 px-3 rounded-xl transition-all cursor-pointer flex items-center justify-center gap-2 ${
                    editClassActiveTab === 'roster'
                      ? 'bg-[#4D5A8F] text-white shadow-3xs'
                      : 'text-[#252B3B] hover:bg-white/80'
                  }`}
                >
                  <Users className="h-3.5 w-3.5" />
                  <span>Class &amp; Student Roster</span>
                  <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-mono ${
                    editClassActiveTab === 'roster' ? 'bg-white/20 text-white' : 'bg-white text-[#4D5A8F] border border-[#B3C0E6]/50'
                  }`}>
                    {currentEnrolledStudents.length}
                  </span>
                </button>
                <button
                  type="button"
                  onClick={() => setEditClassActiveTab('schedule')}
                  className={`flex-1 py-2 px-3 rounded-xl transition-all cursor-pointer flex items-center justify-center gap-2 ${
                    editClassActiveTab === 'schedule'
                      ? 'bg-[#4D5A8F] text-white shadow-3xs'
                      : 'text-[#252B3B] hover:bg-white/80'
                  }`}
                >
                  <Clock className="h-3.5 w-3.5" />
                  <span>Weekly Schedule</span>
                  <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-mono ${
                    editClassActiveTab === 'schedule' ? 'bg-white/20 text-white' : 'bg-white text-[#4D5A8F] border border-[#B3C0E6]/50'
                  }`}>
                    {editingClassTarget.schedule.length}
                  </span>
                </button>
              </div>

              {/* Form Content Scrollable */}
              <form onSubmit={handleSaveEditedClass} className="flex flex-col gap-4 overflow-y-auto pr-1">
                {editClassActiveTab === 'roster' ? (
                  <div className="flex flex-col gap-4">
                    {/* Class Name Input */}
                    <div className="flex flex-col gap-1.5 bg-[#F5F8FF] p-3.5 rounded-2xl border border-[#E8ECF5]">
                      <label className="text-[10px] font-bold text-[#252B3B] uppercase tracking-wider">
                        Class Section Name
                      </label>
                      <input
                        type="text"
                        value={editingClassTarget.newName}
                        onChange={e => {
                          setEditingClassTarget({ ...editingClassTarget, newName: e.target.value });
                          setEditingClassModalError(null);
                        }}
                        className="w-full text-sm font-bold px-3.5 py-2.5 rounded-xl border border-[#E8ECF5] bg-white shadow-3xs focus:outline-none focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] transition-all"
                        placeholder="e.g. Physics 101"
                        required
                      />
                      {editingClassModalError && (
                        <span className="text-[11px] text-rose-600 font-semibold mt-1">
                          {editingClassModalError}
                        </span>
                      )}
                    </div>

                    {/* Student Roster Section */}
                    <div className="flex flex-col gap-2.5">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-1.5">
                          <Users className="h-4 w-4 text-[#4D5A8F]" />
                          <span className="text-xs font-bold text-[#252B3B] uppercase tracking-wider">
                            Enrolled Student Roster
                          </span>
                          <span className="text-[10px] font-mono font-bold bg-[#E3EAFA] text-[#4D5A8F] px-2 py-0.5 rounded-md border border-[#B3C0E6]/50">
                            {currentEnrolledStudents.length} {currentEnrolledStudents.length === 1 ? 'student' : 'students'}
                          </span>
                        </div>

                        {/* ADD A STUDENT BUTTON */}
                        <button
                          type="button"
                          onClick={() => {
                            setAddStudentIdInput('');
                            setAddStudentNameInput('');
                            setAddStudentModalError(null);
                            setIsAddStudentModalOpen(true);
                          }}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-[#4D5A8F] hover:bg-[#3D4873] text-white text-xs font-bold rounded-xl transition-all shadow-3xs cursor-pointer active:scale-95"
                          id="btn-add-student-modal-trigger"
                        >
                          <Plus className="h-3.5 w-3.5" />
                          <span>Add a Student</span>
                        </button>
                      </div>

                      {addStudentSuccessMsg && (
                        <div className="p-2.5 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-bold flex items-center gap-2 animate-fade-in">
                          <Check className="h-4 w-4 text-emerald-600 shrink-0" />
                          <span>{addStudentSuccessMsg}</span>
                        </div>
                      )}

                      {/* Enrolled Students List Card */}
                      <div className="border border-[#E8ECF5] rounded-2xl bg-white overflow-hidden shadow-3xs">
                        {currentEnrolledStudents.length === 0 ? (
                          <div className="p-6 text-center flex flex-col items-center justify-center gap-1.5">
                            <Users className="h-8 w-8 text-[#B3C0E6]" />
                            <p className="text-xs font-bold text-[#252B3B]">No students enrolled yet</p>
                            <p className="text-[11px] text-[#747D91]">
                              Click &quot;Add a Student&quot; above to enroll students into this class section.
                            </p>
                          </div>
                        ) : (
                          <div className="max-h-56 overflow-y-auto divide-y divide-[#E8ECF5]">
                            {currentEnrolledStudents.map(st => {
                              const colorInfo = getStudentHighlightClasses(st.color);
                              return (
                                <div key={st.id} className="p-2.5 px-3 flex items-center justify-between hover:bg-[#F5F8FF] transition-colors text-xs">
                                  <div className="flex items-center gap-2.5 min-w-0">
                                    <span className="font-mono bg-[#E3EAFA] text-[#4D5A8F] px-2 py-0.5 rounded font-bold text-[11px] shrink-0 border border-[#B3C0E6]/50">
                                      ID: {st.id}
                                    </span>
                                    <span className="font-bold text-[#252B3B] truncate">
                                      {st.fullName}
                                    </span>
                                    {st.color && st.color !== 'none' && (
                                      <span className={`w-2 h-2 rounded-full ${colorInfo.border?.replace('border-l-4 ', '') || 'bg-slate-400'} shrink-0`} title={`Color: ${st.color}`} />
                                    )}
                                  </div>

                                  <button
                                    type="button"
                                    onClick={() => {
                                      setConfirmDeleteStudentTarget({
                                        studentId: st.id,
                                        studentName: st.fullName,
                                        className: editingClassTarget.originalName
                                      });
                                    }}
                                    className="p-1 text-[#747D91] hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer shrink-0"
                                    title={`Remove ${st.fullName} from this class`}
                                  >
                                    <Trash2 className="h-3.5 w-3.5" />
                                  </button>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="flex flex-col gap-3.5">
                    {/* Current slots list */}
                    <div className="flex flex-col gap-1.5">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-bold text-[#252B3B] uppercase tracking-wider">
                          Current Timetable Slots ({editingClassTarget.schedule.length})
                        </span>
                        {editingClassTarget.schedule.length > 0 && (
                          <button
                            type="button"
                            onClick={() => setEditingClassTarget({ ...editingClassTarget, schedule: [] })}
                            className="text-[10px] text-rose-600 hover:underline font-bold cursor-pointer"
                          >
                            Clear All Slots
                          </button>
                        )}
                      </div>

                      {editingClassTarget.schedule.length === 0 ? (
                        <p className="text-xs text-[#747D91] italic bg-[#F5F8FF] p-3 rounded-xl border border-dashed border-[#B3C0E6] text-center">
                          No lecture time slots scheduled yet for this section.
                        </p>
                      ) : (
                        <div className="flex flex-col gap-1.5 max-h-40 overflow-y-auto">
                          {editingClassTarget.schedule.map((slot, sIdx) => {
                            const startM = parseTimeToMinutes(slot.startTime);
                            const endM = parseTimeToMinutes(slot.endTime);
                            const durStr = endM > startM ? formatDurationMinutes(endM - startM) : '';
                            return (
                              <div key={slot.id || sIdx} className="flex items-center justify-between p-2.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl text-xs">
                                <div className="flex items-center gap-2">
                                  <span className="font-bold text-[#252B3B] w-22">{slot.day}</span>
                                  <span className="font-mono text-[#4D5A8F] font-bold bg-white px-2 py-0.5 rounded border border-[#E8ECF5] shadow-3xs">
                                    {slot.startTime} to {slot.endTime}
                                  </span>
                                  {durStr && (
                                    <span className="text-[10px] text-[#747D91] font-mono">({durStr})</span>
                                  )}
                                </div>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setEditingClassTarget({
                                      ...editingClassTarget,
                                      schedule: editingClassTarget.schedule.filter((_, i) => i !== sIdx)
                                    });
                                  }}
                                  className="p-1 text-[#747D91] hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                                  title="Remove this slot"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>

                    {/* Add lecture slot(s) box with multi-day selection */}
                    <div className="p-3.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-2xl flex flex-col gap-3">
                      <span className="text-[10.5px] font-bold text-[#252B3B] flex items-center gap-1.5">
                        <Plus className="h-3.5 w-3.5 text-[#4D5A8F]" />
                        Add Lecture Slots (Select Multiple Days)
                      </span>

                      {/* Multi-Day Selector */}
                      <div className="flex flex-col gap-1.5">
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="font-bold text-[#747D91] uppercase">Select Days:</span>
                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              onClick={() => setEditingClassAddSlotDays(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])}
                              className="text-[#4D5A8F] font-bold hover:underline cursor-pointer"
                            >
                              Mon–Fri
                            </button>
                            <span className="text-[#8F9DCC]">•</span>
                            <button
                              type="button"
                              onClick={() => setEditingClassAddSlotDays([...DAYS_OF_WEEK])}
                              className="text-[#4D5A8F] font-bold hover:underline cursor-pointer"
                            >
                              All 7
                            </button>
                            <span className="text-[#8F9DCC]">•</span>
                            <button
                              type="button"
                              onClick={() => setEditingClassAddSlotDays([])}
                              className="text-[#747D91] hover:text-rose-600 font-medium cursor-pointer"
                            >
                              Clear
                            </button>
                          </div>
                        </div>

                        <div className="grid grid-cols-7 gap-1">
                          {DAYS_OF_WEEK.map(d => {
                            const isSelected = editingClassAddSlotDays.includes(d);
                            const shortName = d.slice(0, 3);
                            return (
                              <button
                                key={d}
                                type="button"
                                onClick={() => {
                                  if (isSelected) {
                                    setEditingClassAddSlotDays(prev => prev.filter(x => x !== d));
                                  } else {
                                    setEditingClassAddSlotDays(prev => [...prev, d]);
                                  }
                                }}
                                className={`py-1.5 px-0.5 rounded-lg text-xs font-bold transition-all border text-center cursor-pointer ${
                                  isSelected
                                    ? 'bg-[#4D5A8F] text-white border-[#4D5A8F] shadow-3xs font-black'
                                    : 'bg-white text-[#252B3B] border-[#E8ECF5] hover:bg-[#E3EAFA]'
                                }`}
                                title={d}
                              >
                                <span>{shortName}</span>
                              </button>
                            );
                          })}
                        </div>
                      </div>

                      {/* Time Period Inputs */}
                      <div className="flex flex-col gap-1.5 bg-white p-2.5 rounded-xl border border-[#E8ECF5] shadow-3xs">
                        <span className="text-[9.5px] font-bold text-[#747D91] uppercase tracking-wider">Time Period</span>
                        <div className="flex items-center gap-1.5 justify-between">
                          <div className="flex items-center bg-[#F5F8FF] border border-[#E8ECF5] rounded-lg p-1 text-xs">
                            <input
                              type="text"
                              value={editingClassAddSlotStartHour}
                              onChange={e => setEditingClassAddSlotStartHour(e.target.value)}
                              placeholder="09:00"
                              className="w-12 text-center font-mono font-bold outline-none text-[#252B3B]"
                            />
                            <button
                              type="button"
                              onClick={() => setEditingClassAddSlotStartPeriod(p => p === 'AM' ? 'PM' : 'AM')}
                              className="text-[9px] font-bold text-[#4D5A8F] bg-[#E3EAFA] px-1.5 py-0.5 rounded cursor-pointer hover:bg-[#CED8F2]"
                            >
                              {editingClassAddSlotStartPeriod}
                            </button>
                          </div>
                          <span className="text-[#747D91] text-xs font-bold">to</span>
                          <div className="flex items-center bg-[#F5F8FF] border border-[#E8ECF5] rounded-lg p-1 text-xs">
                            <input
                              type="text"
                              value={editingClassAddSlotEndHour}
                              onChange={e => setEditingClassAddSlotEndHour(e.target.value)}
                              placeholder="10:30"
                              className="w-12 text-center font-mono font-bold outline-none text-[#252B3B]"
                            />
                            <button
                              type="button"
                              onClick={() => setEditingClassAddSlotEndPeriod(p => p === 'AM' ? 'PM' : 'AM')}
                              className="text-[9px] font-bold text-[#4D5A8F] bg-[#E3EAFA] px-1.5 py-0.5 rounded cursor-pointer hover:bg-[#CED8F2]"
                            >
                              {editingClassAddSlotEndPeriod}
                            </button>
                          </div>
                        </div>

                        {/* Quick Presets */}
                        <div className="flex items-center gap-1 flex-wrap pt-1 border-t border-[#E8ECF5]">
                          {[
                            { start: '08:30', sP: 'AM', end: '10:00', eP: 'AM' },
                            { start: '08:30', sP: 'AM', end: '11:20', eP: 'AM' },
                            { start: '09:00', sP: 'AM', end: '10:30', eP: 'AM' },
                            { start: '01:00', sP: 'PM', end: '02:00', eP: 'PM' },
                            { start: '01:00', sP: 'PM', end: '02:30', eP: 'PM' }
                          ].map((preset, idx) => (
                            <button
                              key={idx}
                              type="button"
                              onClick={() => {
                                setEditingClassAddSlotStartHour(preset.start);
                                setNewClassStartPeriod(preset.sP as 'AM' | 'PM');
                                setEditingClassAddSlotEndHour(preset.end);
                                setEditingClassAddSlotEndPeriod(preset.eP as 'AM' | 'PM');
                              }}
                              className="text-[9px] font-mono px-1.5 py-0.5 rounded bg-[#F5F8FF] hover:bg-[#E3EAFA] hover:text-[#4D5A8F] text-[#252B3B] transition-colors cursor-pointer border border-[#E8ECF5]"
                            >
                              {preset.start}{preset.sP}-{preset.end}{preset.eP}
                            </button>
                          ))}
                        </div>
                      </div>

                      <button
                        type="button"
                        disabled={editingClassAddSlotDays.length === 0}
                        onClick={() => {
                          if (editingClassAddSlotDays.length === 0) return;
                          const newSlots: ClassScheduleSlot[] = editingClassAddSlotDays.map((d, i) => ({
                            id: `slot-${Date.now()}-${i}`,
                            day: d,
                            startTime: `${editingClassAddSlotStartHour.trim() || '09:00'} ${editingClassAddSlotStartPeriod}`,
                            endTime: `${editingClassAddSlotEndHour.trim() || '10:30'} ${editingClassAddSlotEndPeriod}`
                          }));
                          setEditingClassTarget({
                            ...editingClassTarget,
                            schedule: [...editingClassTarget.schedule, ...newSlots]
                          });
                          setEditingClassAddSlotDays([]);
                        }}
                        className="self-end px-3.5 py-1.5 bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-95 text-white text-xs font-bold rounded-xl transition-all cursor-pointer shadow-3xs flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
                      >
                        <Plus className="h-3.5 w-3.5" />
                        <span>
                          {editingClassAddSlotDays.length > 0 
                            ? `Append Slots (${editingClassAddSlotDays.length} ${editingClassAddSlotDays.length === 1 ? 'day' : 'days'})`
                            : 'Select Days to Append'}
                        </span>
                      </button>
                    </div>
                  </div>
                )}

                {/* Modal Footer Actions */}
                <div className="flex flex-col-reverse sm:flex-row items-center gap-2 sm:gap-2.5 justify-end mt-2 border-t border-[#E8ECF5] pt-4">
                  <button
                    type="button"
                    onClick={() => setEditingClassTarget(null)}
                    className="w-full sm:w-auto text-center justify-center bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] border border-[#E8ECF5] text-xs font-bold px-4.5 py-2.5 rounded-xl transition-all cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="w-full sm:w-auto text-center justify-center bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-[0.98] text-white text-xs font-bold px-5 py-2.5 rounded-xl transition-all cursor-pointer shadow-xs inline-flex items-center gap-1.5"
                  >
                    <Check className="h-4 w-4" />
                    Save Changes
                  </button>
                </div>
              </form>
            </div>
          </div>
        );
      })()}

      {/* POPUP MODAL: ADD STUDENT TO CLASS ROSTER */}
      {isAddStudentModalOpen && editingClassTarget && (
        <div className="fixed inset-0 z-[120] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-fade-in" id="add-student-modal-overlay">
          <div className="bg-white rounded-3xl p-6 md:p-7 max-w-md w-full shadow-2xl flex flex-col gap-4.5 border border-[#E8ECF5] transform scale-100 transition-all duration-300">
            {/* Header */}
            <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-[#E3EAFA] flex items-center justify-center text-[#4D5A8F]">
                  <UserPlus className="h-4 w-4" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-[#252B3B] tracking-tight">
                    Add a Student
                  </h3>
                  <p className="text-[11px] text-[#4D5A8F] font-mono">
                    Enrolling into: {editingClassTarget.newName || editingClassTarget.originalName}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsAddStudentModalOpen(false);
                  setAddStudentModalError(null);
                }}
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] border border-[#E8ECF5] transition-all cursor-pointer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleAddStudentToClass} className="flex flex-col gap-4">
              <p className="text-xs text-[#747D91] leading-relaxed">
                Fill in the student ID and name to enroll them into this class roster:
              </p>

              {/* Student ID Input */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-[#252B3B] uppercase tracking-wider flex items-center justify-between">
                  <span>Student ID Number *</span>
                  <span className="text-[9.5px] text-[#4D5A8F] font-normal">Must be unique</span>
                </label>
                <input
                  type="text"
                  value={addStudentIdInput}
                  onChange={e => {
                    setAddStudentIdInput(e.target.value);
                    setAddStudentModalError(null);
                  }}
                  placeholder="e.g. 106 or STU-2026-01"
                  className="w-full text-sm font-bold font-mono px-3.5 py-2.5 rounded-xl border border-[#E8ECF5] bg-[#F5F8FF] shadow-3xs focus:outline-none focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] transition-all"
                  autoFocus
                  required
                />
              </div>

              {/* Student Full Name Input */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-[#252B3B] uppercase tracking-wider">
                  Student Full Name *
                </label>
                <input
                  type="text"
                  value={addStudentNameInput}
                  onChange={e => {
                    setAddStudentNameInput(e.target.value);
                    setAddStudentModalError(null);
                  }}
                  placeholder="e.g. Lucas Vance"
                  className="w-full text-sm font-bold px-3.5 py-2.5 rounded-xl border border-[#E8ECF5] bg-[#F5F8FF] shadow-3xs focus:outline-none focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] transition-all"
                  required
                />
              </div>

              {addStudentModalError && (
                <div className="p-2.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-700 text-xs font-semibold flex items-center gap-2">
                  <AlertCircle className="h-4 w-4 shrink-0 text-rose-600" />
                  <span>{addStudentModalError}</span>
                </div>
              )}

              {/* Footer buttons */}
              <div className="flex flex-col-reverse sm:flex-row items-center gap-2 sm:gap-2.5 justify-end mt-2 pt-3 border-t border-[#E8ECF5]">
                <button
                  type="button"
                  onClick={() => {
                    setIsAddStudentModalOpen(false);
                    setAddStudentModalError(null);
                  }}
                  className="w-full sm:w-auto text-center justify-center bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] border border-[#E8ECF5] text-xs font-bold px-4 py-2.5 rounded-xl transition-all cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="w-full sm:w-auto text-center justify-center bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-[0.98] text-white text-xs font-bold px-4.5 py-2.5 rounded-xl transition-all cursor-pointer shadow-xs inline-flex items-center gap-1.5"
                >
                  <Plus className="h-4 w-4" />
                  <span>Add Student</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* QUICK ADD LECTURE SLOT MODAL */}
      {isAddScheduleModalOpen && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-fade-in" id="add-schedule-modal-overlay">
          <div className="bg-white rounded-3xl p-6 md:p-8 max-w-md w-full shadow-2xl flex flex-col gap-5 border border-[#E8ECF5] transform scale-100 transition-all duration-300">
            <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-3.5">
              <span className="text-base font-bold text-[#252B3B] tracking-tight flex items-center gap-2 font-sans">
                <Clock className="h-5 w-5 text-[#4D5A8F]" />
                <span>Schedule Lecture Slot</span>
              </span>
              <button
                type="button"
                onClick={() => setIsAddScheduleModalOpen(false)}
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] border border-[#E8ECF5] transition-all cursor-pointer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleAddScheduleSlotSubmit} className="flex flex-col gap-4">
              {/* Select Class */}
              <div className="flex flex-col gap-1">
                <label className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider">
                  Target Class Section
                </label>
                <select
                  value={scheduleModalClass}
                  onChange={e => setScheduleModalClass(e.target.value)}
                  className="text-xs bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl p-2.5 outline-none font-bold text-[#252B3B] shadow-3xs focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F]"
                  required
                >
                  {classes.map(c => (
                    <option key={c.name} value={c.name}>{c.name}</option>
                  ))}
                </select>
              </div>

              {/* Multi-Day Selection */}
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider">
                    Lecture Days (Select Multiple)
                  </label>
                  <div className="flex items-center gap-1.5 text-[10px]">
                    <button
                      type="button"
                      onClick={() => setScheduleModalDays(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])}
                      className="text-[#4D5A8F] font-bold hover:underline cursor-pointer"
                    >
                      Mon–Fri
                    </button>
                    <span className="text-[#8F9DCC]">•</span>
                    <button
                      type="button"
                      onClick={() => setScheduleModalDays([...DAYS_OF_WEEK])}
                      className="text-[#4D5A8F] font-bold hover:underline cursor-pointer"
                    >
                      All 7
                    </button>
                    <span className="text-[#8F9DCC]">•</span>
                    <button
                      type="button"
                      onClick={() => setScheduleModalDays([])}
                      className="text-[#747D91] hover:text-rose-600 font-medium cursor-pointer"
                    >
                      Clear
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-7 gap-1">
                  {DAYS_OF_WEEK.map(d => {
                    const isSelected = scheduleModalDays.includes(d);
                    const shortName = d.slice(0, 3);
                    return (
                      <button
                        key={d}
                        type="button"
                        onClick={() => {
                          if (isSelected) {
                            setScheduleModalDays(prev => prev.filter(x => x !== d));
                          } else {
                            setScheduleModalDays(prev => [...prev, d]);
                          }
                        }}
                        className={`py-2 px-1 rounded-xl text-xs font-bold transition-all border text-center cursor-pointer ${
                          isSelected
                            ? 'bg-[#4D5A8F] text-white border-[#4D5A8F] shadow-3xs font-black'
                            : 'bg-white text-[#252B3B] border-[#E8ECF5] hover:bg-[#E3EAFA]'
                        }`}
                        title={d}
                      >
                        <span className="block">{shortName}</span>
                      </button>
                    );
                  })}
                </div>

                {scheduleModalDays.length > 0 ? (
                  <span className="text-[10.5px] text-[#4D5A8F] font-semibold">
                    ✓ {scheduleModalDays.length} {scheduleModalDays.length === 1 ? 'day' : 'days'} selected: {scheduleModalDays.join(', ')}
                  </span>
                ) : (
                  <span className="text-[10px] text-amber-600 font-medium italic">
                    Please select at least one day.
                  </span>
                )}
              </div>

              {/* Time Period (__am/pm to __am/pm) */}
              <div className="flex flex-col gap-2 bg-[#F5F8FF] p-3.5 rounded-2xl border border-[#E8ECF5]">
                <label className="text-[10.5px] font-bold text-[#252B3B] flex items-center justify-between">
                  <span>Lecture Time Period</span>
                  <span className="text-[10.5px] text-[#4D5A8F] font-mono font-bold">
                    {scheduleModalStartHour} {scheduleModalStartPeriod} to {scheduleModalEndHour} {scheduleModalEndPeriod}
                  </span>
                </label>

                <div className="grid grid-cols-2 gap-2">
                  <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-[#E8ECF5] shadow-3xs">
                    <span className="text-[9.5px] font-bold text-[#747D91] uppercase">Start Time</span>
                    <div className="flex items-center gap-1">
                      <input
                        type="text"
                        value={scheduleModalStartHour}
                        onChange={e => setScheduleModalStartHour(e.target.value)}
                        placeholder="09:00"
                        className="w-full text-xs font-mono font-bold text-[#252B3B] outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => setScheduleModalStartPeriod(p => p === 'AM' ? 'PM' : 'AM')}
                        className="px-2 py-0.5 text-[10px] font-black rounded-md bg-[#E3EAFA] text-[#4D5A8F] hover:bg-[#CED8F2] cursor-pointer"
                      >
                        {scheduleModalStartPeriod}
                      </button>
                    </div>
                  </div>

                  <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-[#E8ECF5] shadow-3xs">
                    <span className="text-[9.5px] font-bold text-[#747D91] uppercase">End Time</span>
                    <div className="flex items-center gap-1">
                      <input
                        type="text"
                        value={scheduleModalEndHour}
                        onChange={e => setScheduleModalEndHour(e.target.value)}
                        placeholder="10:30"
                        className="w-full text-xs font-mono font-bold text-[#252B3B] outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => setScheduleModalEndPeriod(p => p === 'AM' ? 'PM' : 'AM')}
                        className="px-2 py-0.5 text-[10px] font-black rounded-md bg-[#E3EAFA] text-[#4D5A8F] hover:bg-[#CED8F2] cursor-pointer"
                      >
                        {scheduleModalEndPeriod}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Quick Presets */}
                <div className="flex items-center gap-1 flex-wrap pt-1 border-t border-[#E8ECF5]">
                  <span className="text-[9.5px] text-[#747D91] font-medium">Presets:</span>
                  {[
                    { start: '08:30', sP: 'AM', end: '10:00', eP: 'AM' },
                    { start: '08:30', sP: 'AM', end: '11:20', eP: 'AM' },
                    { start: '09:00', sP: 'AM', end: '10:30', eP: 'AM' },
                    { start: '01:00', sP: 'PM', end: '02:00', eP: 'PM' },
                    { start: '01:00', sP: 'PM', end: '02:30', eP: 'PM' }
                  ].map((preset, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => {
                        setScheduleModalStartHour(preset.start);
                        setScheduleModalStartPeriod(preset.sP as 'AM' | 'PM');
                        setScheduleModalEndHour(preset.end);
                        setScheduleModalEndPeriod(preset.eP as 'AM' | 'PM');
                      }}
                      className="text-[9px] font-mono px-1.5 py-0.5 rounded bg-white hover:bg-[#E3EAFA] hover:text-[#4D5A8F] text-[#252B3B] transition-colors cursor-pointer border border-[#E8ECF5]"
                    >
                      {preset.start}{preset.sP}-{preset.end}{preset.eP}
                    </button>
                  ))}
                </div>
              </div>

              {/* Actions */}
              <div className="flex flex-col-reverse sm:flex-row items-center gap-2 sm:gap-2.5 justify-end mt-2 border-t border-[#E8ECF5] pt-4">
                <button
                  type="button"
                  onClick={() => setIsAddScheduleModalOpen(false)}
                  className="w-full sm:w-auto text-center justify-center bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] border border-[#E8ECF5] text-xs font-bold px-4.5 py-2.5 rounded-xl transition-all cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={scheduleModalDays.length === 0}
                  className="w-full sm:w-auto text-center justify-center bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-[0.98] text-white text-xs font-bold px-5 py-2.5 rounded-xl transition-all cursor-pointer shadow-xs inline-flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Plus className="h-4 w-4" />
                  <span>
                    {scheduleModalDays.length > 0 
                      ? `Schedule on ${scheduleModalDays.length} ${scheduleModalDays.length === 1 ? 'Day' : 'Days'}`
                      : 'Choose Days'}
                  </span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODIFY SINGLE TIME SLOT MODAL (TRIGGERED VIA PEN ICON IN TIMETABLE BAR) */}
      {editingSlotTarget && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-in fade-in duration-200" id="modify-slot-modal-overlay">
          <div className="bg-white rounded-3xl p-6 md:p-7 max-w-md w-full shadow-2xl flex flex-col gap-5 border border-[#E8ECF5] animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between pb-3 border-b border-[#E8ECF5]">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-[#E3EAFA] text-[#4D5A8F] rounded-xl border border-[#B3C0E6]/50">
                  <Pencil className="h-4.5 w-4.5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-[#252B3B] text-base">
                    Modify Lecture Time Slot
                  </h3>
                  <p className="text-xs text-[#747D91] mt-0.5">
                    {editingSlotTarget.className} • {editingSlotTarget.day}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setEditingSlotTarget(null)}
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleSaveEditedSlot} className="flex flex-col gap-4">
              {/* Day of the Week */}
              <div className="flex flex-col gap-1">
                <label className="text-[10px] font-bold text-[#747D91] uppercase tracking-wider">
                  Day of the Week
                </label>
                <select
                  value={editingSlotTarget.day}
                  onChange={e => setEditingSlotTarget({ ...editingSlotTarget, day: e.target.value as DayOfWeek })}
                  className="text-xs bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl p-2.5 outline-none font-bold text-[#252B3B] shadow-3xs cursor-pointer focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F]"
                >
                  {DAYS_OF_WEEK.map(d => (
                    <option key={d} value={d}>{d}</option>
                  ))}
                </select>
              </div>

              {/* Time Period */}
              <div className="flex flex-col gap-2 bg-[#F5F8FF] p-3.5 rounded-2xl border border-[#E8ECF5]">
                <label className="text-[10.5px] font-bold text-[#252B3B] flex items-center justify-between">
                  <span>Lecture Time Period</span>
                  <span className="text-[10.5px] text-[#4D5A8F] font-mono font-bold">
                    {editingSlotTarget.startHour} {editingSlotTarget.startPeriod} to {editingSlotTarget.endHour} {editingSlotTarget.endPeriod}
                  </span>
                </label>

                <div className="grid grid-cols-2 gap-2">
                  <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-[#E8ECF5] shadow-3xs">
                    <span className="text-[9.5px] font-bold text-[#747D91] uppercase">Start Time</span>
                    <div className="flex items-center gap-1">
                      <input
                        type="text"
                        value={editingSlotTarget.startHour}
                        onChange={e => setEditingSlotTarget({ ...editingSlotTarget, startHour: e.target.value })}
                        placeholder="09:00"
                        className="w-full text-xs font-mono font-bold text-[#252B3B] outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => setEditingSlotTarget({ ...editingSlotTarget, startPeriod: editingSlotTarget.startPeriod === 'AM' ? 'PM' : 'AM' })}
                        className="px-2 py-0.5 text-[10px] font-black rounded-md bg-[#E3EAFA] text-[#4D5A8F] hover:bg-[#CED8F2] cursor-pointer"
                      >
                        {editingSlotTarget.startPeriod}
                      </button>
                    </div>
                  </div>

                  <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-[#E8ECF5] shadow-3xs">
                    <span className="text-[9.5px] font-bold text-[#747D91] uppercase">End Time</span>
                    <div className="flex items-center gap-1">
                      <input
                        type="text"
                        value={editingSlotTarget.endHour}
                        onChange={e => setEditingSlotTarget({ ...editingSlotTarget, endHour: e.target.value })}
                        placeholder="10:30"
                        className="w-full text-xs font-mono font-bold text-[#252B3B] outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => setEditingSlotTarget({ ...editingSlotTarget, endPeriod: editingSlotTarget.endPeriod === 'AM' ? 'PM' : 'AM' })}
                        className="px-2 py-0.5 text-[10px] font-black rounded-md bg-[#E3EAFA] text-[#4D5A8F] hover:bg-[#CED8F2] cursor-pointer"
                      >
                        {editingSlotTarget.endPeriod}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Quick presets */}
                <div className="flex items-center gap-1.5 flex-wrap pt-1">
                  <span className="text-[9.5px] text-[#747D91] font-medium">Quick presets:</span>
                  {[
                    { start: '08:30', sP: 'AM', end: '10:00', eP: 'AM' },
                    { start: '08:30', sP: 'AM', end: '11:20', eP: 'AM' },
                    { start: '09:00', sP: 'AM', end: '10:30', eP: 'AM' },
                    { start: '10:30', sP: 'AM', end: '12:00', eP: 'PM' },
                    { start: '01:00', sP: 'PM', end: '02:00', eP: 'PM' },
                    { start: '01:00', sP: 'PM', end: '02:30', eP: 'PM' },
                    { start: '02:30', sP: 'PM', end: '04:00', eP: 'PM' }
                  ].map((preset, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => {
                        setEditingSlotTarget({
                          ...editingSlotTarget,
                          startHour: preset.start,
                          startPeriod: preset.sP as 'AM' | 'PM',
                          endHour: preset.end,
                          endPeriod: preset.eP as 'AM' | 'PM'
                        });
                      }}
                      className="text-[9.5px] font-mono px-1.5 py-0.5 rounded bg-white hover:bg-[#E3EAFA] hover:text-[#4D5A8F] text-[#252B3B] transition-colors cursor-pointer border border-[#E8ECF5]"
                    >
                      {preset.start}{preset.sP}-{preset.end}{preset.eP}
                    </button>
                  ))}
                </div>
              </div>

              {editingSlotError && (
                <div className="text-[11px] bg-rose-50 text-rose-700 p-2.5 rounded-xl border border-rose-200 flex items-center gap-2 font-semibold">
                  <AlertCircle className="h-4 w-4 shrink-0 text-rose-600" />
                  <span>{editingSlotError}</span>
                </div>
              )}

              {/* Actions */}
              <div className="flex flex-col-reverse sm:flex-row items-center gap-2 sm:gap-2.5 justify-end mt-2 border-t border-[#E8ECF5] pt-4">
                <button
                  type="button"
                  onClick={() => setEditingSlotTarget(null)}
                  className="w-full sm:w-auto text-center justify-center bg-[#F5F8FF] hover:bg-[#E3EAFA] text-[#252B3B] border border-[#E8ECF5] text-xs font-bold px-4.5 py-2.5 rounded-xl transition-all cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="w-full sm:w-auto text-center justify-center bg-[#4D5A8F] hover:bg-[#3D4873] active:scale-[0.98] text-white text-xs font-bold px-5 py-2.5 rounded-xl transition-all cursor-pointer shadow-xs inline-flex items-center gap-1.5"
                >
                  <Check className="h-4 w-4" />
                  Save Slot Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Assessment Modal (Add New Assessment & Edit Assessment Parameters) */}
      {isAssessmentModalOpen && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-in fade-in duration-200" id="assessment-configuration-modal">
          <div className="bg-white rounded-3xl p-6 md:p-7 max-w-md w-full shadow-2xl flex flex-col gap-5 border border-[#E8ECF5] animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between pb-3 border-b border-[#E8ECF5]">
              <div>
                <h3 className="font-extrabold text-[#252B3B] text-base">
                  {assessmentModalMode === 'add' ? 'Add New Assessment' : 'Edit Assessment'}
                </h3>
                <p className="text-xs text-[#747D91] mt-0.5">
                  {assessmentModalMode === 'add' 
                    ? `Create an assessment for ${activeClass.name}`
                    : `Update parameters for ${editingAssessmentOriginalName}`}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setIsAssessmentModalOpen(false)}
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleSaveAssessmentModal} className="flex flex-col gap-4">
              {/* Assessment Name */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[11px] font-bold text-[#747D91] uppercase tracking-wider">Assessment Name</label>
                <input
                  type="text"
                  value={assessNameInput}
                  onChange={e => setAssessNameInput(e.target.value)}
                  placeholder="e.g. Test 1, Quiz 2, Final Project"
                  required
                  autoFocus
                  className="text-xs bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl p-2.5 outline-none focus:bg-white focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] placeholder-[#747D91]/60 font-semibold transition-all"
                />
              </div>

              {/* Full Mark (Max Score) */}
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-[11px] font-bold text-[#747D91] uppercase tracking-wider">Full Mark (Max Score)</label>
                  <span className="text-[10px] text-[#747D91]">Total points possible</span>
                </div>
                <input
                  type="number"
                  min={1}
                  max={1000}
                  value={assessMaxMarkInput}
                  onChange={e => setAssessMaxMarkInput(parseInt(e.target.value) || 0)}
                  required
                  placeholder="e.g. 20, 50, 100"
                  className="text-xs bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl p-2.5 outline-none focus:bg-white focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] font-mono font-bold transition-all"
                />
                <span className="text-[10.5px] text-[#747D91]">
                  Example: A test can be out of 20 or 100, while its weight determines final contribution.
                </span>
              </div>

              {/* Weight Out of 100% */}
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-[11px] font-bold text-[#747D91] uppercase tracking-wider">Weight (% of Final Grade)</label>
                  <span className="text-[10px] font-mono text-[#4D5A8F] font-bold">Max total: 100%</span>
                </div>
                <div className="relative flex items-center">
                  <input
                    type="number"
                    min={0}
                    max={100}
                    value={assessWeightInput}
                    onChange={e => setAssessWeightInput(parseInt(e.target.value) || 0)}
                    required
                    className="w-full text-xs bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl p-2.5 pr-8 outline-none focus:bg-white focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F] text-[#252B3B] font-mono font-bold transition-all"
                  />
                  <span className="absolute right-3 font-bold text-[#747D91] text-xs">%</span>
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
                          : 'bg-[#E3EAFA]/70 border-[#B3C0E6] text-[#252B3B]'
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
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5 pt-3 border-t border-[#E8ECF5] mt-2">
                {assessmentModalMode === 'edit' && editingAssessmentOriginalName ? (
                  <button
                    type="button"
                    onClick={() => handleDeleteAssessment(editingAssessmentOriginalName)}
                    className="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-bold text-rose-600 hover:text-rose-700 hover:bg-rose-50 rounded-xl transition-colors cursor-pointer w-full sm:w-auto"
                  >
                    <Trash2 className="h-4 w-4" />
                    <span>Delete Test</span>
                  </button>
                ) : <div className="hidden sm:block" />}

                <div className="flex flex-col-reverse sm:flex-row items-center gap-2 w-full sm:w-auto">
                  <button
                    type="button"
                    onClick={() => setIsAssessmentModalOpen(false)}
                    className="w-full sm:w-auto text-center justify-center px-4 py-2.5 text-xs font-bold text-[#252B3B] bg-[#F5F8FF] hover:bg-[#E3EAFA] border border-[#E8ECF5] rounded-xl transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="w-full sm:w-auto text-center justify-center px-5 py-2.5 text-xs font-bold text-white bg-[#4D5A8F] hover:bg-[#3D4873] rounded-xl shadow-xs transition-all cursor-pointer inline-flex items-center gap-1.5 active:scale-95"
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
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs animate-in fade-in duration-200" id="final-grade-weights-modal">
          <div className="bg-white rounded-3xl p-6 md:p-7 max-w-md w-full shadow-2xl flex flex-col gap-5 border border-[#E8ECF5] animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between pb-3 border-b border-[#E8ECF5]">
              <div>
                <h3 className="font-extrabold text-[#252B3B] text-base flex items-center gap-2">
                  <span>📊 Final Grade Weight Distribution</span>
                </h3>
                <p className="text-xs text-[#747D91] mt-0.5">
                  Set how much each test weighs. All tests combined must weigh 100%.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setIsFinalGradeModalOpen(false)}
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {activeClass.assessments.length === 0 ? (
              <div className="py-8 text-center text-[#747D91] italic text-xs">
                No assessments registered yet for this class section.
              </div>
            ) : (
              <div className="flex flex-col gap-3">
                <div className="flex items-center justify-between text-[11px] text-[#747D91] font-bold uppercase tracking-wider px-1">
                  <span>Assessment Name & Max Mark</span>
                  <span>Weight (%)</span>
                </div>

                <div className="max-h-60 overflow-y-auto flex flex-col gap-2 p-1">
                  {activeClass.assessments.map(a => {
                    const currentVal = tempWeights[a.name] !== undefined ? tempWeights[a.name] : a.weight;
                    return (
                      <div key={a.name} className="flex items-center justify-between p-3 bg-[#F5F8FF] rounded-2xl border border-[#E8ECF5] shadow-3xs">
                        <div className="min-w-0 pr-2">
                          <span className="font-bold text-[#252B3B] text-xs block truncate">{a.name}</span>
                          <span className="text-[10px] text-[#747D91] font-mono">Full Mark: {a.maxMark || 100} pts</span>
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
                            className="w-20 text-xs bg-white border border-[#E8ECF5] rounded-xl p-2 pr-6 text-right font-mono font-bold text-[#252B3B] outline-none focus:ring-2 focus:ring-[#4D5A8F]/20 focus:border-[#4D5A8F]"
                          />
                          <span className="absolute right-2.5 font-bold text-[#747D91] text-xs">%</span>
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
                      className="text-[11px] font-bold text-[#4D5A8F] hover:text-[#3D4873] hover:underline cursor-pointer"
                    >
                      Auto-Distribute Evenly
                    </button>
                  </div>
                )}

                {/* Weight Total Status Banner */}
                {(() => {
                  const currentSum: number = (Object.values(tempWeights) as number[]).reduce((sum: number, w: number) => sum + (Number(w) || 0), 0);
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
                <div className="flex flex-col-reverse sm:flex-row items-center justify-end gap-2 sm:gap-2.5 pt-3 border-t border-[#E8ECF5]">
                  <button
                    type="button"
                    onClick={() => setIsFinalGradeModalOpen(false)}
                    className="w-full sm:w-auto text-center justify-center px-4 py-2.5 text-xs font-bold text-[#252B3B] bg-[#F5F8FF] hover:bg-[#E3EAFA] border border-[#E8ECF5] rounded-xl transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={handleSaveFinalGradeWeights}
                    disabled={((Object.values(tempWeights) as number[]).reduce((sum: number, w: number) => sum + (Number(w) || 0), 0)) > 100}
                    className="w-full sm:w-auto text-center justify-center px-5 py-2.5 text-xs font-bold text-white bg-[#4D5A8F] hover:bg-[#3D4873] disabled:opacity-50 disabled:cursor-not-allowed rounded-xl shadow-xs transition-all cursor-pointer inline-flex items-center gap-1.5 active:scale-95"
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
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-3xs animate-in fade-in duration-150" id="delete-attendance-date-modal">
          <div className="bg-white rounded-2xl max-w-sm w-full p-5 shadow-2xl border border-rose-200 flex flex-col gap-4 animate-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-rose-100 text-rose-600 shrink-0">
                <Trash2 className="h-5 w-5" />
              </div>
              <div className="flex-1 min-w-0">
                <h4 className="text-sm font-bold text-[#252B3B]">Delete Attendance Date Column</h4>
                <p className="text-xs text-[#747D91] mt-1 leading-normal">
                  Are you sure you want to delete the attendance date column for <strong className="text-[#252B3B] font-bold">{confirmDeleteDateTarget.displayDate}</strong> in <strong className="text-[#252B3B] font-bold">{activeClass.name}</strong>?
                </p>
                <p className="text-[11px] text-rose-700 bg-rose-50 border border-rose-200 rounded-lg p-2 mt-2 font-medium">
                  ⚠️ Warning: All attendance check logs for this date will be permanently deleted.
                </p>
              </div>
            </div>

            <div className="flex flex-col-reverse sm:flex-row items-center justify-end gap-2 pt-2 border-t border-[#E8ECF5]">
              <button
                type="button"
                onClick={() => setConfirmDeleteDateTarget(null)}
                className="w-full sm:w-auto text-center justify-center px-3.5 py-2 text-xs font-semibold text-[#252B3B] bg-[#F5F8FF] hover:bg-[#E3EAFA] border border-[#E8ECF5] rounded-xl transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => handleDeleteAttendanceDate(confirmDeleteDateTarget.date)}
                className="w-full sm:w-auto text-center justify-center px-3.5 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-xs transition-all cursor-pointer flex items-center gap-1.5 active:scale-95"
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
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-3xs animate-in fade-in duration-150" id="delete-student-warning-modal">
          <div className="bg-white rounded-2xl max-w-sm w-full p-5 shadow-2xl border border-rose-200 flex flex-col gap-4 animate-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-rose-100 text-rose-600 shrink-0">
                <AlertCircle className="h-5 w-5" />
              </div>
              <div className="flex-1 min-w-0">
                <h4 className="text-sm font-bold text-[#252B3B]">Delete Student Confirmation</h4>
                <p className="text-xs text-[#747D91] mt-1 leading-normal">
                  Are you sure you want to remove <strong className="text-[#252B3B] font-bold">{confirmDeleteStudentTarget.studentName}</strong> (ID: {confirmDeleteStudentTarget.studentId}) from <strong className="text-[#252B3B] font-bold">{confirmDeleteStudentTarget.className}</strong>?
                </p>
                <p className="text-[11px] text-rose-700 bg-rose-50 border border-rose-200 rounded-lg p-2 mt-2 font-medium">
                  ⚠️ Warning: All attendance records and graded scores for this student in this class will be permanently removed.
                </p>
              </div>
            </div>

            <div className="flex flex-col-reverse sm:flex-row items-center justify-end gap-2 pt-2 border-t border-[#E8ECF5]">
              <button
                type="button"
                onClick={() => setConfirmDeleteStudentTarget(null)}
                className="w-full sm:w-auto text-center justify-center px-3.5 py-2 text-xs font-semibold text-[#252B3B] bg-[#F5F8FF] hover:bg-[#E3EAFA] border border-[#E8ECF5] rounded-xl transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => handleDeleteStudentFromClass(confirmDeleteStudentTarget.studentId, confirmDeleteStudentTarget.className)}
                className="w-full sm:w-auto text-center justify-center px-3.5 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-xs transition-all cursor-pointer flex items-center gap-1.5 active:scale-95"
              >
                <Trash2 className="h-3.5 w-3.5" />
                <span>Confirm Delete</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Calendar & Agenda: Event Creator / Editor Modal */}
      {isEventModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-3xs animate-in fade-in duration-150" id="calendar-event-editor-modal">
          <div className="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-[#E8ECF5] flex flex-col gap-5 animate-in zoom-in-95 duration-150 max-h-[90vh] overflow-y-auto">
            {/* Modal Header */}
            <div className="flex items-center justify-between border-b border-[#E8ECF5] pb-3">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-[#E3EAFA] text-[#4D5A8F] rounded-xl border border-[#CED8F2]">
                  <CalendarPlus className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-base font-extrabold text-[#252B3B]">
                    {editingEventId ? 'Edit Calendar Event or Note' : 'Add Calendar Event or Note'}
                  </h3>
                  <p className="text-xs text-[#747D91]">
                    Record an exam, meeting, deadline, or general class note.
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsEventModalOpen(false)}
                className="p-1.5 hover:bg-[#E3EAFA] rounded-full text-[#747D91] hover:text-[#252B3B] transition-colors cursor-pointer"
                title="Close"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            {/* Error Message */}
            {eventModalError && (
              <div className="bg-rose-50 border border-rose-200 text-rose-700 text-xs px-3.5 py-2.5 rounded-xl flex items-center gap-2">
                <AlertCircle className="h-4 w-4 shrink-0 text-rose-600" />
                <span>{eventModalError}</span>
              </div>
            )}

            <form onSubmit={handleSaveEvent} className="flex flex-col gap-4">
              {/* Event Title */}
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-[#252B3B]">
                  Title / Subject <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={eventTitleInput}
                  onChange={(e) => setEventTitleInput(e.target.value)}
                  placeholder="e.g. Midterm Exam, Staff Meeting, Project Deadline..."
                  className="w-full px-3.5 py-2.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl text-xs font-semibold text-[#252B3B] placeholder-[#747D91]/60 focus:bg-white focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20 outline-none transition-all"
                  autoFocus
                />
              </div>

              {/* Event Type Pills */}
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-[#252B3B]">
                  Category Type
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                  {(['exam', 'meeting', 'lecture', 'deadline', 'note'] as CalendarEventType[]).map(typeKey => {
                    const badge = getEventTypeBadgeInfo(typeKey);
                    const isSelected = eventTypeInput === typeKey;
                    return (
                      <button
                        key={typeKey}
                        type="button"
                        onClick={() => setEventTypeInput(typeKey)}
                        className={`flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold border transition-all cursor-pointer ${
                          isSelected
                            ? 'bg-[#4D5A8F] text-white border-[#4D5A8F] shadow-xs'
                            : `${badge.chipBg} border-[#E8ECF5]`
                        }`}
                      >
                        <span className={`w-2 h-2 rounded-full ${isSelected ? 'bg-white' : badge.dotBg}`} />
                        <span>{badge.label}</span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Date & Time Row */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                {/* Date */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-xs font-bold text-[#252B3B]">
                    Date <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="date"
                    value={eventDateInput}
                    onChange={(e) => setEventDateInput(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl text-xs font-semibold text-[#252B3B] focus:bg-white focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20 outline-none transition-all cursor-pointer"
                  />
                </div>

                {/* Time & All-Day */}
                <div className="flex flex-col gap-1.5">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-bold text-[#252B3B]">
                      Meeting Hour / Time
                    </label>
                    <label className="inline-flex items-center gap-1.5 text-[11px] font-semibold text-[#6D7CB0] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={isEventAllDay}
                        onChange={(e) => setIsEventAllDay(e.target.checked)}
                        className="rounded border-[#B3C0E6] text-[#4D5A8F] focus:ring-[#4D5A8F]"
                      />
                      <span>All Day</span>
                    </label>
                  </div>
                  <input
                    type="text"
                    disabled={isEventAllDay}
                    value={isEventAllDay ? 'All Day' : eventTimeInput}
                    onChange={(e) => setEventTimeInput(e.target.value)}
                    placeholder="e.g. 10:00 AM - 11:30 AM"
                    className={`w-full px-3.5 py-2.5 border rounded-xl text-xs font-semibold transition-all ${
                      isEventAllDay
                        ? 'bg-[#E3EAFA]/50 text-[#747D91] border-[#E8ECF5] cursor-not-allowed'
                        : 'bg-[#F5F8FF] text-[#252B3B] border-[#E8ECF5] focus:bg-white focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20 outline-none'
                    }`}
                  />
                </div>
              </div>

              {/* Class Section Target */}
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-[#252B3B]">
                  Associated Class Section
                </label>
                <select
                  value={eventClassInput}
                  onChange={(e) => setEventClassInput(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl text-xs font-semibold text-[#252B3B] focus:bg-white focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20 outline-none transition-all cursor-pointer"
                >
                  <option value="All Classes">All Classes (General Academic Note)</option>
                  {classes.map(c => (
                    <option key={c.name} value={c.name}>{c.name}</option>
                  ))}
                </select>
              </div>

              {/* Description / Detailed Notes */}
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-[#252B3B]">
                  Details &amp; Agenda Notes
                </label>
                <textarea
                  rows={3}
                  value={eventDescriptionInput}
                  onChange={(e) => setEventDescriptionInput(e.target.value)}
                  placeholder="Add exam room, chapters to study, meeting agenda items, or specific instructions for this day..."
                  className="w-full px-3.5 py-2.5 bg-[#F5F8FF] border border-[#E8ECF5] rounded-xl text-xs font-medium text-[#252B3B] placeholder-[#747D91]/60 focus:bg-white focus:border-[#4D5A8F] focus:ring-2 focus:ring-[#4D5A8F]/20 outline-none transition-all resize-y"
                />
              </div>

              {/* Modal Actions */}
              <div className="flex flex-col-reverse sm:flex-row items-stretch sm:items-center justify-end gap-2 sm:gap-2.5 pt-3 border-t border-[#E8ECF5]">
                <button
                  type="button"
                  onClick={() => setIsEventModalOpen(false)}
                  className="w-full sm:w-auto text-center justify-center px-4 py-2.5 text-xs font-semibold text-[#252B3B] bg-[#F5F8FF] hover:bg-[#E3EAFA] border border-[#E8ECF5] rounded-xl transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="w-full sm:w-auto text-center justify-center px-5 py-2.5 text-xs font-bold text-white bg-[#4D5A8F] hover:bg-[#3D4873] rounded-xl shadow-xs transition-all cursor-pointer flex items-center gap-1.5 active:scale-95"
                >
                  <Check className="h-4 w-4" />
                  <span>{editingEventId ? 'Save Changes' : 'Create Entry'}</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Delete Calendar Event Confirmation Modal */}
      {confirmDeleteEventTarget && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-3xs animate-in fade-in duration-150" id="delete-calendar-event-modal">
          <div className="bg-white rounded-2xl max-w-sm w-full p-5 shadow-2xl border border-rose-200 flex flex-col gap-4 animate-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-rose-100 text-rose-600 shrink-0">
                <Trash2 className="h-5 w-5" />
              </div>
              <div className="flex-1 min-w-0">
                <h4 className="text-sm font-bold text-[#252B3B]">Delete Calendar Entry</h4>
                <p className="text-xs text-[#747D91] mt-1 leading-normal">
                  Are you sure you want to delete <strong className="text-[#252B3B] font-bold">{confirmDeleteEventTarget.title}</strong> scheduled for <strong className="text-[#252B3B] font-bold">{formatFullDateString(confirmDeleteEventTarget.date)}</strong>?
                </p>
              </div>
            </div>

            <div className="flex flex-col-reverse sm:flex-row items-center justify-end gap-2 pt-2 border-t border-[#E8ECF5]">
              <button
                type="button"
                onClick={() => setConfirmDeleteEventTarget(null)}
                className="w-full sm:w-auto text-center justify-center px-3.5 py-2 text-xs font-semibold text-[#252B3B] bg-[#F5F8FF] hover:bg-[#E3EAFA] border border-[#E8ECF5] rounded-xl transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => handleDeleteEvent(confirmDeleteEventTarget.id)}
                className="w-full sm:w-auto text-center justify-center px-3.5 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-xs transition-all cursor-pointer flex items-center gap-1.5 active:scale-95"
              >
                <Trash2 className="h-3.5 w-3.5" />
                <span>Confirm Delete</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Mobile Fixed Bottom Navigation Bar - Ergonomic thumb reach for phone users */}
      <nav
        id="omni-mobile-bottom-nav"
        aria-label="Mobile Navigation"
        className="fixed bottom-0 left-0 right-0 z-40 bg-white/95 backdrop-blur-md border-t border-[#E8ECF5] px-1 py-1 flex lg:hidden items-center justify-around shadow-[0_-4px_16px_rgba(0,0,0,0.06)]"
      >
        <button
          type="button"
          onClick={() => handleSelectTab('timetable')}
          className={`flex flex-col items-center justify-center flex-1 py-1 px-1 rounded-xl transition-all min-h-[46px] cursor-pointer ${
            activeTab === 'timetable'
              ? 'text-[#4D5A8F] font-bold bg-[#E3EAFA]/70'
              : 'text-[#747D91] hover:text-[#4D5A8F] font-medium'
          }`}
        >
          <Clock className={`h-4.5 w-4.5 mb-0.5 ${activeTab === 'timetable' ? 'text-[#4D5A8F]' : 'text-[#747D91]'}`} />
          <span className="text-[10px] leading-tight">Timetable</span>
        </button>

        <button
          type="button"
          onClick={() => handleSelectTab('calendar')}
          className={`flex flex-col items-center justify-center flex-1 py-1 px-1 rounded-xl transition-all min-h-[46px] cursor-pointer ${
            activeTab === 'calendar'
              ? 'text-[#4D5A8F] font-bold bg-[#E3EAFA]/70'
              : 'text-[#747D91] hover:text-[#4D5A8F] font-medium'
          }`}
        >
          <CalendarDays className={`h-4.5 w-4.5 mb-0.5 ${activeTab === 'calendar' ? 'text-[#4D5A8F]' : 'text-[#747D91]'}`} />
          <span className="text-[10px] leading-tight">Calendar</span>
        </button>

        <button
          type="button"
          onClick={() => handleSelectTab('students')}
          className={`flex flex-col items-center justify-center flex-1 py-1 px-1 rounded-xl transition-all min-h-[46px] cursor-pointer ${
            activeTab === 'students'
              ? 'text-[#4D5A8F] font-bold bg-[#E3EAFA]/70'
              : 'text-[#747D91] hover:text-[#4D5A8F] font-medium'
          }`}
        >
          <Users className={`h-4.5 w-4.5 mb-0.5 ${activeTab === 'students' ? 'text-[#4D5A8F]' : 'text-[#747D91]'}`} />
          <span className="text-[10px] leading-tight">Students</span>
        </button>

        <button
          type="button"
          onClick={() => handleSelectTab('attendance')}
          className={`flex flex-col items-center justify-center flex-1 py-1 px-1 rounded-xl transition-all min-h-[46px] cursor-pointer ${
            activeTab === 'attendance'
              ? 'text-[#4D5A8F] font-bold bg-[#E3EAFA]/70'
              : 'text-[#747D91] hover:text-[#4D5A8F] font-medium'
          }`}
        >
          <CalendarCheck className={`h-4.5 w-4.5 mb-0.5 ${activeTab === 'attendance' ? 'text-[#4D5A8F]' : 'text-[#747D91]'}`} />
          <span className="text-[10px] leading-tight">Attendance</span>
        </button>

        <button
          type="button"
          onClick={() => handleSelectTab('grades')}
          className={`flex flex-col items-center justify-center flex-1 py-1 px-1 rounded-xl transition-all min-h-[46px] cursor-pointer ${
            activeTab === 'grades'
              ? 'text-[#4D5A8F] font-bold bg-[#E3EAFA]/70'
              : 'text-[#747D91] hover:text-[#4D5A8F] font-medium'
          }`}
        >
          <GraduationCap className={`h-4.5 w-4.5 mb-0.5 ${activeTab === 'grades' ? 'text-[#4D5A8F]' : 'text-[#747D91]'}`} />
          <span className="text-[10px] leading-tight">Grades</span>
        </button>
      </nav>

      {/* footer credits */}
      <footer className="bg-white px-6 py-3 pb-20 lg:pb-3 border-t border-[#E8ECF5] text-center text-[10px] text-[#747D91] flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2" id="omni-class-footer">
        <span>Omni Class — Academic Roster, Attendance & Grade Records Manager</span>
        <span className="font-mono text-[#8F9DCC]">Offline-Ready & Cloud-Connected</span>
      </footer>
    </div>
  );
}
