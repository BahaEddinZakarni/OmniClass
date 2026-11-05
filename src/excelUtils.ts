import * as XLSX from 'xlsx';
import { Student, Assessment, ClassSection } from './types';

export interface ParsedClassSectionData {
  className: string;
  students: Student[];
  attendanceDates: string[];
  attendance: Record<string, Record<string, 'Not Taken' | 'Attended' | 'Absent'>>;
  assessments: Assessment[];
  grades: Record<string, Record<string, number>>;
}

export interface ParseExcelResult {
  isMultiSection: boolean;
  sections: ParsedClassSectionData[];
  summary: string;
  totalStudents: number;
}

/**
 * Standardizes any date string into canonical YYYY-MM-DD format.
 * Accurately parses DD-MM-YY, DD/MM/YY, DD-MM-YYYY, YYYY-MM-DD, and Date objects.
 */
export function parseDateToStandardYMD(d: string): string | null {
  if (!d) return null;
  const trimmed = String(d).trim();
  if (!trimmed) return null;

  // 1. Matches YYYY-MM-DD or YYYY/MM/DD
  const ymdMatch = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (ymdMatch) {
    const y = ymdMatch[1];
    const m = String(parseInt(ymdMatch[2], 10)).padStart(2, '0');
    const day = String(parseInt(ymdMatch[3], 10)).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  // 2. Matches DD-MM-YYYY or DD/MM/YYYY
  const dmyLongMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyLongMatch) {
    const day = String(parseInt(dmyLongMatch[1], 10)).padStart(2, '0');
    const m = String(parseInt(dmyLongMatch[2], 10)).padStart(2, '0');
    const y = dmyLongMatch[3];
    return `${y}-${m}-${day}`;
  }

  // 3. Matches DD-MM-YY or DD/MM/YY (Our exported format)
  const dmyShortMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/);
  if (dmyShortMatch) {
    const day = String(parseInt(dmyShortMatch[1], 10)).padStart(2, '0');
    const m = String(parseInt(dmyShortMatch[2], 10)).padStart(2, '0');
    const yrNum = parseInt(dmyShortMatch[3], 10);
    const fullYear = yrNum < 70 ? 2000 + yrNum : 1900 + yrNum;
    return `${fullYear}-${m}-${day}`;
  }

  // 4. Fallback: Parse with Date object
  const parsed = new Date(trimmed);
  if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1990 && parsed.getFullYear() < 2100) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const day = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  return null;
}

/**
 * Checks if a string looks like an ID column header.
 */
function isIdColumn(colName: string): boolean {
  if (!colName) return false;
  const clean = colName.toLowerCase().replace(/[^a-z0-9#]/g, '');
  const idKeywords = [
    'idnumber', 'studentid', '#id', 'id#', 'id', 'studentnumber',
    'student#', '#student', 'rollno', 'rollnumber', 'roll#',
    'matriculation', 'matricule', 'regno', 'registrationnumber',
    'registration', 'code', 'academid', 'usercode', 'no', '#'
  ];
  if (idKeywords.includes(clean)) return true;
  if (/^id\b/i.test(colName.trim()) || /\bid$/i.test(colName.trim()) || /\bid\b/i.test(colName.trim())) return true;
  if (clean.includes('studentid') || clean.includes('idnumber')) return true;
  return false;
}

/**
 * Checks if a column header designates a username, email handle, account, or login credential.
 * These must NEVER be chosen as the official student name column.
 */
export function isUsernameOrEmailHeader(header: string): boolean {
  if (!header) return false;
  const clean = header.toLowerCase().replace(/[^a-z0-9]/g, '');
  const usernameKeywords = [
    'username', 'user', 'login', 'loginname', 'loginid', 'userlogin',
    'email', 'mail', 'emailaddress', 'useremail', 'studentemail',
    'account', 'useraccount', 'handle', 'netid', 'sso', 'systemuser',
    'alias', 'screenname', 'nickname', 'slack', 'discord', 'github'
  ];
  if (usernameKeywords.includes(clean)) return true;
  if (/^(?:user\s*name|login\s*name|email|user\s*account|net\s*id)$/i.test(header.trim())) return true;
  if (clean.includes('email') || clean.includes('username') || clean.includes('userlogin') || clean.includes('loginname')) return true;
  return false;
}

/**
 * Checks if a string value looks like a username, email address, or system handle
 * (e.g. contains '@', contains underscores '_', contains internal dots like 'john.doe',
 * or has trailing system numbers).
 */
export function isLikelyUsernameOrEmail(val: string): boolean {
  if (!val) return false;
  const s = String(val).trim();
  if (!s) return false;

  // 1. Contains @ (full email address e.g. "john.doe@gmail.com")
  if (s.includes('@')) return true;

  // 2. Contains underscores (e.g. "john_doe", "sophia_alvarez_01")
  if (s.includes('_')) return true;

  // 3. Contains dots within words with no space (e.g. "john.doe", "b.13.bad50", "s.alvarez")
  if (/^[a-zA-Z0-9_-]+\.[a-zA-Z0-9_.-]+$/.test(s)) return true;

  // 4. Multiple dots without spaces
  if ((s.match(/\./g) || []).length >= 2 && !s.includes(' ')) return true;

  // 5. Ends with digits or embedded digits without space (e.g. "johndoe123", "user42", "b13bad50")
  if (/^[a-zA-Z]+[0-9]{2,}$/.test(s) || /^[a-zA-Z0-9_-]+[0-9]{3,}$/.test(s)) return true;

  // 6. Contains email domain extensions
  if (/\.(com|edu|org|net|gov|mil|io|ai|co|uk|ca|de|fr|ru|es|me)$/i.test(s)) return true;

  return false;
}

/**
 * Transforms any raw name or username-style string into a proper, clean government/legal name:
 * - Reverses "Last, First" formats to "First Last"
 * - Converts dots and underscores into natural spaces ("sophia.alvarez" -> "Sophia Alvarez")
 * - Strips any email domain portion ("sophia.alvarez@school.edu" -> "Sophia Alvarez")
 * - Removes random trailing system digits ("sophia_alvarez99" -> "Sophia Alvarez")
 * - Capitalizes every name word accurately (preserving hyphens and apostrophes like "O'Connor", "Mary-Jane")
 */
export function formatToProperName(rawName: string): string {
  if (!rawName) return '';
  let str = String(rawName).trim();
  if (!str) return '';

  // 1. If it's an email address, strip the @domain part
  if (str.includes('@')) {
    str = str.split('@')[0].trim();
  }

  // 2. Handle "Last, First" or "Last, First Middle" format (common in official school exports)
  if (str.includes(',') && !str.includes('.') && !str.includes('_')) {
    const commaParts = str.split(',').map(p => p.trim()).filter(Boolean);
    if (commaParts.length >= 2) {
      str = `${commaParts[1]} ${commaParts[0]}`;
    }
  }

  // 3. Remove dots and underscores (e.g. "john.doe" or "john_doe" or "b.13.bad50")
  if (str.includes('.') || str.includes('_')) {
    str = str.replace(/[._]+/g, ' ');
  }

  // 4. Strip trailing system numbers (e.g. "alvarez01" -> "alvarez")
  str = str.replace(/\b([a-zA-Z]+)\d+\b/g, '$1');

  // 5. Clean up any remaining non-name punctuation (keep hyphens and apostrophes)
  str = str.replace(/[^a-zA-Z\s\-']/g, ' ');

  // 6. Split words and capitalize properly
  const words = str
    .split(/\s+/)
    .map(w => w.trim())
    .filter(Boolean);

  if (words.length === 0) return rawName.trim();

  const formattedWords = words.map(w => {
    // Handle apostrophes like O'Connor or D'Angelo
    if (w.includes("'")) {
      return w.split("'").map(part => {
        if (!part) return '';
        return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
      }).join("'");
    }
    // Handle hyphens like Mary-Jane
    if (w.includes("-")) {
      return w.split("-").map(part => {
        if (!part) return '';
        return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
      }).join("-");
    }
    if (w.length <= 1) return w.toUpperCase();
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  });

  return formattedWords.join(' ');
}

/**
 * Clean and identify name columns, prioritizing genuine government/legal names
 * and rejecting email/username columns.
 */
interface NameColumnMapping {
  singleFullNameIdx: number;
  firstNameIdx: number;
  middleNameIdx: number;
  lastNameIdx: number;
  additionalNameIndices: number[];
}

function resolveNameColumns(headers: string[], sampleRows?: any[][]): NameColumnMapping {
  let singleFullNameIdx = -1;
  let firstNameIdx = -1;
  let middleNameIdx = -1;
  let lastNameIdx = -1;
  const additionalNameIndices: number[] = [];

  const firstKeywords = ['firstname', 'first', 'givenname', 'given', 'forename', 'prenom'];
  const middleKeywords = ['middlename', 'middle', 'secondname', 'mi', 'middleinitial'];
  const lastKeywords = ['lastname', 'last', 'surname', 'familyname', 'nomdefamille', 'apellido'];
  
  // Prioritized government/official name keywords
  const governmentKeywords = ['governmentname', 'govermentname', 'legalname', 'officialname'];
  const fullKeywords = ['fullname', 'studentfullname', 'studentname', 'formalname', 'actualname', 'realname', 'name', 'nom', 'nombre', 'learner', 'student'];

  headers.forEach((header, idx) => {
    if (!header) return;

    // RULE 1: Never select username or email headers as a name column!
    if (isUsernameOrEmailHeader(header)) {
      return;
    }

    const clean = header.toLowerCase().replace(/[^a-z0-9]/g, '');

    // Check if column sample values look like usernames/emails
    let usernameScore = 0;
    if (sampleRows && sampleRows.length > 0) {
      const nonEmpties = sampleRows.map(r => String(r?.[idx] ?? '').trim()).filter(Boolean);
      if (nonEmpties.length > 0) {
        const usernameCount = nonEmpties.filter(isLikelyUsernameOrEmail).length;
        usernameScore = usernameCount / nonEmpties.length;
      }
    }

    // If more than 40% of rows look like usernames or emails, do NOT pick as a name column!
    if (usernameScore > 0.4) {
      return;
    }

    // Check explicit government name keywords first
    if (governmentKeywords.includes(clean)) {
      singleFullNameIdx = idx;
      return;
    }

    if (firstKeywords.includes(clean)) {
      if (firstNameIdx === -1) firstNameIdx = idx;
    } else if (middleKeywords.includes(clean)) {
      if (middleNameIdx === -1) middleNameIdx = idx;
    } else if (lastKeywords.includes(clean)) {
      if (lastNameIdx === -1) lastNameIdx = idx;
    } else if (fullKeywords.includes(clean)) {
      if (singleFullNameIdx === -1) {
        singleFullNameIdx = idx;
      } else {
        // Prefer more descriptive headers like "Full Name" or "Student Name" over generic "Name"
        if (clean.includes('fullname') || clean.includes('studentname') || clean.includes('studentfullname')) {
          singleFullNameIdx = idx;
        }
      }
    } else if (clean.includes('name') && !clean.includes('class') && !clean.includes('course') && !clean.includes('school')) {
      additionalNameIndices.push(idx);
    }
  });

  return { singleFullNameIdx, firstNameIdx, middleNameIdx, lastNameIdx, additionalNameIndices };
}

/**
 * Check if a column header represents an attendance date.
 */
function resolveAttendanceDate(header: string, sampleRowValues: string[]): string | null {
  if (!header) return null;
  const trimmed = header.trim();
  if (!trimmed) return null;
  const lower = trimmed.toLowerCase();
  if (
    lower.includes('final grade') ||
    lower.includes('grade') ||
    lower.includes('id') ||
    lower.includes('name') ||
    lower.includes('lecture') ||
    lower.includes('summary') ||
    lower.includes('overview') ||
    lower.includes('total') ||
    lower.includes('attended') ||
    lower.includes('status')
  ) {
    return null;
  }

  // Check if header is a date
  const parsedDate = parseDateToStandardYMD(trimmed);
  if (parsedDate) {
    return parsedDate;
  }

  // Check sample values in this column for attendance keywords
  const attendanceKeywords = ['attended', 'absent', 'not taken', 'present', 'p', 'a'];
  const hasAttendanceValues = sampleRowValues.some(v => 
    attendanceKeywords.includes(String(v || '').toLowerCase().trim())
  );
  if (hasAttendanceValues) {
    // Try to parse header with relaxed parsing
    const cleaned = header.replace(/[^0-9/.-]/g, '');
    const tryDate = parseDateToStandardYMD(cleaned);
    if (tryDate) return tryDate;
  }

  return null;
}

/**
 * Checks if a column header represents an assessment column.
 * Matches: 'Quiz 1 (Max 100, 30%)' or any general test/quiz/assignment column.
 */
interface ParsedAssessmentHeader {
  name: string;
  maxMark: number;
  weight: number;
}

function resolveAssessmentColumn(header: string, sampleRowValues: string[]): ParsedAssessmentHeader | null {
  if (!header) return null;
  const trimmed = header.trim();
  if (!trimmed) return null;
  const lower = trimmed.toLowerCase();
  if (
    lower === 'final grade' ||
    lower.includes('id') ||
    lower.includes('name') ||
    lower.includes('attendance') ||
    lower.includes('lecture') ||
    lower.includes('summary') ||
    lower.includes('overview') ||
    lower.includes('status')
  ) {
    return null;
  }

  // Check if header matches our exported assessment format: "Name (Max 100, 30%)"
  const exportedMatch = trimmed.match(/^(.*?)\s*\((?:max|full mark)?\s*(\d+)\s*,\s*(\d+)%?\)$/i);
  if (exportedMatch) {
    return {
      name: exportedMatch[1].trim(),
      maxMark: parseInt(exportedMatch[2], 10) || 100,
      weight: parseInt(exportedMatch[3], 10) || 0
    };
  }

  // Check if sample values are numeric scores
  const numericCount = sampleRowValues.filter(v => {
    if (v === '' || v === null || v === undefined) return false;
    const num = parseFloat(String(v).trim());
    return !isNaN(num) && num >= 0 && num <= 1000;
  }).length;

  if (numericCount > 0 && numericCount >= Math.floor(sampleRowValues.length * 0.3)) {
    // Treat as an assessment column
    return {
      name: trimmed,
      maxMark: 100,
      weight: 0
    };
  }

  return null;
}

/**
 * Parses a 2D array of rows representing a class section table.
 */
function parseSectionMatrix(
  rows: any[][],
  defaultClassName: string
): ParsedClassSectionData | null {
  if (!rows || rows.length === 0) return null;

  // 1. Locate the header row (contains at least one ID and/or Name column)
  let headerRowIdx = -1;
  let idColIdx = -1;
  let nameMapping: NameColumnMapping = {
    singleFullNameIdx: -1,
    firstNameIdx: -1,
    middleNameIdx: -1,
    lastNameIdx: -1,
    additionalNameIndices: []
  };

  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;

    const strHeaders = row.map(c => String(c ?? '').trim());
    let foundId = -1;
    strHeaders.forEach((h, colIdx) => {
      if (foundId === -1 && isIdColumn(h)) {
        foundId = colIdx;
      }
    });

    const sampleData = rows.slice(r + 1, r + 15);
    const mapping = resolveNameColumns(strHeaders, sampleData);
    const hasName = mapping.singleFullNameIdx !== -1 ||
                    mapping.firstNameIdx !== -1 ||
                    mapping.lastNameIdx !== -1 ||
                    mapping.additionalNameIndices.length > 0;

    if (foundId !== -1 || hasName) {
      headerRowIdx = r;
      idColIdx = foundId;
      nameMapping = mapping;
      break;
    }
  }

  if (headerRowIdx === -1) {
    // Fallback: If no headers matched by name, assume Row 0: Col 0 is ID, Col 1 is Name
    headerRowIdx = 0;
    idColIdx = 0;
    nameMapping.singleFullNameIdx = 1;
  }

  const rawHeaders = rows[headerRowIdx].map(c => String(c ?? '').trim());

  // If ID column still not located, pick the first column that has mostly short strings or numbers
  if (idColIdx === -1) {
    idColIdx = 0;
  }

  // Sample data rows to detect attendance date columns & assessment columns
  const dataRows = rows.slice(headerRowIdx + 1);
  const sampleValuesPerCol: string[][] = rawHeaders.map((_, colIdx) => 
    dataRows.slice(0, 10).map(r => String(r?.[colIdx] ?? '').trim()).filter(Boolean)
  );

  const attendanceDateCols: { colIdx: number; dateYMD: string }[] = [];
  const assessmentCols: { colIdx: number; meta: ParsedAssessmentHeader }[] = [];

  rawHeaders.forEach((header, colIdx) => {
    if (colIdx === idColIdx) return;
    if (colIdx === nameMapping.singleFullNameIdx ||
        colIdx === nameMapping.firstNameIdx ||
        colIdx === nameMapping.middleNameIdx ||
        colIdx === nameMapping.lastNameIdx ||
        nameMapping.additionalNameIndices.includes(colIdx)) {
      return;
    }

    const dateYMD = resolveAttendanceDate(header, sampleValuesPerCol[colIdx] || []);
    if (dateYMD) {
      attendanceDateCols.push({ colIdx, dateYMD });
      return;
    }

    const assessMeta = resolveAssessmentColumn(header, sampleValuesPerCol[colIdx] || []);
    if (assessMeta) {
      assessmentCols.push({ colIdx, meta: assessMeta });
    }
  });

  const students: Student[] = [];
  const attendance: Record<string, Record<string, 'Not Taken' | 'Attended' | 'Absent'>> = {};
  const grades: Record<string, Record<string, number>> = {};
  const seenStudentIds = new Set<string>();

  dataRows.forEach(row => {
    if (!row || !Array.isArray(row)) return;

    // Check if row is a section break or empty
    const firstCell = String(row[0] ?? '').trim();
    if (firstCell.startsWith('🏫') || firstCell.toLowerCase().startsWith('class section')) {
      return;
    }

    const rawId = String(row[idColIdx] ?? '').trim();
    if (!rawId) return;

    // Extract Name
    let fullName = '';
    const hasExplicitParts = nameMapping.firstNameIdx !== -1 || nameMapping.lastNameIdx !== -1;

    if (hasExplicitParts) {
      const parts: string[] = [];
      if (nameMapping.firstNameIdx !== -1) parts.push(String(row[nameMapping.firstNameIdx] ?? '').trim());
      if (nameMapping.middleNameIdx !== -1) parts.push(String(row[nameMapping.middleNameIdx] ?? '').trim());
      if (nameMapping.lastNameIdx !== -1) parts.push(String(row[nameMapping.lastNameIdx] ?? '').trim());
      nameMapping.additionalNameIndices.forEach(idx => {
        parts.push(String(row[idx] ?? '').trim());
      });
      fullName = parts.filter(Boolean).join(' ');
    } else if (nameMapping.singleFullNameIdx !== -1) {
      fullName = String(row[nameMapping.singleFullNameIdx] ?? '').trim();
    } else if (nameMapping.additionalNameIndices.length > 0) {
      fullName = nameMapping.additionalNameIndices
        .map(idx => String(row[idx] ?? '').trim())
        .filter(Boolean)
        .join(' ');
    }

    // If full name is empty or looks like an email/username handle, check other columns for a genuine government name
    if (!fullName || isLikelyUsernameOrEmail(fullName)) {
      for (let c = 0; c < row.length; c++) {
        if (c === idColIdx) continue;
        const colHeader = rawHeaders[c] || '';
        if (isUsernameOrEmailHeader(colHeader)) continue;

        const cellText = String(row[c] ?? '').trim();
        if (cellText && !isLikelyUsernameOrEmail(cellText)) {
          fullName = cellText;
          break;
        }
      }

      // If still empty, try column adjacent to ID as last resort
      if (!fullName) {
        const altIdx = idColIdx === 0 ? 1 : 0;
        fullName = String(row[altIdx] ?? '').trim();
      }
    }

    if (!fullName) return; // Skip rows without name

    // Convert into a clean, proper government name (strips email domains, converts dots/underscores to spaces, capitalizes)
    const cleanGovernmentName = formatToProperName(fullName);
    if (!cleanGovernmentName) return;

    // Avoid duplicate IDs within the same section
    let studentId = rawId;
    if (seenStudentIds.has(studentId)) {
      studentId = `${rawId}-${students.length + 1}`;
    }
    seenStudentIds.add(studentId);

    const nameParts = cleanGovernmentName.split(/\s+/);
    const firstName = nameParts[0] || cleanGovernmentName;
    const lastName = nameParts.slice(1).join(' ');

    students.push({
      id: studentId,
      firstName,
      lastName,
      fullName: cleanGovernmentName
    });

    // Populate Attendance
    if (attendanceDateCols.length > 0) {
      attendance[studentId] = {};
      attendanceDateCols.forEach(({ colIdx, dateYMD }) => {
        const rawStatus = String(row[colIdx] ?? '').trim().toLowerCase();
        let status: 'Not Taken' | 'Attended' | 'Absent' = 'Not Taken';
        if (rawStatus === 'attended' || rawStatus === 'present' || rawStatus === 'p' || rawStatus === 'yes' || rawStatus === '1') {
          status = 'Attended';
        } else if (rawStatus === 'absent' || rawStatus === 'a' || rawStatus === 'no' || rawStatus === '0') {
          status = 'Absent';
        }
        attendance[studentId][dateYMD] = status;
      });
    }

    // Populate Grades
    if (assessmentCols.length > 0) {
      grades[studentId] = {};
      assessmentCols.forEach(({ colIdx, meta }) => {
        const cellRaw = row[colIdx];
        if (cellRaw !== undefined && cellRaw !== null && cellRaw !== '') {
          const num = parseFloat(String(cellRaw).trim());
          if (!isNaN(num)) {
            grades[studentId][meta.name] = num;
          }
        }
      });
    }
  });

  if (students.length === 0) return null;

  // Sort students alphabetically by firstName
  students.sort((a, b) => a.firstName.localeCompare(b.firstName));

  const attendanceDates = attendanceDateCols.map(c => c.dateYMD);
  const assessments = assessmentCols.map(c => c.meta);

  return {
    className: defaultClassName,
    students,
    attendanceDates,
    attendance,
    assessments,
    grades
  };
}

/**
 * Universal Excel / CSV Parser
 * Automatically detects:
 * - Single roster files with arbitrary column names (Name, Full Name, First+Last, Student ID, #ID, Roll No, etc.)
 * - Our exported Attendance Logs (restoring roster and date columns)
 * - Our exported Grades Reports (restoring assessments, weights, and scores)
 * - Merged files for all classes (detects multiple sections separated by banners)
 * - Multi-sheet Excel files (each sheet treated as a section)
 */
export async function parseAnyExcelFile(
  file: File,
  activeClassName?: string
): Promise<ParseExcelResult> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });

  const sections: ParsedClassSectionData[] = [];

  // Check all sheets in workbook
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    // Convert sheet to 2D array of rows
    const sheetAoA = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: '' });
    if (!sheetAoA || sheetAoA.length === 0) continue;

    // Check if sheet contains multiple merged sections indicated by section banners (e.g. 🏫 Class Section A)
    const sectionSplits: { startIndex: number; className: string }[] = [];
    sheetAoA.forEach((row, rowIdx) => {
      if (!row || !Array.isArray(row)) return;
      const cell0 = String(row[0] ?? '').trim();
      const hasOtherCells = row.slice(1).some(c => String(c ?? '').trim() !== '');

      // Identify section header like: "🏫 Class Section A" or "Class Section A"
      if (!hasOtherCells && (cell0.startsWith('🏫') || /^(?:class\s+section|section\s+[a-z0-9]|class\s*:)/i.test(cell0))) {
        const cleanName = cell0.replace(/^🏫\s*/, '').replace(/^class\s*:\s*/i, '').trim();
        sectionSplits.push({ startIndex: rowIdx, className: cleanName });
      }
    });

    if (sectionSplits.length > 1) {
      // Multiple sections detected within this single sheet!
      for (let i = 0; i < sectionSplits.length; i++) {
        const curr = sectionSplits[i];
        const next = sectionSplits[i + 1];
        const sectionRows = sheetAoA.slice(curr.startIndex + 1, next ? next.startIndex : undefined);
        const parsed = parseSectionMatrix(sectionRows, curr.className);
        if (parsed && parsed.students.length > 0) {
          sections.push(parsed);
        }
      }
    } else {
      // Single section in this sheet
      let fallbackName = activeClassName || 'Imported Section';
      if (sectionSplits.length === 1) {
        fallbackName = sectionSplits[0].className;
      } else if (workbook.SheetNames.length > 1) {
        const cleanSheet = sheetName.trim();
        const facetMatch = cleanSheet.match(/^(.*?)\s*[-_–:(]\s*(?:attendance|grades|assessments|roster|students|full\s*info|ledger|overview|scores)\)?$/i);
        if (facetMatch && facetMatch[1].trim()) {
          fallbackName = facetMatch[1].trim();
        } else if (/^(?:attendance|grades|assessments|roster|students|full\s*info|ledger|overview|scores|sheet\d+)$/i.test(cleanSheet)) {
          fallbackName = activeClassName || 'Imported Section';
        } else {
          fallbackName = cleanSheet;
        }
      }

      // If the first row was a banner, skip it
      const startIdx = sectionSplits.length === 1 ? sectionSplits[0].startIndex + 1 : 0;
      const parsed = parseSectionMatrix(sheetAoA.slice(startIdx), fallbackName);
      if (parsed && parsed.students.length > 0) {
        sections.push(parsed);
      }
    }
  }

  if (sections.length === 0) {
    throw new Error("Could not detect any valid student records. Please ensure your file contains student IDs and names.");
  }

  // Consolidate sections that share the same class name (e.g. from multiple sheets, facets, or duplicates)
  const consolidatedMap = new Map<string, ParsedClassSectionData>();

  sections.forEach(sec => {
    // Clean up class name (strip facet suffixes like " - Attendance", " (Grades)", " - Full Info", etc.)
    let baseName = sec.className.trim();
    const facetSuffixMatch = baseName.match(/^(.*?)\s*[-_–:(]\s*(?:attendance|grades|assessments|roster|students|full\s*info|ledger|overview|scores)\)?$/i);
    if (facetSuffixMatch && facetSuffixMatch[1].trim()) {
      baseName = facetSuffixMatch[1].trim();
    } else if (/^(?:attendance|grades|assessments|roster|students|full\s*info|ledger|overview|scores|sheet\d+)$/i.test(baseName)) {
      baseName = activeClassName || 'Imported Section';
    }

    const key = baseName.toLowerCase();
    if (!consolidatedMap.has(key)) {
      consolidatedMap.set(key, {
        className: baseName,
        students: [...sec.students],
        attendanceDates: [...sec.attendanceDates],
        attendance: { ...sec.attendance },
        assessments: [...sec.assessments],
        grades: { ...sec.grades }
      });
    } else {
      const existing = consolidatedMap.get(key)!;
      // 1. Merge students
      const stMap = new Map<string, Student>();
      existing.students.forEach(s => stMap.set(s.id, s));
      sec.students.forEach(s => {
        if (stMap.has(s.id)) {
          stMap.set(s.id, { ...stMap.get(s.id)!, ...s });
        } else {
          stMap.set(s.id, s);
        }
      });
      existing.students = Array.from(stMap.values()).sort((a, b) => a.firstName.localeCompare(b.firstName));

      // 2. Merge attendance dates
      const mergedDates = Array.from(new Set([...existing.attendanceDates, ...sec.attendanceDates]));
      existing.attendanceDates = mergedDates;

      // 3. Merge attendance records
      if (sec.attendance) {
        existing.attendance = existing.attendance || {};
        Object.keys(sec.attendance).forEach(stId => {
          existing.attendance[stId] = {
            ...(existing.attendance[stId] || {}),
            ...(sec.attendance[stId] || {})
          };
        });
      }

      // 4. Merge assessments
      const assessMap = new Map<string, Assessment>();
      existing.assessments.forEach(a => assessMap.set(a.name.toLowerCase(), a));
      sec.assessments.forEach(a => {
        const aKey = a.name.toLowerCase();
        if (!assessMap.has(aKey) || (a.weight > 0 && assessMap.get(aKey)!.weight === 0)) {
          assessMap.set(aKey, a);
        }
      });
      existing.assessments = Array.from(assessMap.values());

      // 5. Merge grades
      if (sec.grades) {
        existing.grades = existing.grades || {};
        Object.keys(sec.grades).forEach(stId => {
          existing.grades[stId] = {
            ...(existing.grades[stId] || {}),
            ...(sec.grades[stId] || {})
          };
        });
      }
    }
  });

  const finalSections = Array.from(consolidatedMap.values());
  const isMultiSection = finalSections.length > 1;
  const totalStudents = finalSections.reduce((sum, s) => sum + s.students.length, 0);

  let summary = '';
  if (isMultiSection) {
    summary = `Identified ${finalSections.length} class sections (${totalStudents} total students): ${finalSections.map(s => `${s.className} (${s.students.length})`).join(', ')}.`;
  } else {
    const s = finalSections[0];
    const details: string[] = [`${s.students.length} students`];
    if (s.attendanceDates.length > 0) details.push(`${s.attendanceDates.length} attendance dates`);
    if (s.assessments.length > 0) details.push(`${s.assessments.length} assessments`);
    summary = `Parsed ${s.className}: ${details.join(', ')}.`;
  }

  return {
    isMultiSection,
    sections: finalSections,
    summary,
    totalStudents
  };
}

/**
 * Calculates optimal column widths for SheetJS worksheets so that
 * long student names, IDs, dates, and test titles never truncate or require double clicking.
 */
export function applyAutoColumnWidths(
  worksheet: XLSX.WorkSheet,
  aoa: (string | number | null | undefined)[][]
): void {
  const colWidths: { wch: number }[] = [];

  const maxCols = Math.max(0, ...aoa.map(r => (Array.isArray(r) ? r.length : 0)));

  aoa.forEach(row => {
    if (!row || !Array.isArray(row)) return;
    row.forEach((cellVal, colIdx) => {
      if (cellVal === null || cellVal === undefined) return;
      const strVal = String(cellVal).trim();
      // Skip wide banner headers from blowing up column 0
      if (colIdx === 0 && (strVal.startsWith('🏫') || /^(?:class\s+section|merged)/i.test(strVal))) {
        return;
      }
      const len = strVal.length;
      if (!colWidths[colIdx]) {
        colWidths[colIdx] = { wch: 0 };
      }
      if (len > colWidths[colIdx].wch) {
        colWidths[colIdx].wch = len;
      }
    });
  });

  const formattedCols: { wch: number }[] = [];
  for (let idx = 0; idx < maxCols; idx++) {
    const c = colWidths[idx];
    // Empty distinction separator column between grades and attendance
    if (!c || c.wch === 0) {
      formattedCols.push({ wch: 4 });
      continue;
    }
    let width = c.wch + 4;
    // ID column
    if (idx === 0) width = Math.max(width, 16);
    // Student Name column - guarantee at least 30 chars width so long names fit easily
    else if (idx === 1) width = Math.max(width, 30);
    // Grade / assessment / attendance summary / date columns
    else width = Math.max(width, 14);

    formattedCols.push({ wch: Math.min(width, 50) });
  }

  worksheet['!cols'] = formattedCols;
}

/**
 * Convenience helper to apply auto column widths to json_to_sheet outputs.
 */
export function applyAutoColumnWidthsToTable(
  worksheet: XLSX.WorkSheet,
  tableData: Record<string, any>[]
): void {
  if (!tableData || tableData.length === 0) return;
  const headers = Object.keys(tableData[0]);
  const aoa: any[][] = [headers];
  tableData.forEach(row => {
    aoa.push(headers.map(h => row[h]));
  });
  applyAutoColumnWidths(worksheet, aoa);
}
