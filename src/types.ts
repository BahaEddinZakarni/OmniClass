export interface Student {
  id: string;
  firstName: string;
  lastName: string;
  fullName: string;
  color?: string; // Predetermined background highlight color
}

export interface Assessment {
  name: string;
  weight: number; // Percentage out of 100%
  maxMark?: number; // Full mark of assessment (e.g. 20, 50, 100)
}

export type DayOfWeek = 'Monday' | 'Tuesday' | 'Wednesday' | 'Thursday' | 'Friday' | 'Saturday' | 'Sunday';

export interface ClassScheduleSlot {
  id?: string;
  day: DayOfWeek;
  startTime: string; // e.g. "09:00 AM"
  endTime: string;   // e.g. "10:30 AM"
  room?: string;
}

export interface ClassSection {
  name: string;
  students: Student[]; // Sorted alphabetically by firstName
  attendanceDates?: string[]; // Preserves column ordering (newest next to names)
  // Indexed by student.id -> { dateString: status }
  attendance: {
    [studentId: string]: {
      [date: string]: 'Not Taken' | 'Attended' | 'Absent';
    };
  };
  assessments: Assessment[];
  // Indexed by student.id -> { assessmentName: score }
  grades: {
    [studentId: string]: {
      [assessmentName: string]: number;
    };
  };
  schedule?: ClassScheduleSlot[];
}

export type CalendarEventType = 'exam' | 'meeting' | 'lecture' | 'deadline' | 'note';

export interface CalendarEvent {
  id: string;
  date: string; // Canonical YYYY-MM-DD
  title: string;
  time?: string; // e.g. "10:00 AM" or "All Day"
  className?: string; // Class name or empty/General
  type: CalendarEventType;
  description?: string;
  color?: string;
  createdAt?: string | number;
}

