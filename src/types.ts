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
}
