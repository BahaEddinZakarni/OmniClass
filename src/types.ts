export interface Student {
  id: string;
  firstName: string;
  lastName: string;
  fullName: string;
  color?: string; // Predetermined background highlight color
}

export interface Assessment {
  name: string;
  weight: number; // e.g. 20
}

export interface ClassSection {
  name: string;
  students: Student[]; // Sorted alphabetically by firstName
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
