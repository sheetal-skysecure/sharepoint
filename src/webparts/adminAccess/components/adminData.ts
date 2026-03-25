export interface UserEnrollment {
    id: string;
    userId: string;
    userName: string;
    userEmail: string;
    courseId: string;
    courseName: string;
    provider: string;
    status: 'scheduled' | 'completed' | 'in-progress' | 'failed';
    date: string;
    department: 'IT' | 'Sales' | 'Marketing' | 'HR' | 'Security';
    score?: number;
}

export const mockEnrollments: UserEnrollment[] = [
    { id: '1', userId: 'u1', userName: 'Sheetal Sinha', userEmail: 'sheetal@example.com', courseId: 'AZ-104', courseName: 'Azure Administrator Associate', provider: 'Microsoft', status: 'in-progress', date: '2024-03-05', department: 'IT' },
    { id: '2', userId: 'u2', userName: 'John Doe', userEmail: 'john.doe@example.com', courseId: 'AZ-104', courseName: 'Azure Administrator Associate', provider: 'Microsoft', status: 'completed', date: '2024-02-15', score: 850, department: 'IT' },
    { id: '3', userId: 'u3', userName: 'Jane Smith', userEmail: 'jane.smith@example.com', courseId: 'MS-700', courseName: 'Teams Administrator Associate', provider: 'Microsoft', status: 'scheduled', date: '2024-04-10', department: 'Sales' },
    { id: '4', userId: 'u4', userName: 'Robert Brown', userEmail: 'robert.b@example.com', courseId: 'AZ-104', courseName: 'Azure Administrator Associate', provider: 'Microsoft', status: 'in-progress', date: '2024-03-01', department: 'IT' },
    { id: '5', userId: 'u5', userName: 'Alice Wilson', userEmail: 'alice.w@example.com', courseId: 'DP-203', courseName: 'Azure Data Engineer Associate', provider: 'Microsoft', status: 'completed', date: '2024-01-20', score: 920, department: 'Marketing' },
    { id: '6', userId: 'u1', userName: 'Sheetal Sinha', userEmail: 'sheetal@example.com', courseId: 'SC-300', courseName: 'Identity and Access Administrator', provider: 'Microsoft', status: 'scheduled', date: '2024-05-12', department: 'IT' },
    { id: '7', userId: 'u6', userName: 'Kevin Lee', userEmail: 'kevin.l@example.com', courseId: 'AZ-104', courseName: 'Azure Administrator Associate', provider: 'Microsoft', status: 'failed', date: '2024-02-28', score: 640, department: 'Security' },
    { id: '8', userId: 'u7', userName: 'Sarah Connor', userEmail: 's.connor@example.com', courseId: 'MS-102', courseName: 'Microsoft 365 Administrator Expert', provider: 'Microsoft', status: 'in-progress', date: '2024-03-15', department: 'Security' },
    { id: '9', userId: 'u8', userName: 'Thomas Muller', userEmail: 't.m@example.com', courseId: 'AZ-104', courseName: 'Azure Administrator Associate', provider: 'Microsoft', status: 'completed', date: '2023-12-10', score: 780, department: 'HR' },
    { id: '10', userId: 'u2', userName: 'John Doe', userEmail: 'john.doe@example.com', courseId: 'SC-200', courseName: 'Security Operations Analyst', provider: 'Microsoft', status: 'scheduled', date: '2024-06-01', department: 'IT' },
];
