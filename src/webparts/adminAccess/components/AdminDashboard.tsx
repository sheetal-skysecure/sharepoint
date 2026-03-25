import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { IAdminDashboardProps } from './IAdminDashboardProps';
import { mockEnrollments } from './adminData';
import {
    Users,
    BookOpen,
    CheckCircle,
    TrendingUp,
    Search,
    Download,
    Mail,
    MoreVertical,
    ChevronRight,
    Clock,
    Activity,
    AlertCircle,
    Check
} from 'lucide-react';

export default function AdminDashboard(props: IAdminDashboardProps) {
    const navigate = useNavigate();
    const [searchTerm, setSearchTerm] = useState('');
    const [statusFilter, setStatusFilter] = useState('all');
    const [deptFilter, setDeptFilter] = useState('all');
    const [currentTime, setCurrentTime] = useState(new Date());
    const [lastSync, setLastSync] = useState(new Date());
    const [notifications] = useState<{ id: number, text: string, time: string, type: 'info' | 'success' | 'alert' }[]>([
        { id: 1, text: "Sheetal Sinha started SC-300", time: "2 mins ago", type: 'info' },
        { id: 2, text: "John Doe completed AZ-104", time: "15 mins ago", type: 'success' },
        { id: 3, text: "Kevin Lee failed assessment", time: "1 hour ago", type: 'alert' }
    ]);

    // Real-time updates simulation
    useEffect(() => {
        const timer = setInterval(() => setCurrentTime(new Date()), 1000);
        const syncTimer = setInterval(() => setLastSync(new Date()), 30000); // Sync every 30s

        return () => {
            clearInterval(timer);
            clearInterval(syncTimer);
        };
    }, []);

    const filteredEnrollments = useMemo(() => {
        return mockEnrollments.filter(e => {
            const matchesSearch =
                e.userName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                e.courseName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                e.userEmail.toLowerCase().includes(searchTerm.toLowerCase());

            const matchesStatus = statusFilter === 'all' || e.status === statusFilter;
            const matchesDept = deptFilter === 'all' || e.department === deptFilter;

            return matchesSearch && matchesStatus && matchesDept;
        });
    }, [searchTerm, statusFilter, deptFilter]);

    // Analytics Calculations
    const totalUsers = new Set(mockEnrollments.map(e => e.userId)).size;
    const totalEnrollments = mockEnrollments.length;
    const completedCount = mockEnrollments.filter(e => e.status === 'completed').length;
    const completionRate = Math.round((completedCount / totalEnrollments) * 100);

    const deptStats = useMemo(() => {
        const stats: { [key: string]: number } = {};
        mockEnrollments.forEach(e => {
            stats[e.department] = (stats[e.department] || 0) + 1;
        });
        return stats;
    }, []);

    const courseStats = useMemo(() => {
        const stats: { [key: string]: { name: string, count: number, completed: number } } = {};
        mockEnrollments.forEach(e => {
            if (!stats[e.courseId]) {
                stats[e.courseId] = { name: e.courseName, count: 0, completed: 0 };
            }
            stats[e.courseId].count++;
            if (e.status === 'completed') stats[e.courseId].completed++;
        });
        return Object.keys(stats).map((key: string) => stats[key]).sort((a: any, b: any) => b.count - a.count);
    }, []);

    const handleSendReminder = (email: string) => {
        alert(`Reminder sent to ${email}`);
    };

    return (
        <div style={{
            padding: '2rem',
            backgroundColor: '#f1f5f9',
            minHeight: '100vh',
            fontFamily: "'Inter', sans-serif",
            color: '#1e293b'
        }}>
            {/* Top Bar with Real-time Clock */}
            <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: '2rem',
                backgroundColor: 'white',
                padding: '1rem 2rem',
                borderRadius: '16px',
                boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                    <div style={{ padding: '0.5rem', backgroundColor: '#e2e8f0', borderRadius: '8px' }}>
                        <Clock size={20} color="#64748b" />
                    </div>
                    <div>
                        <div style={{ fontSize: '0.85rem', color: '#64748b', fontWeight: 600 }}>System Time</div>
                        <div style={{ fontSize: '1rem', fontWeight: 700 }}>{currentTime.toLocaleTimeString()}</div>
                    </div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '2rem' }}>
                    <div style={{ textAlign: 'right' }}>
                        <div style={{ fontSize: '0.85rem', color: '#64748b', fontWeight: 600 }}>Last Data Sync</div>
                        <div style={{ fontSize: '0.9rem', fontWeight: 700, color: '#10b981' }}>
                            {Math.floor((currentTime.getTime() - lastSync.getTime()) / 1000)}s ago
                        </div>
                    </div>
                    <div style={{ width: '1px', height: '30px', backgroundColor: '#e2e8f0' }}></div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                        <div style={{ textAlign: 'right' }}>
                            <div style={{ fontSize: '0.9rem', fontWeight: 700 }}>Admin Portal</div>
                            <div style={{ fontSize: '0.75rem', color: '#64748b' }}>v2.4.0-alpha</div>
                        </div>
                    </div>
                </div>
            </div>

            {/* Header */}
            <header style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: '2.5rem'
            }}>
                <div>
                    <h1 style={{ fontSize: '2.25rem', fontWeight: 800, color: '#0f172a', margin: 0, letterSpacing: '-0.025em' }}>
                        Governance & Analytics
                    </h1>
                    <p style={{ color: '#64748b', marginTop: '0.5rem', fontSize: '1.1rem' }}>
                        Overseeing <span style={{ color: '#3b82f6', fontWeight: 700 }}>{totalEnrollments}</span> active learning paths across <span style={{ color: '#3b82f6', fontWeight: 700 }}>{totalUsers}</span> team members.
                    </p>
                </div>
                <div style={{ display: 'flex', gap: '1rem' }}>
                    <button style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '0.5rem',
                        padding: '0.8rem 1.5rem',
                        backgroundColor: 'white',
                        border: '1px solid #e2e8f0',
                        borderRadius: '12px',
                        color: '#475569',
                        fontWeight: 600,
                        cursor: 'pointer',
                        transition: 'all 0.2s'
                    }} className="hover-lift">
                        <Download size={18} /> Export CSV
                    </button>
                    <button style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '0.5rem',
                        padding: '0.8rem 1.5rem',
                        backgroundColor: '#3b82f6',
                        border: 'none',
                        borderRadius: '12px',
                        color: 'white',
                        fontWeight: 700,
                        cursor: 'pointer',
                        boxShadow: '0 4px 12px rgba(59, 130, 246, 0.3)'
                    }}>
                        Bulk Manage
                    </button>
                </div>
            </header>

            {/* Stats Grid */}
            <div style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
                gap: '1.5rem',
                marginBottom: '3rem'
            }}>
                <StatCard
                    icon={<Users size={22} />}
                    label="Professionals"
                    value={totalUsers}
                    trend="+12%"
                    color="#3b82f6"
                    bg="#eff6ff"
                />
                <StatCard
                    icon={<BookOpen size={22} />}
                    label="Active Tracks"
                    value={totalEnrollments}
                    sub="Currently active"
                    color="#8b5cf6"
                    bg="#f5f3ff"
                />
                <StatCard
                    icon={<CheckCircle size={22} />}
                    label="Completion Rate"
                    value={`${completionRate}%`}
                    progress={completionRate}
                    color="#10b981"
                    bg="#ecfdf5"
                />
                <StatCard
                    icon={<TrendingUp size={22} />}
                    label="Avg. Proficiency"
                    value="84.2"
                    trend="+2.4"
                    color="#f59e0b"
                    bg="#fffbeb"
                />
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr', gap: '2rem' }}>
                {/* Main Content Area */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
                    {/* Enrollment Table */}
                    <div style={{ backgroundColor: 'white', borderRadius: '24px', padding: '2rem', border: '1px solid #e2e8f0', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
                            <h3 style={{ fontSize: '1.25rem', fontWeight: 800, color: '#1e293b', margin: 0 }}>Enrollment Tracker</h3>
                            <div style={{ display: 'flex', gap: '0.75rem' }}>
                                <div style={{ position: 'relative' }}>
                                    <Search size={18} style={{ position: 'absolute', left: '1rem', top: '50%', transform: 'translateY(-50%)', color: '#94a3b8' }} />
                                    <input
                                        type="text"
                                        placeholder="Search..."
                                        value={searchTerm}
                                        onChange={(e) => setSearchTerm(e.target.value)}
                                        style={{
                                            padding: '0.7rem 1rem 0.7rem 2.75rem',
                                            borderRadius: '12px',
                                            border: '1px solid #e2e8f0',
                                            fontSize: '0.9rem',
                                            width: '220px',
                                            backgroundColor: '#f8fafc'
                                        }}
                                    />
                                </div>
                                <select
                                    value={statusFilter}
                                    onChange={(e) => setStatusFilter(e.target.value)}
                                    style={{ padding: '0.7rem', borderRadius: '12px', border: '1px solid #e2e8f0', backgroundColor: 'white', fontSize: '0.9rem' }}
                                >
                                    <option value="all">Status: All</option>
                                    <option value="completed">Completed</option>
                                    <option value="in-progress">In Progress</option>
                                    <option value="scheduled">Scheduled</option>
                                </select>
                                <select
                                    value={deptFilter}
                                    onChange={(e) => setDeptFilter(e.target.value)}
                                    style={{ padding: '0.7rem', borderRadius: '12px', border: '1px solid #e2e8f0', backgroundColor: 'white', fontSize: '0.9rem' }}
                                >
                                    <option value="all">Dept: All</option>
                                    {Object.keys(deptStats).map(dept => (
                                        <option key={dept} value={dept}>{dept}</option>
                                    ))}
                                </select>
                            </div>
                        </div>

                        <div style={{ overflowX: 'auto' }}>
                            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                <thead>
                                    <tr style={{ borderBottom: '2px solid #f1f5f9' }}>
                                        <th style={{ textAlign: 'left', padding: '1rem', color: '#64748b', fontSize: '0.8rem', fontWeight: 700, textTransform: 'uppercase' }}>User</th>
                                        <th style={{ textAlign: 'left', padding: '1rem', color: '#64748b', fontSize: '0.8rem', fontWeight: 700, textTransform: 'uppercase' }}>Department</th>
                                        <th style={{ textAlign: 'left', padding: '1rem', color: '#64748b', fontSize: '0.8rem', fontWeight: 700, textTransform: 'uppercase' }}>Certification</th>
                                        <th style={{ textAlign: 'left', padding: '1rem', color: '#64748b', fontSize: '0.8rem', fontWeight: 700, textTransform: 'uppercase' }}>Status</th>
                                        <th style={{ textAlign: 'right', padding: '1rem', color: '#64748b', fontSize: '0.8rem', fontWeight: 700, textTransform: 'uppercase' }}>Actions</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filteredEnrollments.map((e) => (
                                        <tr key={e.id} style={{ borderBottom: '1px solid #f8fafc' }} className="table-row-hover">
                                            <td style={{ padding: '1rem' }}>
                                                <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                                                    <div style={{ width: '36px', height: '36px', backgroundColor: '#e2e8f0', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800, color: '#475569' }}>
                                                        {e.userName[0]}
                                                    </div>
                                                    <div>
                                                        <div style={{ fontWeight: 700, fontSize: '0.95rem' }}>{e.userName}</div>
                                                        <div style={{ fontSize: '0.8rem', color: '#94a3b8' }}>{e.userEmail}</div>
                                                    </div>
                                                </div>
                                            </td>
                                            <td style={{ padding: '1rem' }}>
                                                <span style={{ fontSize: '0.85rem', color: '#475569', backgroundColor: '#f1f5f9', padding: '0.25rem 0.6rem', borderRadius: '6px', fontWeight: 600 }}>
                                                    {e.department}
                                                </span>
                                            </td>
                                            <td style={{ padding: '1rem' }}>
                                                <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{e.courseName}</div>
                                                <div style={{ fontSize: '0.75rem', color: '#94a3b8' }}>{e.provider}</div>
                                            </td>
                                            <td style={{ padding: '1rem' }}>
                                                <StatusBadge status={e.status} />
                                            </td>
                                            <td style={{ padding: '1rem', textAlign: 'right' }}>
                                                <button
                                                    onClick={() => handleSendReminder(e.userEmail)}
                                                    style={{ border: 'none', background: 'none', color: '#3b82f6', cursor: 'pointer', padding: '0.5rem', borderRadius: '8px' }}
                                                    title="Send Reminder"
                                                >
                                                    <Mail size={18} />
                                                </button>
                                                <button style={{ border: 'none', background: 'none', color: '#94a3b8', cursor: 'pointer', padding: '0.5rem' }}>
                                                    <MoreVertical size={18} />
                                                </button>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                {/* Sidebar */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
                    {/* Activity Feed */}
                    <div style={{ backgroundColor: 'white', borderRadius: '24px', padding: '1.5rem', border: '1px solid #e2e8f0' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', marginBottom: '1.5rem' }}>
                            <Activity size={20} color="#3b82f6" />
                            <h3 style={{ fontSize: '1.1rem', fontWeight: 800, margin: 0 }}>Live Activity</h3>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                            {notifications.map(n => (
                                <div key={n.id} style={{ display: 'flex', gap: '1rem' }}>
                                    <div style={{
                                        width: '8px',
                                        height: '8px',
                                        borderRadius: '50%',
                                        marginTop: '6px',
                                        backgroundColor: n.type === 'success' ? '#10b981' : n.type === 'alert' ? '#ef4444' : '#3b82f6'
                                    }}></div>
                                    <div>
                                        <div style={{ fontSize: '0.875rem', fontWeight: 600 }}>{n.text}</div>
                                        <div style={{ fontSize: '0.75rem', color: '#94a3b8' }}>{n.time}</div>
                                    </div>
                                </div>
                            ))}
                        </div>
                    </div>

                    {/* Popular Courses */}
                    <div style={{ backgroundColor: 'white', borderRadius: '24px', padding: '1.5rem', border: '1px solid #e2e8f0' }}>
                        <h3 style={{ fontSize: '1.1rem', fontWeight: 800, marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                            <TrendingUp size={18} color="#8b5cf6" /> Popular Courses
                        </h3>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                            {courseStats.slice(0, 3).map((stat, idx) => (
                                <div key={idx}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.85rem', marginBottom: '0.5rem' }}>
                                        <span style={{ fontWeight: 600 }}>{stat.name}</span>
                                        <span style={{ color: '#64748b' }}>{stat.count}</span>
                                    </div>
                                    <div style={{ height: '4px', backgroundColor: '#f1f5f9', borderRadius: '2px' }}>
                                        <div style={{ width: `${(stat.count / totalEnrollments) * 100}%`, height: '100%', backgroundColor: '#8b5cf6' }}></div>
                                    </div>
                                </div>
                            ))}
                        </div>
                    </div>

                    {/* Quick Analytics */}
                    <div style={{ backgroundColor: '#1e293b', borderRadius: '24px', padding: '1.5rem', color: 'white' }}>
                        <h3 style={{ fontSize: '1.1rem', fontWeight: 700, marginBottom: '1.25rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                            <AlertCircle size={18} color="#fbbf24" /> Insight
                        </h3>
                        <p style={{ fontSize: '0.85rem', color: '#94a3b8', lineHeight: '1.6' }}>
                            IT Department is leading with <span style={{ color: 'white', fontWeight: 700 }}>92%</span> completion rate this quarter.
                            Security certifications have increased by <span style={{ color: '#10b981', fontWeight: 700 }}>14%</span>.
                        </p>
                        <div style={{ marginTop: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.8rem' }}>
                                <span>Platform Health</span>
                                <span style={{ color: '#10b981' }}>Operational</span>
                            </div>
                            <div style={{ height: '4px', backgroundColor: '#334155', borderRadius: '2px' }}>
                                <div style={{ width: '98%', height: '100%', backgroundColor: '#10b981' }}></div>
                            </div>
                        </div>
                    </div>

                    {/* Resources */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                        <ResourceLink icon={<BookOpen size={16} />} text="Learning Center" onClick={() => navigate('/learning-center')} />
                        <ResourceLink icon={<TrendingUp size={16} />} text="Course Catalog" onClick={() => navigate('/learning-center')} />
                        <ResourceLink icon={<Users size={16} />} text="User Management" onClick={() => navigate('/admin')} />
                    </div>
                </div>
            </div>
        </div>
    );
}

function StatCard({ icon, label, value, trend, sub, progress, color, bg }: any) {
    return (
        <div style={{ backgroundColor: 'white', padding: '1.5rem', borderRadius: '20px', border: '1px solid #e2e8f0', boxShadow: '0 1px 2px rgba(0,0,0,0.05)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '1rem' }}>
                <div style={{ padding: '0.6rem', backgroundColor: bg, borderRadius: '12px', color: color }}>
                    {icon}
                </div>
                {trend && <span style={{ color: '#10b981', fontSize: '0.8rem', fontWeight: 700, backgroundColor: '#f0fdf4', padding: '2px 8px', borderRadius: '100px' }}>{trend}</span>}
            </div>
            <div style={{ color: '#64748b', fontSize: '0.875rem', fontWeight: 600, marginBottom: '0.25rem' }}>{label}</div>
            <div style={{ fontSize: '1.75rem', fontWeight: 800 }}>{value}</div>
            {sub && <div style={{ fontSize: '0.75rem', color: '#94a3b8', marginTop: '0.5rem' }}>{sub}</div>}
            {progress !== undefined && (
                <div style={{ height: '6px', backgroundColor: '#f1f5f9', borderRadius: '10px', marginTop: '1rem', overflow: 'hidden' }}>
                    <div style={{ width: `${progress}%`, height: '100%', backgroundColor: color }}></div>
                </div>
            )}
        </div>
    );
}

function StatusBadge({ status }: { status: string }) {
    const styles: any = {
        'completed': { bg: '#dcfce7', text: '#166534', icon: <Check size={12} /> },
        'in-progress': { bg: '#eff6ff', text: '#1e40af', icon: <Activity size={12} /> },
        'scheduled': { bg: '#fef9c3', text: '#854d0e', icon: <Clock size={12} /> },
        'failed': { bg: '#fee2e2', text: '#991b1b', icon: <AlertCircle size={12} /> }
    };
    const s = styles[status] || styles['scheduled'];
    return (
        <div style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '4px',
            padding: '4px 10px',
            borderRadius: '100px',
            fontSize: '0.75rem',
            fontWeight: 700,
            backgroundColor: s.bg,
            color: s.text,
            textTransform: 'uppercase'
        }}>
            {s.icon} {status.replace('-', ' ')}
        </div>
    );
}

function ResourceLink({ icon, text, onClick }: { icon: any, text: string, onClick: () => void }) {
    return (
        <button
            type="button"
            onClick={onClick}
            style={{
            display: 'flex',
            alignItems: 'center',
            gap: '0.75rem',
            padding: '1rem',
            backgroundColor: 'white',
            borderRadius: '12px',
            color: '#475569',
            fontWeight: 600,
            fontSize: '0.9rem',
            border: '1px solid #e2e8f0',
            transition: 'all 0.2s',
            cursor: 'pointer',
            width: '100%',
            textAlign: 'left'
        }} className="resource-link-hover">
            {icon} {text}
            <ChevronRight size={14} style={{ marginLeft: 'auto', opacity: 0.5 }} />
        </button>
    );
}
