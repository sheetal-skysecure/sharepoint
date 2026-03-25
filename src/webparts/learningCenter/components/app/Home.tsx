import * as React from 'react';
import { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { ArrowRight, BookOpen, Users, BookMarked, CalendarDays, MessageCircle, Download, X, HelpCircle, Trophy, GraduationCap, FileCheck, TrendingUp, ClipboardCheck } from 'lucide-react';

export default function Home() {
    const [showFAQ, setShowFAQ] = useState(false);
    const navigate = useNavigate();

    const handleResourceClick = (title: string) => {
        switch (title) {
            case 'Help center':
                window.open('https://learn.microsoft.com/credentials/certifications/register-schedule-exam', '_blank');
                break;
            case 'Study guides':
                window.open('https://learn.microsoft.com/docs/', '_blank');
                break;
            case 'Course FAQs':
                setShowFAQ(true);
                break;
            case 'Course calendar':
                window.open('https://outlook.office.com/calendar', '_blank');
                break;
            case 'Textbooks':
                window.open('https://learn.microsoft.com/en-us/training/browse/', '_blank');
                break;
            case 'Message the instructor':
                window.location.href = 'mailto:learning@skysecuretech.com?subject=Learning Center Inquiry';
                break;
            default:
                break;
        }
    };

    const faqs = [
        { q: "How do I schedule an exam?", a: "Go to the Certification portal, search for your provider (e.g., Microsoft), and select 'Schedule' on the course card." },
        { q: "What is 'Smart Verification'?", a: "It's an automated check that ensures your uploaded certificate matches the exam code before marking it as complete." },
        { q: "Can I cancel a scheduled course?", a: "Yes, currently courses stay in your scheduled list, but you can always pick a different target date." }
    ];

    return (
        <div className="animate-fade-in" style={{ paddingBottom: '5rem', backgroundColor: 'var(--bg-color)', minHeight: '100vh' }}>

            {/* Hero Section */}
            <section style={{
                background: 'var(--gradient-primary)',
                color: 'white',
                position: 'relative',
                overflow: 'hidden',
                padding: '5rem 0 8rem'
            }}>
                {/* Decorative Elements */}
                <div style={{ position: 'absolute', top: '-100px', right: '-100px', width: '400px', height: '400px', borderRadius: '50%', background: 'rgba(255,255,255,0.05)' }} />
                <div style={{ position: 'absolute', bottom: '-50px', left: '10%', width: '200px', height: '200px', borderRadius: '50%', background: 'rgba(255,255,255,0.03)' }} />

                <div className="container" style={{ position: 'relative', zIndex: 2 }}>
                    <div style={{ maxWidth: '700px' }}>
                        <div style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '0.5rem',
                            backgroundColor: 'rgba(255,255,255,0.15)',
                            backdropFilter: 'blur(10px)',
                            padding: '0.5rem 1.25rem',
                            borderRadius: '100px',
                            fontSize: '0.875rem',
                            fontWeight: 600,
                            marginBottom: '2rem',
                            border: '1px solid rgba(255,255,255,0.2)'
                        }}>
                            <Trophy size={16} />
                            Your Professional Journey Starts Here
                        </div>

                        <h1 style={{
                            fontSize: '3.5rem',
                            fontWeight: 800,
                            marginBottom: '1.5rem',
                            lineHeight: '1.1',
                            letterSpacing: '-0.02em'
                        }}>
                            Master New Skills with <br />
                            <span style={{ color: '#bae6fd' }}>Skysecure Learning</span>
                        </h1>

                        <p style={{
                            fontSize: '1.25rem',
                            opacity: 0.9,
                            marginBottom: '3rem',
                            lineHeight: '1.6',
                            maxWidth: '600px'
                        }}>
                            Empowering you with a streamlined path to certification. Access curated courses,
                            smart verification, and expert resources all in one professional portal.
                        </p>

                        <div style={{ display: 'flex', gap: '1rem' }}>
                            <Link to="/learning-center" className="btn" style={{
                                backgroundColor: 'white',
                                color: 'var(--primary)',
                                padding: '1rem 2rem',
                                borderRadius: '100px',
                                fontSize: '1rem',
                                fontWeight: 700,
                                boxShadow: '0 10px 15px -3px rgba(0,0,0,0.1)'
                            }}>
                                View Certifications
                                <ArrowRight size={20} />
                            </Link>
                            <button onClick={() => setShowFAQ(true)} className="btn" style={{
                                backgroundColor: 'rgba(255,255,255,0.1)',
                                color: 'white',
                                border: '1px solid rgba(255,255,255,0.3)',
                                padding: '1rem 2rem',
                                borderRadius: '100px',
                                fontSize: '1rem'
                            }}>
                                How it works
                            </button>
                        </div>
                    </div>
                </div>
            </section>

            {/* Feature Cards Container */}
            <section className="container" style={{ marginTop: '-4rem', marginBottom: '4rem', zIndex: 10, position: 'relative' }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '2rem' }}>

                    <div className="glass-card" style={{ padding: '2rem', borderRadius: 'var(--border-radius-lg)', display: 'flex', gap: '1.5rem', alignItems: 'flex-start' }}>
                        <div style={{ backgroundColor: 'var(--primary-light)', color: 'var(--primary)', padding: '1rem', borderRadius: 'var(--border-radius-md)' }}>
                            <GraduationCap size={32} />
                        </div>
                        <div>
                            <h3 style={{ fontSize: '1.25rem', fontWeight: 700, marginBottom: '0.5rem' }}>Curated Tracks</h3>
                            <p style={{ color: 'var(--text-muted)', fontSize: '0.9375rem', lineHeight: '1.5' }}>
                                Expertly chosen certification paths from Microsoft, Google, and AWS.
                            </p>
                        </div>
                    </div>

                    <div className="glass-card" style={{ padding: '2rem', borderRadius: 'var(--border-radius-lg)', display: 'flex', gap: '1.5rem', alignItems: 'flex-start' }}>
                        <div style={{ backgroundColor: '#f0fdf4', color: '#16a34a', padding: '1rem', borderRadius: 'var(--border-radius-md)' }}>
                            <FileCheck size={32} />
                        </div>
                        <div>
                            <h3 style={{ fontSize: '1.25rem', fontWeight: 700, marginBottom: '0.5rem' }}>Smart Verification</h3>
                            <p style={{ color: 'var(--text-muted)', fontSize: '0.9375rem', lineHeight: '1.5' }}>
                                Automatic exam code detection to verify your accomplishments instantly.
                            </p>
                        </div>
                    </div>

                    <div className="glass-card" style={{ padding: '2rem', borderRadius: 'var(--border-radius-lg)', display: 'flex', gap: '1.5rem', alignItems: 'flex-start' }}>
                        <div style={{ backgroundColor: '#fff7ed', color: '#ea580c', padding: '1rem', borderRadius: 'var(--border-radius-md)' }}>
                            <CalendarDays size={32} />
                        </div>
                        <div>
                            <h3 style={{ fontSize: '1.25rem', fontWeight: 700, marginBottom: '0.5rem' }}>Goal Tracking</h3>
                            <p style={{ color: 'var(--text-muted)', fontSize: '0.9375rem', lineHeight: '1.5' }}>
                                Set target dates and keep your career growth organized and on schedule.
                            </p>
                        </div>
                    </div>

                </div>
            </section>

            {/* Resources Area */}
            <section className="container">
                <div style={{ textAlign: 'center', marginBottom: '3rem' }}>
                    <h2 style={{ fontSize: '2.25rem', fontWeight: 800, color: 'var(--text-main)', marginBottom: '1rem' }}>Learning Dashboard</h2>
                    <p style={{ color: 'var(--text-muted)', fontSize: '1.125rem' }}>Quick access to secondary tools and support documentation.</p>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '1.5rem' }}>
                    <PremiumResourceBox
                        title="Help center"
                        desc="Contact support for portal questions"
                        icon={<Users size={24} />}
                        onClick={() => handleResourceClick('Help center')}
                        color="#3b82f6"
                    />
                    <PremiumResourceBox
                        title="Study guides"
                        desc="Official Microsoft learning paths"
                        icon={<BookOpen size={24} />}
                        onClick={() => handleResourceClick('Study guides')}
                        color="#8b5cf6"
                    />
                    <PremiumResourceBox
                        title="Course FAQs"
                        desc="Common questions about the curriculum"
                        icon={<MessageCircle size={24} />}
                        onClick={() => handleResourceClick('Course FAQs')}
                        color="#ec4899"
                    />
                    <PremiumResourceBox
                        title="Course calendar"
                        desc="Manage your exam schedule"
                        icon={<CalendarDays size={24} />}
                        onClick={() => handleResourceClick('Course calendar')}
                        color="#f59e0b"
                    />
                    <PremiumResourceBox
                        title="Practice Assessments"
                        desc="Test your knowledge before the exam"
                        icon={<ClipboardCheck size={24} />}
                        onClick={() => navigate('/assessments')}
                        color="#10b981"
                    />
                    <PremiumResourceBox
                        title="Curriculum"
                        desc="Course syllabus and textbooks"
                        icon={<BookMarked size={24} />}
                        onClick={() => handleResourceClick('Textbooks')}
                        color="#6366f1"
                    />
                    <PremiumResourceBox
                        title="Contact Instructor"
                        desc="Get 1-on-1 expert guidance"
                        icon={<Download size={24} />}
                        onClick={() => handleResourceClick('Message the instructor')}
                        color="#06b6d4"
                    />
                    <PremiumResourceBox
                        title="Admin Governance"
                        desc="Analytics and enrollment monitoring"
                        icon={<TrendingUp size={24} />}
                        onClick={() => navigate('/admin')}
                        color="#4f46e5"
                    />
                </div>
            </section>

            {/* FAQ Modal */}
            {showFAQ && (
                <div className="modal-overlay" onClick={() => setShowFAQ(false)}>
                    <div className="modal-content" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '600px', borderRadius: '24px' }}>
                        <div className="modal-header" style={{ padding: '2rem' }}>
                            <h3 style={{ fontSize: '1.5rem', fontWeight: 800, margin: 0, display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                                <HelpCircle size={28} color="var(--primary)" />
                                Knowledge Base
                            </h3>
                            <button onClick={() => setShowFAQ(false)} style={{ backgroundColor: '#f1f5f9', padding: '0.5rem', borderRadius: '50%' }}>
                                <X size={20} />
                            </button>
                        </div>
                        <div className="modal-body" style={{ padding: '2rem' }}>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
                                {faqs.map((faq, i) => (
                                    <div key={i} style={{ backgroundColor: '#f8fafc', padding: '1.5rem', borderRadius: '16px' }}>
                                        <div style={{ fontWeight: 800, color: 'var(--text-main)', marginBottom: '1rem', fontSize: '1.125rem', display: 'flex', gap: '0.75rem' }}>
                                            <span style={{ color: 'var(--primary)' }}>Q.</span> {faq.q}
                                        </div>
                                        <div style={{ color: 'var(--text-muted)', fontSize: '1rem', lineHeight: '1.6', borderLeft: '3px solid #e2e8f0', paddingLeft: '1rem' }}>{faq.a}</div>
                                    </div>
                                ))}
                            </div>
                        </div>
                        <div className="modal-footer" style={{ padding: '1.5rem 2rem' }}>
                            <button className="btn btn-primary" style={{ padding: '0.75rem 2rem', borderRadius: '12px' }} onClick={() => setShowFAQ(false)}>Understood</button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}

function PremiumResourceBox({ title, desc, icon, onClick, color }: { title: string, desc: string, icon: any, onClick: () => void, color: string }) {
    return (
        <div
            onClick={onClick}
            style={{
                backgroundColor: 'white',
                padding: '1.5rem',
                borderRadius: 'var(--border-radius-lg)',
                display: 'flex',
                alignItems: 'center',
                gap: '1.25rem',
                cursor: 'pointer',
                boxShadow: 'var(--shadow-sm)',
                border: '1px solid var(--border-color)',
                transition: 'var(--transition)'
            }}
            className="hover:translate-y-[-4px] hover:shadow-lg"
        >
            <div style={{
                backgroundColor: `${color}15`,
                color: color,
                padding: '0.875rem',
                borderRadius: '12px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
            }}>
                {icon}
            </div>
            <div>
                <h4 style={{ fontWeight: 800, fontSize: '1.125rem', color: 'var(--text-main)', marginBottom: '0.25rem' }}>{title}</h4>
                <p style={{ color: 'var(--text-muted)', fontSize: '0.875rem' }}>{desc}</p>
            </div>
        </div>
    );
}
