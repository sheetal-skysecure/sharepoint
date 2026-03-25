import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { ClipboardCheck, Search, Filter, CheckCircle2, AlertCircle, Clock, Calendar } from 'lucide-react';
import { microsoftCertifications } from './data';
import { SharePointService, type IAssessmentAssignmentRecord } from '../../services/SharePointService';

export default function Assessments(props: { userDisplayName?: string, userEmail?: string, context?: any }) {
    const [searchTerm, setSearchTerm] = useState('');
    const [filterStatus, setFilterStatus] = useState('all');

    const [recordedResults, setRecordedResults] = useState<any>({});
    const [assignmentState, setAssignmentState] = useState<{
        loading: boolean;
        items: IAssessmentAssignmentRecord[];
        error: string | null;
    }>({
        loading: true,
        items: [],
        error: null
    });
    const currentUserEmail = (props.userEmail || props.context?.pageContext?.user?.email || '').toString().trim().toLowerCase();
    const loadedAssignmentsEmailRef = useRef<string>('');

    useEffect(() => {
        const syncRecordedResults = () => {
            const savedResults = localStorage.getItem('lmsAssessmentResults');
            if (!savedResults) {
                setRecordedResults({});
                return;
            }

            try {
                setRecordedResults(JSON.parse(savedResults) || {});
            } catch (error) {
                console.error('Failed to parse local assessment results', error);
                setRecordedResults({});
            }
        };

        syncRecordedResults();
        window.addEventListener('storage', syncRecordedResults);
        return () => {
            window.removeEventListener('storage', syncRecordedResults);
        };
    }, []);

    useEffect(() => {
        let isCancelled = false;

        const loadAssignments = async (): Promise<void> => {
            if (!currentUserEmail) {
                return;
            }

            try {
                const items = await SharePointService.getAssessmentAssignmentsForUser(currentUserEmail);
                if (isCancelled) {
                    return;
                }

                setAssignmentState((prev) => {
                    const previousJson = JSON.stringify(prev.items);
                    const nextJson = JSON.stringify(items);

                    if (!prev.loading && !prev.error && previousJson === nextJson) {
                        return prev;
                    }

                    return {
                        loading: false,
                        items,
                        error: null
                    };
                });
            } catch (error) {
                if (isCancelled) {
                    return;
                }

                const message = error instanceof Error ? error.message : 'Failed to load assigned assessments from SharePoint.';
                setAssignmentState((prev) => ({
                    loading: false,
                    items: prev.items,
                    error: message
                }));
            }
        };

        if (!currentUserEmail) {
            loadedAssignmentsEmailRef.current = '';
            setAssignmentState({
                loading: false,
                items: [],
                error: null
            });
            return;
        }

        if (loadedAssignmentsEmailRef.current === currentUserEmail) {
            return;
        }

        loadedAssignmentsEmailRef.current = currentUserEmail;
        void loadAssignments();

        return () => {
            isCancelled = true;
        };
    }, [currentUserEmail]);

    const allAssessments = useMemo(() => {
        const assessments: any[] = [];

        assignmentState.items.forEach((assignment) => {
            const payload = assignment.assessmentPayload;
            const lookupCode = (payload?.certCode || payload?.assessmentName || assignment.assessmentName || '').toString().trim();
            let detail: any = payload ? {
                id: payload.id || assignment.id,
                title: payload.title || assignment.title,
                certName: payload.assessmentName || assignment.assessmentName,
                certCode: payload.certCode || assignment.assessmentName,
                provider: payload.provider || 'Internal',
                duration: payload.duration || '20 Mins',
                questions: payload.questions || payload.questionsArr?.length || 10,
                questionsArr: payload.questionsArr || null,
                threshold: payload.threshold ?? 70
            } : null;

            microsoftCertifications.forEach((category) => {
                category.certs?.forEach((cert) => {
                    if (cert.code === lookupCode || cert.name === lookupCode || cert.assessment?.title === assignment.title) {
                        detail = {
                            ...(detail || {}),
                            title: detail?.title || cert.assessment?.title || assignment.title || `${cert.name} Practice Assessment`,
                            id: detail?.id || assignment.id,
                            certName: detail?.certName || cert.name,
                            certCode: detail?.certCode || cert.code,
                            provider: detail?.provider || 'Microsoft',
                            duration: detail?.duration || '20 Mins',
                            questions: detail?.questions || cert.assessment?.questions || 10,
                            questionsArr: detail?.questionsArr || (cert.assessment as any)?.questionsArr || null,
                            threshold: detail?.threshold ?? 70
                        };
                    }
                });
            });

            if (!detail) {
                detail = {
                    id: assignment.id,
                    title: assignment.title || `${assignment.assessmentName} Practice Assessment`,
                    certName: assignment.assessmentName,
                    certCode: assignment.assessmentName,
                    provider: 'Internal',
                    duration: '20 Mins',
                    questions: 10,
                    questionsArr: null,
                    threshold: 70
                };
            }

            const lookupKey = (assignment.id || detail.id || detail.certCode || assignment.assessmentName || 'unknown').toString();
            const result = recordedResults?.[lookupKey] || recordedResults?.[detail.certCode] || recordedResults?.[assignment.assessmentName];

            assessments.push({
                ...detail,
                id: assignment.id,
                title: assignment.title || detail.title,
                certName: detail.certName || assignment.assessmentName,
                certCode: detail.certCode || assignment.assessmentName,
                status: result ? 'completed' : 'pending',
                score: result ? result.score : null,
                availabilityStatus: 'available',
                orderIndex: assignment.orderIndex,
                assignedGroup: assignment.assignedGroup
            });
        });

        return assessments;
    }, [assignmentState.items, recordedResults]);
    const filtered = allAssessments.filter(a => {
        const titleMatch = (a.title || '').toLowerCase().includes(searchTerm.toLowerCase());
        const codeMatch = (a.certCode || '').toLowerCase().includes(searchTerm.toLowerCase());
        
        // Handle 'assigned' filter - in this context, all items in allAssessments 
        // are technically assigned to the user, but we can filter for 'pending' specifically 
        // if that's what's intended, or just treat 'assigned' as showing all targeted ones.
        const matchesStatus = filterStatus === 'all' || 
                             (filterStatus === 'assigned' && a.status === 'pending') ||
                             a.status === filterStatus;

        return (titleMatch || codeMatch) && matchesStatus;
    });

    const mockQuestions = [
        { id: 1, q: "Which of the following describes the most efficient method to manage role-based access control for multiple resources in the environment?", options: ["Apply individual permissions to each user account directly on every resource", "Use security groups and assign roles to the groups for centralized management", "Disable all default access and use manual override for every session", "Implement a script that rotates passwords daily across all systems"], correct: 1 },
        { id: 2, q: "What is the primary benefit of implementing a Zero Trust security model?", options: ["It eliminates the need for any firewalls", "It assumes every request is a potential breach and requires verification", "It allows all internal users to have administrative privileges", "It simplifies the network by removing encryption requirements"], correct: 1 },
        { id: 3, q: "In a shared responsibility model, who is typically responsible for the physical security of the data center in a cloud environment?", options: ["The customer", "The cloud service provider", "Both the customer and the provider", "The government agencies"], correct: 1 },
        { id: 4, q: "Which protocol is primarily used for secure web browsing communication?", options: ["HTTP", "FTP", "HTTPS", "SMTP"], correct: 2 },
        { id: 5, q: "What does MFA stand for in the context of security?", options: ["Multi-Function Access", "Multi-Factor Authentication", "Main Frame Authorization", "Modern File Access"], correct: 1 },
        { id: 6, q: "Which Microsoft service is used for identity and access management?", options: ["Microsoft Sentinel", "Microsoft Entra ID", "Microsoft Defender", "Microsoft Purview"], correct: 1 },
        { id: 7, q: "What is the primary purpose of a Firewall?", options: ["To speed up internet connection", "To filter incoming and outgoing network traffic", "To store backup data", "To manage user passwords"], correct: 1 },
        { id: 8, q: "Which type of malware is designed to lock files and demand payment?", options: ["Spyware", "Ransomware", "Trojan", "Adware"], correct: 1 },
        { id: 9, q: "What is 'Phishing'?", options: ["A way to catch fish in the office", "A social engineering attack to steal credentials", "A method of network load balancing", "A hardware component in servers"], correct: 1 },
        { id: 10, q: "What is the purpose of Data Encryption?", options: ["To compress files for storage", "To protect the confidentiality of data", "To increase data processing speed", "To delete redundant data automatically"], correct: 1 }
    ];

    const [taking, setTaking] = useState<any>(null);
    const [currentStep, setCurrentStep] = useState(0);
    const [selectedAnswer, setSelectedAnswer] = useState<number | null>(null);
    const [score, setScore] = useState(0);
    const [timeLeft, setTimeLeft] = useState(1200); // 20 minutes in seconds

    useEffect(() => {
        let timer: any;
        if (taking) {
            timer = setInterval(() => {
                setTimeLeft(prev => {
                    if (prev <= 0) {
                        clearInterval(timer);
                        return 0;
                    }
                    return prev - 1;
                });
            }, 1000);
        } else {
            setTimeLeft(1200);
        }
        return () => clearInterval(timer);
    }, [taking]);

    const formatTime = (seconds: number) => {
        const mins = Math.floor(seconds / 60);
        const secs = (seconds % 60).toString();
        return `${mins}:${secs.length === 1 ? '0' + secs : secs}`;
    };

    const activeUserName = props.userDisplayName || 'Unknown Learner';

    const questionsSource = useMemo(() => {
        if (!taking) return [];
        const raw = (taking.questionsArr && taking.questionsArr.length > 0) ? taking.questionsArr : mockQuestions;
        return [...raw].sort(() => 0.5 - Math.random());
    }, [taking?.id]);

    if (taking) {

        const totalQuestions = questionsSource.length;
        const currentQuestion: any = questionsSource[currentStep % totalQuestions];

        return (
            <div className="container animate-fade-in" style={{ padding: '3rem 2rem' }}>
                <header style={{ marginBottom: '3.5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                    <div style={{ flex: 1 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', color: 'var(--primary)', fontWeight: 800, textTransform: 'uppercase', fontSize: '0.85rem' }}>
                            <div style={{ width: '8px', height: '8px', borderRadius: '50%', background: 'var(--primary)', animation: 'pulse 2s infinite' }} />
                            LIVE ASSESSMENT IN PROGRESS
                        </div>
                        <h1 style={{ fontSize: '2.5rem', fontWeight: '800', margin: '0.5rem 0 0.25rem 0', letterSpacing: '-0.02em', color: '#1e293b' }}>{taking.certCode} Practice</h1>
                        <p style={{ color: '#64748b', fontSize: '1.1rem', fontWeight: 600 }}>Question {currentStep + 1} of {totalQuestions}</p>
                    </div>

                    <div style={{
                        background: 'rgba(14, 165, 233, 0.1)',
                        padding: '1.25rem 2rem',
                        borderRadius: '24px',
                        border: '2px solid var(--primary)',
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        minWidth: '220px',
                        boxShadow: '0 10px 15px -3px rgba(14, 165, 233, 0.1)'
                    }}>
                        <div style={{ fontSize: '0.85rem', fontWeight: 800, color: 'var(--primary)', marginBottom: '0.5rem', letterSpacing: '0.05em' }}>REMAINING TIME</div>
                        <div style={{ fontSize: '2.25rem', fontWeight: 950, color: timeLeft < 60 ? '#ef4444' : '#1e293b', display: 'flex', alignItems: 'center', gap: '12px', lineHeight: 1 }}>
                            <Clock size={32} style={{ color: timeLeft < 60 ? '#ef4444' : 'var(--primary)' }} /> {formatTime(timeLeft)}
                        </div>
                    </div>
                </header>

                <div className="assessment-card" style={{ background: 'white', padding: '3.5rem', borderRadius: '32px', border: '1px solid #e2e8f0', boxShadow: '0 25px 50px -12px rgba(0,0,0,0.08)' }}>
                    {/* ... question content remains same ... */}
                    <h3 style={{ fontSize: '1.6rem', fontWeight: 800, color: '#1e293b', marginBottom: '3rem', lineHeight: 1.4 }}>
                        {currentQuestion.q}
                    </h3>

                    <div style={{ display: 'grid', gap: '1.25rem' }}>
                        {currentQuestion.options.map((opt: string, i: number) => (
                            <button
                                key={i}
                                onClick={() => setSelectedAnswer(i)}
                                className="hover-lift"
                                style={{
                                    textAlign: 'left',
                                    padding: '1.75rem 2.5rem',
                                    borderRadius: '20px',
                                    border: '2px solid',
                                    background: selectedAnswer === i ? 'rgba(14, 165, 233, 0.05)' : 'white',
                                    borderColor: selectedAnswer === i ? 'var(--primary)' : '#f1f5f9',
                                    cursor: 'pointer',
                                    fontSize: '1.15rem',
                                    fontWeight: 700,
                                    color: selectedAnswer === i ? '#0f172a' : '#475569',
                                    transition: 'all 0.25s cubic-bezier(0.4, 0, 0.2, 1)',
                                    outline: 'none',
                                    boxShadow: selectedAnswer === i ? '0 10px 15px -3px rgba(14, 165, 233, 0.1)' : 'none'
                                }}
                            >
                                <div style={{ display: 'flex', alignItems: 'center', gap: '1.25rem' }}>
                                    <div style={{
                                        width: '28px',
                                        height: '28px',
                                        borderRadius: '50%',
                                        border: '2px solid',
                                        borderColor: selectedAnswer === i ? 'var(--primary)' : '#cbd5e1',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        background: selectedAnswer === i ? 'var(--primary)' : 'transparent',
                                        transition: 'all 0.2s'
                                    }}>
                                        {selectedAnswer === i && <div style={{ width: '10px', height: '10px', borderRadius: '50%', background: 'white' }} />}
                                    </div>
                                    <span>{opt}</span>
                                </div>
                            </button>
                        ))}
                    </div>

                    <div style={{ marginTop: '5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <button className="btn-secondary" onClick={() => {
                            if (window.confirm("Are you sure you want to exit? Your progress will be lost.")) {
                                setTaking(null);
                                setCurrentStep(0);
                                setSelectedAnswer(null);
                            }
                        }} style={{ padding: '1.25rem 2.5rem', borderRadius: '20px', fontSize: '1.1rem', fontWeight: 800 }}>Cancel</button>
                        <button
                            className="btn-primary"
                            disabled={selectedAnswer === null}
                            onClick={() => {
                                let newScore = score;
                                if (selectedAnswer === currentQuestion.correct) {
                                    newScore += 1;
                                }
                                setScore(newScore);

                                if (currentStep < totalQuestions - 1) {
                                    setCurrentStep(currentStep + 1);
                                    setSelectedAnswer(null);
                                } else {
                                    // Current newScore already includes the current question result
                                    const finalPerc = Math.round((newScore / totalQuestions) * 100);

                                    // Record the result
                                    const assessmentId = taking.id || taking.certCode || 'unknown';
                                    const prevResult = recordedResults[assessmentId];
                                    const attempts = (prevResult?.attempts || 0) + 1;
                                    const newResults = {
                                        ...recordedResults,
                                        [assessmentId]: {
                                            score: finalPerc,
                                            date: new Date().toISOString(),
                                            attempts,
                                            title: taking.title || 'Internal Assessment',
                                            certCode: taking.certCode,
                                            user: activeUserName
                                        }
                                    };
                                    setRecordedResults(newResults);
                                    localStorage.setItem('lmsAssessmentResults', JSON.stringify(newResults));

                                    alert(`Assessment completed! Your final score: ${finalPerc}%. Results have been recorded for ${taking.certCode || 'this assessment'}.`);
                                    setTaking(null);
                                    setCurrentStep(0);
                                    setSelectedAnswer(null);
                                    setScore(0);
                                }
                            }}
                            style={{
                                padding: '1.25rem 3rem',
                                borderRadius: '20px',
                                fontSize: '1.15rem',
                                fontWeight: 800,
                                opacity: selectedAnswer === null ? 0.5 : 1,
                                cursor: selectedAnswer === null ? 'not-allowed' : 'pointer'
                            }}
                        >
                            {currentStep < totalQuestions - 1 ? 'Next Question' : 'Submit Assessment'} <ClipboardCheck size={22} />
                        </button>
                    </div>
                </div>
            </div>
        );
    }

    return (
        <div className="container animate-fade-in" style={{ padding: '3rem 2rem' }}>
            <header style={{ marginBottom: '4rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '1.25rem', marginBottom: '1.25rem' }}>
                    <div style={{ padding: '1rem', background: 'linear-gradient(135deg, var(--primary) 0%, #0ea5e9 100%)', borderRadius: '16px', color: 'white', boxShadow: '0 10px 15px -3px rgba(14, 165, 233, 0.4)' }}>
                        <ClipboardCheck size={36} />
                    </div>
                    <div>
                        <h1 style={{ fontSize: '2.8rem', fontWeight: '900', margin: 0, letterSpacing: '-0.03em', color: '#0f172a' }}>Assessment Hub</h1>
                        <p style={{ color: '#64748b', fontSize: '1.25rem', fontWeight: 500, marginTop: '0.25rem' }}>Sharpen your skills with professional certification prep.</p>
                    </div>
                </div>
            </header>

            {assignmentState.error && (
                <div style={{ marginBottom: '2rem', padding: '1rem 1.25rem', borderRadius: '16px', border: '1px solid #fecaca', background: '#fef2f2', color: '#b91c1c', fontWeight: 700 }}>
                    {assignmentState.error}
                </div>
            )}

            {assignmentState.loading && allAssessments.length === 0 && (
                <div style={{ marginBottom: '2rem', padding: '1rem 1.25rem', borderRadius: '16px', border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', fontWeight: 700 }}>
                    Loading your assigned assessments from SharePoint...
                </div>
            )}

            {Object.keys(recordedResults).length > 0 && (
                <div style={{ marginBottom: '5rem' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '2rem' }}>
                        <div style={{ width: '4px', height: '24px', background: 'var(--primary)', borderRadius: '4px' }} />
                        <h2 style={{ fontSize: '1.75rem', fontWeight: 900, color: '#1e293b', margin: 0, letterSpacing: '-0.01em' }}>Recent Performance</h2>
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '2rem' }}>
                        {allAssessments.filter(as => as.status === 'completed').map(as => {
                            const isPassed = (as.score || 0) >= (as.threshold || 70);
                            return (
                                <div
                                    key={`res_${as.id}`}
                                    className="hover-lift"
                                    style={{
                                        padding: '2rem',
                                        background: 'white',
                                        borderRadius: '24px',
                                        border: '1.5px solid #f1f5f9',
                                        display: 'flex',
                                        justifyContent: 'space-between',
                                        alignItems: 'center',
                                        boxShadow: '0 4px 6px -1px rgba(0,0,0,0.02), 0 10px 15px -3px rgba(0,0,0,0.03)',
                                        position: 'relative',
                                        overflow: 'hidden'
                                    }}
                                >
                                    <div style={{ position: 'absolute', left: 0, top: 0, bottom: 0, width: '6px', background: isPassed ? '#10b981' : '#ef4444' }} />
                                    <div>
                                        <div style={{ fontSize: '0.75rem', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: '0.5rem' }}>{(as.certCode || 'CERT')} RESULT</div>
                                        <div style={{ fontSize: '1.25rem', fontWeight: 900, color: '#1e293b', lineHeight: 1.2 }}>{(as.title || 'Assessment').split(':')[0]}</div>
                                        <div style={{ marginTop: '0.75rem', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                            <div style={{ width: '8px', height: '8px', borderRadius: '50%', background: isPassed ? '#10b981' : '#ef4444' }} />
                                            <span style={{ fontSize: '0.85rem', fontWeight: 800, color: isPassed ? '#059669' : '#dc2626' }}>{isPassed ? 'CERTIFIED' : 'RETAKE SUGGESTED'}</span>
                                        </div>
                                    </div>
                                    <div style={{ textAlign: 'right', background: isPassed ? '#ecfdf5' : '#fef2f2', padding: '1rem', borderRadius: '20px', minWidth: '90px' }}>
                                        <div style={{ fontSize: '2rem', fontWeight: 950, color: isPassed ? '#059669' : '#dc2626', lineHeight: 1 }}>{as.score}%</div>
                                        <div style={{ fontSize: '0.7rem', fontWeight: 900, color: isPassed ? '#059669' : '#dc2626', marginTop: '4px', letterSpacing: '0.05em' }}>SCORE</div>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>
            )}

            <div className="search-box-unified" style={{ marginBottom: '3rem', maxWidth: '100%', height: '60px', padding: '0 1.5rem', borderRadius: '16px' }}>
                <Search size={24} />
                <input
                    type="text"
                    placeholder="Search assessments, certifications..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    style={{ fontSize: '1.1rem' }}
                />
                <div style={{ width: '2px', height: '30px', background: '#e2e8f0', margin: '0 1rem' }} />
                <Filter size={20} style={{ color: '#94a3b8' }} />
                <select
                    style={{ border: 'none', background: 'transparent', fontWeight: 700, fontSize: '1rem', outline: 'none', padding: '0 0.5rem', cursor: 'pointer' }}
                    value={filterStatus}
                    onChange={(e) => setFilterStatus(e.target.value)}
                >
                    <option value="all">All Status</option>
                    <option value="assigned">Assigned to me</option>
                    <option value="completed">Completed</option>
                    <option value="pending">Pending</option>
                </select>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: '2rem' }}>
                {filtered.map((as, idx) => (
                    <div
                        key={as.id}
                        className="assessment-card"
                        style={{
                            background: 'white',
                            borderRadius: '24px',
                            padding: '2rem',
                            border: '1px solid #e2e8f0',
                            boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)',
                            display: 'flex',
                            flexDirection: 'column',
                            gap: '1.5rem',
                            position: 'relative',
                            overflow: 'hidden'
                        }}
                    >
                        {as.status === 'completed' && (
                            <div style={{ position: 'absolute', top: 0, right: 0, padding: '0.5rem 1rem', background: (as.score || 0) >= (as.threshold || 70) ? '#dcfce7' : '#fee2e2', color: (as.score || 0) >= (as.threshold || 70) ? '#15803d' : '#991b1b', fontSize: '0.75rem', fontWeight: 900, borderRadius: '0 0 0 12px' }}>
                                {(as.score || 0) >= (as.threshold || 70) ? 'PASSED' : 'FAILED'}
                            </div>
                        )}

                        <div>
                            <span style={{ fontSize: '0.75rem', fontWeight: 900, color: 'var(--primary)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>{as.certCode} | {as.provider}</span>
                            <h3 style={{ fontSize: '1.25rem', fontWeight: 800, color: '#1e293b', marginTop: '0.5rem', lineHeight: 1.3 }}>{as.title}</h3>
                            <p style={{ fontSize: '0.9rem', color: '#64748b', fontWeight: 600, marginTop: '0.25rem' }}>For {as.certName}</p>
                        </div>

                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#64748b' }}>
                                <Clock size={16} />
                                <span style={{ fontSize: '0.85rem', fontWeight: 700 }}>{as.duration}</span>
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#64748b' }}>
                                <ClipboardCheck size={16} />
                                <span style={{ fontSize: '0.85rem', fontWeight: 700 }}>{as.questions} Questions</span>
                            </div>
                        </div>

                        <div style={{ marginTop: '0.5rem', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                            <div style={{ fontSize: '0.75rem', fontWeight: 700, color: '#94a3b8' }}>
                                SharePoint Assignment:
                            </div>
                            <div style={{ fontSize: '0.85rem', fontWeight: 800, color: '#475569', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Calendar size={14} /> Order {as.orderIndex} | {as.assignedGroup}
                            </div>
                        </div>

                        <div style={{ marginTop: 'auto', paddingTop: '1.5rem', borderTop: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            {as.status === 'completed' ? (
                                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                    <div style={{ textAlign: 'center' }}>
                                        <div style={{ fontSize: '0.7rem', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase' }}>Your Score</div>
                                        <div style={{ fontSize: '1.25rem', fontWeight: 900, color: '#10b981' }}>{as.score}%</div>
                                    </div>
                                    <CheckCircle2 size={24} style={{ color: '#10b981' }} />
                                </div>
                            ) : as.availabilityStatus === 'upcoming' ? (
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#6366f1' }}>
                                    <Clock size={20} />
                                    <span style={{ fontSize: '0.85rem', fontWeight: 800 }}>Opens {as.startDate}</span>
                                </div>
                            ) : as.availabilityStatus === 'expired' ? (
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#ef4444' }}>
                                    <AlertCircle size={20} />
                                    <span style={{ fontSize: '0.85rem', fontWeight: 800 }}>Expired {as.endDate}</span>
                                </div>
                            ) : (
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#f59e0b' }}>
                                    <AlertCircle size={20} />
                                    <span style={{ fontSize: '0.85rem', fontWeight: 800 }}>Not Attempted</span>
                                </div>
                            )}

                            <button
                                className={as.status === 'completed' ? 'btn-secondary' : 'btn-primary'}
                                style={{ 
                                    padding: '0.75rem 1.25rem', 
                                    borderRadius: '14px', 
                                    fontSize: '0.9rem',
                                    opacity: as.availabilityStatus !== 'available' && as.status !== 'completed' ? 0.5 : 1,
                                    cursor: as.availabilityStatus !== 'available' && as.status !== 'completed' ? 'not-allowed' : 'pointer'
                                }}
                                onClick={() => {
                                    if (as.status !== 'completed' && as.availabilityStatus !== 'available') {
                                        alert(`This assessment is only available from ${as.startDate} to ${as.endDate}`);
                                        return;
                                    }
                                    setTaking(as);
                                }}
                                disabled={as.availabilityStatus !== 'available' && as.status !== 'completed'}
                            >
                                {as.status === 'completed' ? 'Retake Test' : 'Start Assessment'}
                            </button>
                        </div>
                    </div>
                ))}
            </div>

            {!assignmentState.loading && filtered.length === 0 && (
                <div style={{ marginTop: '2rem', padding: '2.5rem', borderRadius: '24px', border: '1px solid #e2e8f0', background: '#ffffff', textAlign: 'center', color: '#64748b', fontWeight: 700 }}>
                    No assessments are currently assigned to your account in the SharePoint list `Asseswment`.
                </div>
            )}
        </div>
    );
}
