import * as React from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { PlusCircle, ClipboardCheck, BookOpen } from 'lucide-react';
import { companies, getCertsByCompany } from './data';

export default function LearningCenterDashboard() {
    const [searchParams] = useSearchParams();
    const navigate = useNavigate();
    const searchTerm = (searchParams.get('q') || '').trim().toLowerCase();
    const providerCardStyle: React.CSSProperties = {
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '3rem',
        backgroundColor: 'var(--card-bg)',
        border: '1px solid var(--border-color)',
        borderRadius: 'var(--border-radius-lg)',
        boxShadow: 'var(--shadow-md)',
        transition: 'var(--transition)',
        textDecoration: 'none',
        color: 'var(--text-main)',
        gap: '1.5rem',
        width: '100%',
        cursor: 'pointer'
    };

    const handleProviderClick = React.useCallback((providerId: string) => {
        const normalizedProviderId = (providerId || '').toString().trim().toLowerCase();
        navigate(`/learning-center/${normalizedProviderId}`, {
            state: { provider: normalizedProviderId }
        });
    }, [navigate]);

    const filteredCompanies = companies.filter((company) => {
        if (!searchTerm) {
            return true;
        }

        if ((company.name || '').toLowerCase().includes(searchTerm)) {
            return true;
        }

        const providerSections = getCertsByCompany(company.id) || [];
        return providerSections.some((section: any) =>
            (section.category || '').toLowerCase().includes(searchTerm) ||
            (section.level || '').toLowerCase().includes(searchTerm) ||
            (section.certs || []).some((cert: any) =>
                (cert.name || '').toLowerCase().includes(searchTerm) ||
                (cert.code || '').toLowerCase().includes(searchTerm) ||
                (cert.roles || []).some((role: string) => (role || '').toLowerCase().includes(searchTerm))
            )
        );
    });

    return (
        <div className="container animate-fade-in" style={{ padding: '3rem 2rem' }}>
            <header style={{ marginBottom: '3.5rem' }}>
                <h1 style={{ fontSize: '2.5rem', fontWeight: '800', marginBottom: '0.75rem', letterSpacing: '-0.02em', color: '#1e293b' }}>Learning Center</h1>
                <p style={{ color: 'var(--text-muted)', fontSize: '1.15rem' }}>
                    {searchTerm
                        ? `Showing providers and certification paths matching "${searchParams.get('q')}".`
                        : 'Select a provider to view available certifications and schedule your exams.'}
                </p>
            </header>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '2rem' }}>
                {filteredCompanies.map((company, index) => (
                    <button
                        type="button"
                        key={company.id}
                        aria-label={`Open ${company.name} certifications`}
                        className="animate-fade-in"
                        style={{
                            ...providerCardStyle,
                            animationDelay: `${index * 100}ms`,
                        }}
                        onClick={() => handleProviderClick(company.id)}
                        onMouseOver={(e) => {
                            e.currentTarget.style.transform = 'translateY(-4px)';
                            e.currentTarget.style.boxShadow = 'var(--shadow-lg)';
                            e.currentTarget.style.borderColor = 'var(--primary)';
                        }}
                        onMouseOut={(e) => {
                            e.currentTarget.style.transform = 'translateY(0)';
                            e.currentTarget.style.boxShadow = 'var(--shadow-md)';
                            e.currentTarget.style.borderColor = 'var(--border-color)';
                        }}
                    >
                        <img
                            src={company.logoUrl}
                            alt={`${company.name} logo`}
                            style={{ width: '80px', height: '80px', objectFit: 'contain' }}
                        />
                        <h3 style={{ fontSize: '1.25rem', fontWeight: 500 }}>{company.name}</h3>
                    </button>
                ))}

                {searchTerm && filteredCompanies.length === 0 && (
                    <div
                        style={{
                            gridColumn: '1 / -1',
                            padding: '3rem',
                            backgroundColor: 'var(--card-bg)',
                            border: '1px dashed var(--border-color)',
                            borderRadius: 'var(--border-radius-lg)',
                            color: 'var(--text-muted)',
                            fontWeight: 700,
                            textAlign: 'center'
                        }}
                    >
                        No providers or certification paths matched "{searchParams.get('q')}".
                    </div>
                )}

                {/* Self Explore Card */}
                <Link
                    to="/explore"
                    className="animate-fade-in"
                    style={{
                        animationDelay: `${companies.length * 100}ms`,
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        justifyContent: 'center',
                        padding: '3rem',
                        backgroundColor: 'var(--card-bg)',
                        border: '1px solid var(--border-color)',
                        borderRadius: 'var(--border-radius-lg)',
                        boxShadow: 'var(--shadow-md)',
                        transition: 'var(--transition)',
                        textDecoration: 'none',
                        color: 'var(--text-main)',
                        gap: '1.5rem',
                        background: 'linear-gradient(to bottom right, #ffffff, #f8fafc)'
                    }}
                    onMouseOver={(e) => {
                        e.currentTarget.style.transform = 'translateY(-4px)';
                        e.currentTarget.style.boxShadow = 'var(--shadow-lg)';
                        e.currentTarget.style.borderColor = 'var(--primary)';
                    }}
                    onMouseOut={(e) => {
                        e.currentTarget.style.transform = 'translateY(0)';
                        e.currentTarget.style.boxShadow = 'var(--shadow-md)';
                        e.currentTarget.style.borderColor = 'var(--border-color)';
                    }}
                >
                    <div style={{
                        width: '80px',
                        height: '80px',
                        backgroundColor: '#eff6ff',
                        borderRadius: '20px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: 'var(--primary)',
                        border: '2px dashed var(--primary-light)'
                    }}>
                        <PlusCircle size={40} />
                    </div>
                    <h3 style={{ fontSize: '1.25rem', fontWeight: 700, color: 'var(--primary)' }}>Self Explore</h3>
                </Link>

                <Link
                    to="/assessments"
                    className="animate-fade-in"
                    style={{
                        animationDelay: `${(companies.length + 1) * 100}ms`,
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        justifyContent: 'center',
                        padding: '3rem',
                        backgroundColor: 'var(--card-bg)',
                        border: '1px solid var(--border-color)',
                        borderRadius: 'var(--border-radius-lg)',
                        boxShadow: 'var(--shadow-md)',
                        transition: 'var(--transition)',
                        textDecoration: 'none',
                        color: 'var(--text-main)',
                        gap: '1.5rem',
                        background: 'linear-gradient(to bottom right, #ffffff, #f0fdf4)'
                    }}
                    onMouseOver={(e) => {
                        e.currentTarget.style.transform = 'translateY(-4px)';
                        e.currentTarget.style.boxShadow = 'var(--shadow-lg)';
                        e.currentTarget.style.borderColor = '#10b981';
                    }}
                    onMouseOut={(e) => {
                        e.currentTarget.style.transform = 'translateY(0)';
                        e.currentTarget.style.boxShadow = 'var(--shadow-md)';
                        e.currentTarget.style.borderColor = 'var(--border-color)';
                    }}
                >
                    <div style={{
                        width: '80px',
                        height: '80px',
                        backgroundColor: '#f0fdf4',
                        borderRadius: '20px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: '#10b981',
                        border: '2px dashed #bbf7d0'
                    }}>
                        <ClipboardCheck size={40} />
                    </div>
                    <h3 style={{ fontSize: '1.25rem', fontWeight: 700, color: '#10b981' }}>Assessments</h3>
                </Link>

                <Link
                    to="/library"
                    className="animate-fade-in"
                    style={{
                        animationDelay: `${(companies.length + 2) * 100}ms`,
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        justifyContent: 'center',
                        padding: '3rem',
                        backgroundColor: 'var(--card-bg)',
                        border: '1px solid var(--border-color)',
                        borderRadius: 'var(--border-radius-lg)',
                        boxShadow: 'var(--shadow-md)',
                        transition: 'var(--transition)',
                        textDecoration: 'none',
                        color: 'var(--text-main)',
                        gap: '1.5rem',
                        background: 'linear-gradient(to bottom right, #ffffff, #faf5ff)'
                    }}
                    onMouseOver={(e) => {
                        e.currentTarget.style.transform = 'translateY(-4px)';
                        e.currentTarget.style.boxShadow = 'var(--shadow-lg)';
                        e.currentTarget.style.borderColor = '#c084fc';
                    }}
                    onMouseOut={(e) => {
                        e.currentTarget.style.transform = 'translateY(0)';
                        e.currentTarget.style.boxShadow = 'var(--shadow-md)';
                        e.currentTarget.style.borderColor = 'var(--border-color)';
                    }}
                >
                    <div style={{
                        width: '80px',
                        height: '80px',
                        backgroundColor: '#faf5ff',
                        borderRadius: '20px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: '#c084fc',
                        border: '2px dashed #e9d5ff'
                    }}>
                        <BookOpen size={40} />
                    </div>
                    <h3 style={{ fontSize: '1.25rem', fontWeight: 700, color: '#c084fc' }}>Content Library</h3>
                </Link>
            </div>
        </div>
    );
}
