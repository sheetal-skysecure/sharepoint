import * as React from 'react';
import { useState, useEffect } from 'react';
import { Trash2, Search, Award, ExternalLink, X, PlusCircle } from 'lucide-react';

export default function SelfExplore() {
    const [customCerts, setCustomCerts] = useState<any[]>([]);
    const [showForm, setShowForm] = useState(false);
    const [formData, setFormData] = useState({
        name: '',
        provider: '',
        code: '',
        description: '',
        url: ''
    });
    const [activeUser, setActiveUser] = useState<any>(null);

    useEffect(() => {
        const stored = localStorage.getItem('selfExploreCerts');
        if (stored) {
            try {
                setCustomCerts(JSON.parse(stored));
            } catch (e) {
                console.error(e);
            }
        }

        const userStr = localStorage.getItem('lmsActiveUser');
        if (userStr) {
            try {
                const user = JSON.parse(userStr);
                setActiveUser(user);

                // Patch existing records if they belong to this user session
                const existing = localStorage.getItem('selfExploreCerts');
                if (existing) {
                    const parsed = JSON.parse(existing);
                    let changed = false;
                    const patched = parsed.map((c: any) => {
                        if (!c.userName || c.userName === 'Unknown User') {
                            changed = true;
                            return { ...c, userName: user.name, email: user.email };
                        }
                        return c;
                    });
                    if (changed) {
                        setCustomCerts(patched);
                    }
                }
            } catch (e) { }
        }
    }, []);

    useEffect(() => {
        localStorage.setItem('selfExploreCerts', JSON.stringify(customCerts));
    }, [customCerts]);

    const handleAdd = (e: React.FormEvent) => {
        e.preventDefault();
        if (!formData.name || !formData.provider) return;

        const newCert = {
            ...formData,
            id: Date.now(),
            dateAdded: new Date().toLocaleDateString(),
            status: 'Self-Paced',
            userName: activeUser?.name || 'Unknown User',
            email: activeUser?.email || ''
        };

        setCustomCerts([...customCerts, newCert]);
        setFormData({ name: '', provider: '', code: '', description: '', url: '' });
        setShowForm(false);
    };

    const handleRemove = (id: number) => {
        if (window.confirm("Remove this certification?")) {
            setCustomCerts(customCerts.filter(c => c.id !== id));
        }
    };

    return (
        <div className="container" style={{ padding: '4rem 0', width: '100%' }}>
            <div style={{ marginBottom: '3rem', textAlign: 'center' }}>
                <h1 style={{ fontSize: '3rem', fontWeight: 900, color: '#111827', marginBottom: '1rem', letterSpacing: '-0.04em' }}>
                    Self <span style={{ color: 'var(--primary)' }}>Explore</span>
                </h1>
                <p style={{ color: '#64748b', fontSize: '1.1rem', maxWidth: '600px', margin: '0 auto' }}>
                    Track certifications you're pursuing independently. Add your own custom paths and milestones beyond our curated list.
                </p>
            </div>

            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '4rem' }}>
                <button
                    onClick={() => setShowForm(true)}
                    className="btn-primary"
                    style={{
                        padding: '1rem 2.5rem',
                        fontSize: '1.1rem',
                        borderRadius: '100px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '0.75rem',
                        boxShadow: '0 20px 25px -5px rgba(15, 98, 254, 0.2)'
                    }}
                >
                    <PlusCircle size={24} /> Add Personal Certification
                </button>
            </div>

            {showForm && (
                <div style={{
                    position: 'fixed',
                    inset: 0,
                    backgroundColor: 'rgba(0,0,0,0.5)',
                    backdropFilter: 'blur(8px)',
                    zIndex: 1000,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    padding: '1.5rem'
                }}>
                    <div className="fade-in" style={{
                        backgroundColor: 'white',
                        padding: '2.5rem',
                        borderRadius: '32px',
                        width: '100%',
                        maxWidth: '550px',
                        boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)'
                    }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
                            <h2 style={{ fontSize: '1.5rem', fontWeight: 800, margin: 0 }}>Add Custom Journey</h2>
                            <button onClick={() => setShowForm(false)} style={{ background: 'none', border: 'none', color: '#94a3b8', cursor: 'pointer' }}><X size={24} /></button>
                        </div>

                        <form onSubmit={handleAdd}>
                            <div style={{ display: 'grid', gap: '1.5rem' }}>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.5rem' }}>Certification Name</label>
                                    <input
                                        type="text"
                                        className="input-field"
                                        placeholder="e.g. Google Cloud Professional Architect"
                                        value={formData.name}
                                        onChange={e => setFormData({ ...formData, name: e.target.value })}
                                        required
                                    />
                                </div>
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
                                    <div>
                                        <label style={{ display: 'block', fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.5rem' }}>Provider</label>
                                        <input
                                            type="text"
                                            className="input-field"
                                            placeholder="e.g. Google Cloud"
                                            value={formData.provider}
                                            onChange={e => setFormData({ ...formData, provider: e.target.value })}
                                            required
                                        />
                                    </div>
                                    <div>
                                        <label style={{ display: 'block', fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.5rem' }}>Exam Code</label>
                                        <input
                                            type="text"
                                            className="input-field"
                                            placeholder="e.g. PCA-2024"
                                            value={formData.code}
                                            onChange={e => setFormData({ ...formData, code: e.target.value })}
                                        />
                                    </div>
                                </div>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.5rem' }}>Resource URL (Optional)</label>
                                    <input
                                        type="url"
                                        className="input-field"
                                        placeholder="https://learn.provider.com/..."
                                        value={formData.url}
                                        onChange={e => setFormData({ ...formData, url: e.target.value })}
                                    />
                                </div>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.5rem' }}>Key Skills / Interests</label>
                                    <textarea
                                        className="input-field"
                                        style={{ height: '100px', resize: 'none' }}
                                        placeholder="What motivated you to take this certification?"
                                        value={formData.description}
                                        onChange={e => setFormData({ ...formData, description: e.target.value })}
                                    />
                                </div>
                                <button type="submit" className="btn-primary" style={{ padding: '1.1rem', borderRadius: '14px', fontWeight: 800 }}>Create Personal Path</button>
                            </div>
                        </form>
                    </div>
                </div>
            )}

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(350px, 1fr))', gap: '2rem' }}>
                {customCerts.map(cert => (
                    <div key={cert.id} className="premium-card fade-in" style={{ padding: '2rem', display: 'flex', flexDirection: 'column', position: 'relative' }}>
                        <button
                            onClick={() => handleRemove(cert.id)}
                            style={{ position: 'absolute', top: '1.25rem', right: '1.25rem', color: '#94a3b8', background: 'none', border: 'none', cursor: 'pointer' }}
                            className="hover:text-red-500"
                        >
                            <Trash2 size={18} />
                        </button>

                        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '1.5rem' }}>
                            <div style={{ width: '48px', height: '48px', backgroundColor: '#f0fdf4', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#16a34a' }}>
                                <Award size={24} />
                            </div>
                            <div>
                                <div style={{ fontSize: '0.8rem', fontWeight: 800, color: 'var(--primary)', textTransform: 'uppercase' }}>{cert.provider}</div>
                                <div style={{ fontWeight: 800, fontSize: '1.25rem', color: '#1e293b' }}>{cert.name}</div>
                            </div>
                        </div>

                        <div style={{ flex: 1, color: '#64748b', fontSize: '0.95rem', lineHeight: '1.6', marginBottom: '1.5rem' }}>
                            {cert.description || "No description provided for this custom path."}
                        </div>

                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingTop: '1.5rem', borderTop: '1px solid #f1f5f9' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', color: '#1e293b', fontWeight: 700, fontSize: '0.85rem' }}>
                                <div style={{ width: '6px', height: '6px', background: '#10b981', borderRadius: '50%' }}></div>
                                {cert.status}
                            </div>
                            {cert.url && (
                                <a href={cert.url} target="_blank" rel="noopener noreferrer" style={{ display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--primary)', fontSize: '0.85rem', fontWeight: 700, textDecoration: 'none' }}>
                                    Visit Portal <ExternalLink size={14} />
                                </a>
                            )}
                        </div>
                    </div>
                ))}

                {customCerts.length === 0 && (
                    <div style={{
                        gridColumn: '1 / -1',
                        padding: '10rem 2rem',
                        textAlign: 'center',
                        backgroundColor: '#f8fafc',
                        borderRadius: '48px',
                        border: '3px dashed #e2e8f0',
                        margin: '0 auto',
                        width: '100%',
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        justifyContent: 'center'
                    }}>
                        <div style={{ padding: '2.5rem', backgroundColor: 'white', borderRadius: '50%', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.05)', marginBottom: '2.5rem' }}>
                            <Search size={64} style={{ color: '#94a3b8' }} />
                        </div>
                        <h3 style={{ fontSize: '2.5rem', fontWeight: 900, color: '#1e293b', marginBottom: '1rem' }}>No Custom Journeys Yet</h3>
                        <p style={{ color: '#64748b', fontSize: '1.25rem', maxWidth: '500px' }}>Your personalized certification catalog is waiting. Click the button above to start your self-directed learning journey.</p>
                    </div>
                )}
            </div>
        </div>
    );
}
