import React, { useState } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import * as Icons from "../../ui/icons";

export const LoginCockpit: React.FC = () => {
    const cockpit = useCockpit();
    if (!cockpit) return null;
    const { login } = cockpit;
    const [url, setUrl] = useState("https://divitek.thinkopen.solutions");
    const [db, setDb] = useState("divitek_studio");
    const [username, setUsername] = useState("pedrolopes@divitek.pt");
    const [password, setPassword] = useState("1234");
    const [error, setError] = useState<string | null>(null);
    const [isSubmitting, setIsSubmitting] = useState(false);

    const handleSubmit = async (e: React.FormEvent) => {
        e.preventDefault();
        setError(null);
        setIsSubmitting(true);

        try {
            await login({ url, db, login: username, password });
        } catch (err: any) {
            setError(err.message || "Erro ao fazer login. Verifica as credenciais.");
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <div style={S.container}>
            <div style={S.header}>
                <div style={S.iconBox}>
                    <Icons.Lock size={32} color="var(--iccc-pill-active-bg)" />
                </div>
                <h2 style={S.title}>Bem-vindo ao InboxCockpit</h2>
                <p style={S.subtitle}>Liga a tua conta Odoo para começar</p>
            </div>

            <form style={S.form} onSubmit={handleSubmit}>
                <div style={S.inputGroup}>
                    <label style={S.label}>URL do Odoo</label>
                    <input
                        style={S.input}
                        type="url"
                        placeholder="https://o-teu-odoo.com"
                        value={url}
                        onChange={(e) => setUrl(e.target.value)}
                        required
                    />
                </div>

                <div style={S.inputGroup}>
                    <label style={S.label}>Base de Dados</label>
                    <input
                        style={S.input}
                        type="text"
                        placeholder="odoo_db_name"
                        value={db}
                        onChange={(e) => setDb(e.target.value)}
                        required
                    />
                </div>

                <div style={S.inputGroup}>
                    <label style={S.label}>Login / Email</label>
                    <input
                        style={S.input}
                        type="email"
                        placeholder="email@exemplo.com"
                        value={username}
                        onChange={(e) => setUsername(e.target.value)}
                        required
                    />
                </div>

                <div style={S.inputGroup}>
                    <label style={S.label}>Palavra-passe / API Key</label>
                    <input
                        style={S.input}
                        type="password"
                        placeholder="••••••••"
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        required
                    />
                </div>

                {error && <div style={S.error}>{error}</div>}

                <button style={S.loginBtn} type="submit" disabled={isSubmitting}>
                    {isSubmitting ? (
                        <Icons.RotateCcw size={18} style={{ animation: "spin 1s linear infinite" }} />
                    ) : (
                        "Entrar"
                    )}
                </button>
            </form>

            <div style={S.footer}>
                Tens dúvidas? Fala com o administrador do Odoo.
            </div>

            <style>{`
                @keyframes spin { 100% { transform: rotate(360deg); } }
            `}</style>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    container: {
        display: "flex",
        flexDirection: "column",
        padding: "32px 24px",
        height: "100%",
        justifyContent: "center",
        maxWidth: "400px",
        margin: "0 auto",
    },
    header: {
        textAlign: "center",
        marginBottom: "32px",
    },
    iconBox: {
        display: "inline-flex",
        padding: "16px",
        background: "rgba(37, 99, 235, 0.1)",
        borderRadius: "20px",
        marginBottom: "16px",
    },
    title: {
        fontSize: "16px",
        fontWeight: 600,
        color: "#172B4D",
        margin: "0 0 8px 0",
    },
    subtitle: {
        fontSize: "14px",
        color: "var(--iccc-text-muted)",
        margin: 0,
    },
    form: {
        display: "flex",
        flexDirection: "column",
        gap: "16px",
    },
    inputGroup: {
        display: "flex",
        flexDirection: "column",
        gap: "6px",
    },
    label: {
        fontSize: "11px",
        fontWeight: 600,
        color: "#6B778C",
        paddingLeft: "2px",
    },
    input: {
        padding: "8px 12px",
        background: "#FAFBFC",
        border: "2px solid #DFE1E6",
        borderRadius: "3px",
        fontSize: "13px",
        color: "#172B4D",
        outline: "none",
        transition: "background 0.2s, border-color 0.2s",
    },
    loginBtn: {
        marginTop: "12px",
        background: "#0052CC",
        color: "white",
        border: "none",
        borderRadius: "3px",
        padding: "10px",
        fontSize: "14px",
        fontWeight: 500,
        cursor: "pointer",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        height: "40px",
    },
    error: {
        padding: "12px",
        background: "#fee2e2",
        color: "#991b1b",
        borderRadius: "8px",
        fontSize: "12px",
        fontWeight: 600,
        textAlign: "center",
    },
    footer: {
        marginTop: "40px",
        textAlign: "center",
        fontSize: "11px",
        color: "var(--iccc-text-muted)",
    },
};
