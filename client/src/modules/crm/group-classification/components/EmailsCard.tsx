import React from "react";
import { PanelState } from "@/ui/PanelState";
import { makeEmailKey, buildCompactEmailMeta, buildEmailPreviewText } from "../documentUtils";
import { type RelatedEmailEntry } from "@/api";

export interface EmailsCardProps {
  style: React.CSSProperties;
  loading: boolean;
  visibleEmails: RelatedEmailEntry[];
  selectedEmail: RelatedEmailEntry | null;
  emailSearch: string;
  setEmailSearch: (value: string) => void;
  selectAllVisibleEmails: () => void;
  clearSelectedTargets: () => void;
  selectedTargetCount: number;
  selectedTargetEmailKeys: string[];
  toggleTargetEmailKey: (key: string) => void;
  expandedEmailKeys: string[];
  toggleExpandedEmailKey: (key: string) => void;
  setSelectedEmailKey: (key: string) => void;
}

const EmailsCard: React.FC<EmailsCardProps> = ({
  style,
  loading,
  visibleEmails,
  selectedEmail,
  emailSearch,
  setEmailSearch,
  selectAllVisibleEmails,
  clearSelectedTargets,
  selectedTargetCount,
  selectedTargetEmailKeys,
  toggleTargetEmailKey,
  expandedEmailKeys,
  toggleExpandedEmailKey,
  setSelectedEmailKey
}) => {
  return (
    <section style={style} data-testid="emails-card">
      <div style={S.sectionHeaderCompact}>
        <div>
          <div style={S.sectionTitle}>Emails</div>
          <div style={S.sectionSubtitle}>Exploracao do caso</div>
        </div>
        <span style={S.cardMeta}>Selecionados: {selectedTargetCount}</span>
      </div>
      <div style={S.emailControlsRow}>
        <input 
          style={S.input} 
          value={emailSearch} 
          onChange={(event) => setEmailSearch(event.target.value)} 
          placeholder="Pesquisar por assunto, remetente ou texto..." 
        />
        <div style={S.emailToolsInline}>
          <button type="button" style={S.linkBtn} onClick={selectAllVisibleEmails}>Todos visiveis</button>
          <button data-testid="clear-selected-emails" type="button" style={S.linkBtn} onClick={clearSelectedTargets}>Limpar</button>
        </div>
      </div>
      <div style={S.topCardScroll} data-testid="emails-list">
        {loading ? <PanelState compact tone="loading" title="A carregar emails" description="A preparar a lista desta nova janela." /> : null}
        {!loading && !visibleEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Ajusta os filtros ou muda a fonte da lista." /> : null}
        {!loading && visibleEmails.map((email) => {
          const emailKey = makeEmailKey(email);
          const expanded = expandedEmailKeys.includes(emailKey);
          const active = emailKey === makeEmailKey(selectedEmail || {});
          return (
            <div
              key={`compact-${emailKey}`}
              style={active ? S.emailOn : S.email}
              role="button"
              tabIndex={0}
              onClick={() => setSelectedEmailKey(emailKey)}
              onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === " ") {
                  event.preventDefault();
                  setSelectedEmailKey(emailKey);
                }
              }}
            >
              <div style={S.emailTop}>
                <label style={S.emailPick} onClick={(event) => event.stopPropagation()}>
                  <input
                    type="checkbox"
                    checked={selectedTargetEmailKeys.includes(emailKey)}
                    onChange={() => toggleTargetEmailKey(emailKey)}
                  />
                  <span style={S.emailSubject}>{email.subject || "(sem assunto)"}</span>
                </label>
                <div style={S.emailTopRight}>
                  <span style={S.emailMeta}>{buildCompactEmailMeta(email) || "--"}</span>
                  {Array.isArray(email.attachments) && email.attachments.length ? <span style={S.counter}>{email.attachments.length}</span> : null}
                  <button
                    type="button"
                    style={S.chevronBtn}
                    onClick={(event) => {
                      event.stopPropagation();
                      toggleExpandedEmailKey(emailKey);
                    }}
                  >
                    {expanded ? "\u2303" : "\u2304"}
                  </button>
                </div>
              </div>
              {expanded ? (
                <div style={S.emailSnippet}>{buildEmailPreviewText(email) || "Sem preview curto disponivel."}</div>
              ) : null}
            </div>
          );
        })}
      </div>
    </section>
  );
};

const S: Record<string, React.CSSProperties> = {
  sectionHeaderCompact: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 },
  sectionTitle: { fontSize: 12, fontWeight: 600, color: "var(--skin-text-main)" },
  sectionSubtitle: { fontSize: 10, color: "var(--skin-text-muted)" },
  cardMeta: { fontSize: 10, color: "var(--skin-text-muted)" },
  emailControlsRow: { display: "flex", gap: 8, marginBottom: 8, alignItems: "center" },
  input: { flex: 1, height: 28, fontSize: 11, padding: "0 8px", borderRadius: 4, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-input)", color: "var(--skin-text-main)" },
  emailToolsInline: { display: "flex", gap: 12 },
  linkBtn: { background: "none", border: "none", padding: 0, color: "var(--skin-accent-main)", fontSize: 10, cursor: "pointer", textDecoration: "underline" },
  topCardScroll: { flex: 1, overflowY: "auto", minHeight: 0, display: "flex", flexDirection: "column", gap: 1 },
  email: { padding: "6px 8px", background: "var(--skin-bg-card)", border: "1px solid var(--skin-border-main)", borderRadius: 4, cursor: "pointer", display: "flex", flexDirection: "column", gap: 2, textAlign: "left", width: "100%", position: "relative" },
  emailOn: { padding: "6px 8px", background: "var(--skin-bg-active)", border: "1px solid var(--skin-accent-main)", borderRadius: 4, cursor: "pointer", display: "flex", flexDirection: "column", gap: 2, textAlign: "left", width: "100%", position: "relative" },
  emailTop: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 },
  emailPick: { display: "flex", alignItems: "center", gap: 6, flex: 1, overflow: "hidden", cursor: "pointer" },
  emailSubject: { fontSize: 11, fontWeight: 500, color: "var(--skin-text-main)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  emailTopRight: { display: "flex", alignItems: "center", gap: 6, flexShrink: 0 },
  emailMeta: { fontSize: 10, color: "var(--skin-text-muted)" },
  counter: { fontSize: 9, background: "var(--skin-bg-muted)", padding: "1px 4px", borderRadius: 10, color: "var(--skin-text-muted)" },
  chevronBtn: { background: "none", border: "none", padding: 0, color: "var(--skin-text-muted)", cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", justifyContent: "center", width: 16, height: 16 },
  emailSnippet: { fontSize: 10, color: "var(--skin-text-muted)", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden", marginTop: 4, lineHeight: "1.3" },
};

export default EmailsCard;
