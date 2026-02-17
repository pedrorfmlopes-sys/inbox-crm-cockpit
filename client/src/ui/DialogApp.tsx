import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  createOdoo,
  linkEmailToRecord,
  odooPing,
  readOdoo,
  searchOdoo,
  searchOdooDomain,
  writeOdoo,
} from "../api";

import DebugPanel from "./DebugPanel";
import { getSettings } from "../settings";
import { applySkin } from "./skins";

type Mode = "new" | "add" | "edit";
type Entity = "project.task" | "project.project" | "crm.lead" | "res.partner";

type Recipient = { name: string; email: string };

function parseRecipientsParam(raw: string): Recipient[] {
  if (!raw) return [];
  return raw
    .split(";")
    .map((part) => {
      const [name, email] = part.split("|");
      return { name: String(name || "").trim(), email: String(email || "").trim() };
    })
    .filter((r) => r.email);
}

function qp() {
  return new URLSearchParams(window.location.search);
}

function getMode(): Mode {
  const m = (qp().get("mode") || "new").toLowerCase();
  return m === "add" || m === "edit" ? (m as Mode) : "new";
}

type Ctx = {
  conversationId: string;
  internetMessageId: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  receivedAtIso: string;
  emailWebLink?: string;

  toR: Recipient[];
  ccR: Recipient[];
};

function getCtxFromQuery(): Ctx {
  const p = qp();
  return {
    conversationId: p.get("conversationId") || "",
    internetMessageId: p.get("internetMessageId") || "",
    subject: p.get("subject") || "",
    fromEmail: p.get("fromEmail") || "",
    fromName: p.get("fromName") || "",
    receivedAtIso: p.get("receivedAtIso") || p.get("receivedDateTimeIso") || "",
    emailWebLink: p.get("emailWebLink") || "",
    toR: parseRecipientsParam(p.get("toR") || ""),
    ccR: parseRecipientsParam(p.get("ccR") || ""),
  };
}

function closeDialog() {
  // @ts-ignore global
  if (typeof Office !== "undefined" && Office?.context?.ui?.messageParent) {
    // @ts-ignore global
    Office.context.ui.messageParent("close");
    return;
  }
  window.close();
}

function shortId(s: string, head = 10, tail = 8) {
  if (!s) return "—";
  if (s.length <= head + tail + 3) return s;
  return `${s.slice(0, head)}...${s.slice(-tail)}`;
}

async function copyToClipboard(text: string) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // fallback
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  }
}

type TypeaheadPickerProps = {
  label: string;
  placeholder: string;
  model: string;
  fields?: string[];
  limit?: number;
  pickedId: number | null;
  pickedName: string;
  onPick: (it: any) => void;
  extraDomain?: (q: string) => any[];
};

function TypeaheadPicker({
  label,
  placeholder,
  model,
  fields = ["id", "name", "display_name"],
  limit = 15,
  pickedId,
  pickedName,
  onPick,
  extraDomain,
}: TypeaheadPickerProps) {
  const [q, setQ] = useState("");
  const [items, setItems] = useState<any[]>([]);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const debounceRef = useRef<number | null>(null);

  const effectiveText = pickedId ? pickedName : q;

  async function load(query: string) {
    setBusy(true);
    try {
      if (extraDomain) {
        const domain = extraDomain(query);
        const rows = await searchOdooDomain(model, domain, fields, limit);
        setItems(Array.isArray(rows) ? rows : []);
      } else {
        const rows = await searchOdoo(model, query, limit);
        setItems(Array.isArray(rows) ? rows : []);
      }
    } finally {
      setBusy(false);
    }
  }

  function scheduleLoad(query: string) {
    if (debounceRef.current) window.clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => load(query), 250);
  }

  useEffect(() => {
    if (!open) return;
    // quando abre, carrega logo (mesmo vazio) para mostrar 10–15
    scheduleLoad(pickedId ? "" : q);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  useEffect(() => {
    if (!open) return;
    if (pickedId) return; // quando já está selecionado, não pesquisa
    scheduleLoad(q);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q]);

  return (
    <div style={{ marginTop: 10, position: "relative" }}>
      <label style={S.labBlock}>{label}</label>

      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <input
          style={{ ...S.input, flex: 1, minWidth: 0 }}
          value={effectiveText}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          onChange={(e) => {
            const v = e.target.value;
            if (pickedId) {
              // se começar a escrever, limpa seleção
              setQ(v);
              onPick({ id: null, name: "" });
            } else {
              setQ(v);
            }
            setOpen(true);
          }}
          placeholder={placeholder}
        />

        {pickedId ? (
          <button
            style={S.btn2}
            onClick={() => {
              onPick({ id: null, name: "" });
              setQ("");
              setOpen(true);
              load("");
            }}
            title="Limpar seleção"
          >
            Limpar
          </button>
        ) : (
          <button style={S.btn2} onClick={() => load(q)} disabled={busy} title="Forçar pesquisa">
            {busy ? "…" : "Pesquisar"}
          </button>
        )}
      </div>

      {pickedId ? (
        <div style={{ marginTop: 6, fontSize: 12, color: "#666" }}>
          Selecionado: {pickedName} (#{pickedId})
        </div>
      ) : null}

      {open && (items?.length || busy) ? (
        <div style={S.pickList}>
          {busy && !items.length ? (
            <div style={{ padding: 10, color: "#777", fontSize: 12 }}>A procurar…</div>
          ) : null}

          {items.map((it: any) => (
            <button
              key={it.id}
              style={S.pickItem}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onPick(it);
                setOpen(false);
                setQ("");
              }}
            >
              <b>{it.display_name || it.name || `#${it.id}`}</b>
              <span style={{ color: "#777" }}>#{it.id}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export default function DialogApp() {
  const mode = useMemo(getMode, []);
  const editModel = qp().get("model") || "";
  const editRecordId = Number(qp().get("recordId") || "0");

  const [ctx, setCtx] = useState<Ctx>(() => getCtxFromQuery());
  const [showThread, setShowThread] = useState(false);
  const [entity, setEntity] = useState<Entity>("project.task");
  const [status, setStatus] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      try {
        const st = await getSettings();
        applySkin(st.skinId || 'classic');
      } catch {
        applySkin('classic');
      }
    })();
  }, []);

  // fallback: se query params vierem vazios, tenta ler do Office.js
  useEffect(() => {
    (async () => {
      try {
        await odooPing(); // só para validar que o proxy/API está ok
      } catch (e: any) {
        setStatus(`API/Proxy falhou: ${e?.message || e}`);
      }

      if (ctx.subject || ctx.fromEmail || ctx.conversationId) return;

      // @ts-ignore global
      const OfficeAny = typeof Office !== "undefined" ? Office : null;
      const item = OfficeAny?.context?.mailbox?.item;
      if (!item) return;

      const subject = item.subject || "";
      const from = item.from;
      const fromEmail = from?.emailAddress || "";
      const fromName = from?.displayName || "";
      const conversationId = item.conversationId || "";
      const internetMessageId = item.internetMessageId || "";
      const normalize = (arr: any): Recipient[] =>
        Array.isArray(arr)
          ? arr
              .map((r: any) => ({ name: String(r?.displayName || "").trim(), email: String(r?.emailAddress || "").trim() }))
              .filter((r: any) => r.email)
          : [];

      setCtx((c) => ({
        ...c,
        subject,
        fromEmail,
        fromName,
        conversationId,
        internetMessageId,
        toR: c.toR?.length ? c.toR : normalize(item.to),
        ccR: c.ccR?.length ? c.ccR : normalize(item.cc),
      }));
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (mode === "edit" && editModel) {
      if (editModel === "project.task") setEntity("project.task");
      else if (editModel === "project.project") setEntity("project.project");
      else if (editModel === "crm.lead") setEntity("crm.lead");
      else if (editModel === "res.partner") setEntity("res.partner");
    }
  }, [mode, editModel]);

  return (
    <div style={S.page}>
      <div style={S.top}>
        <div style={S.titleBlock}>
          <img src="/icon-32.png" alt="" style={S.titleLogo} />
          <div style={{ minWidth: 0 }}>
            <div style={S.h1}>Inbox CRM Cockpit</div>
          <div style={S.h2}>{mode === "new" ? "Criar" : mode === "add" ? "Ligar / Atualizar" : "Editar"}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={S.btn2} onClick={() => closeDialog()}>Fechar</button>
        </div>
      </div>

      <div style={S.banner}>
        <div><b>Email:</b> {ctx.subject || "—"}</div>
                <div style={{ color: "#666" }}>{ctx.fromName ? `${ctx.fromName} <${ctx.fromEmail}>` : (ctx.fromEmail || "—")}</div>

        <div style={{ color: "#666", marginTop: 6, fontSize: 12 }}>
          <b title="Destinatários (Para)">Para:</b>{" "}
          {ctx.toR?.length ? ctx.toR.map((r) => r.email).join("; ") : "—"}
        </div>
        <div style={{ color: "#666", marginTop: 2, fontSize: 12 }}>
          <b title="Destinatários (Cc)">Cc:</b>{" "}
          {ctx.ccR?.length ? ctx.ccR.map((r) => r.email).join("; ") : "—"}
        </div>

        <div style={{ color: "#999", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
          {showThread ? (
            <>
              <span>Thread:</span>
              <code title={ctx.conversationId || ""} style={{ fontSize: 11 }}>{shortId(ctx.conversationId)}</code>
              {ctx.conversationId ? (
                <button style={S.btn3} onClick={() => copyToClipboard(ctx.conversationId)}>Copiar</button>
              ) : null}
              <button style={S.threadToggle} onClick={() => setShowThread(false)} title="Ocultar thread">▴</button>
            </>
          ) : (
            <button style={S.threadToggle} onClick={() => setShowThread(true)} title="Mostrar thread">Thread ▾</button>
          )}
        </div>
      </div>

      <div style={S.card}>
        <div style={S.row}>
          <label style={S.lab}>Tipo</label>
          <select style={S.sel} value={entity} onChange={(e) => setEntity(e.target.value as Entity)} disabled={mode === "edit"}>
            <option value="project.task">Tarefa</option>
            <option value="project.project">Projeto</option>
            <option value="crm.lead">Lead</option>
            <option value="res.partner">Contacto</option>
          </select>
        </div>

        <div style={{ marginTop: 6, fontSize: 12, color: "#557" }} title="Modelo técnico no Odoo">
          Modelo: <code style={{ fontSize: 11 }}>{entity}</code>
        </div>

        {mode === "add" ? (
          <AddExistingPanel entity={entity} ctx={ctx} onStatus={setStatus} />
        ) : entity === "project.task" ? (
          <TaskForm mode={mode} ctx={ctx} editId={mode === "edit" ? editRecordId : 0} onStatus={setStatus} />
        ) : entity === "project.project" ? (
          <ProjectForm mode={mode} ctx={ctx} editId={mode === "edit" ? editRecordId : 0} onStatus={setStatus} />
        ) : entity === "crm.lead" ? (
          <LeadForm mode={mode} ctx={ctx} editId={mode === "edit" ? editRecordId : 0} onStatus={setStatus} />
        ) : entity === "res.partner" ? (
          <ContactHubForm mode={mode} ctx={ctx} editId={mode === "edit" ? editRecordId : 0} onStatus={setStatus} />
        ) : (
          <GenericMiniForm mode={mode} ctx={ctx} model={entity} editId={mode === "edit" ? editRecordId : 0} onStatus={setStatus} />
        )}

        {status && <div style={S.alert}>{status}</div>}
      </div>

      <div style={S.footer}>
        <div style={{ color: "#666", fontSize: 12 }}>v6 • Dialog • typeahead (Projetos/Leads/Etapas)</div>
      </div>

      <DebugPanel compact />
    </div>
  );
}

function AddExistingPanel({ entity, ctx, onStatus }: any) {
  const [pickedId, setPickedId] = useState<number | null>(null);
  const [pickedName, setPickedName] = useState("");

  async function link() {
    if (!pickedId) return onStatus("Escolhe um registo para ligar.");
    try {
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: entity,
        recordId: pickedId,
        recordName: pickedName,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <TypeaheadPicker
        label="Selecionar existente"
        placeholder={`Pesquisar ${entity}...`}
        model={entity}
        pickedId={pickedId}
        pickedName={pickedName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPickedId(id);
          setPickedName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={link} disabled={!pickedId}>Ligar ao email</button>
        <button style={S.btn2} onClick={() => closeDialog()}>Cancelar</button>
      </div>
    </div>
  );
}

function TaskForm({ mode, ctx, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [description, setDescription] = useState("");

  const [projectId, setProjectId] = useState<number | null>(null);
  const [projectName, setProjectName] = useState("");

  const [assigneeId, setAssigneeId] = useState<number | null>(null);
  const [assigneeName, setAssigneeName] = useState("");

  const [deadline, setDeadline] = useState("");
  const [stageId, setStageId] = useState<number | null>(null);
  const [stageName, setStageName] = useState("");
  const [stagePick, setStagePick] = useState<any[]>([]);

  const [isSub, setIsSub] = useState(false);
  const [parentId, setParentId] = useState<number | null>(null);
  const [parentName, setParentName] = useState("");

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("project.task", [editId], ["name", "description", "project_id", "user_ids", "date_deadline", "stage_id", "parent_id"]);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setDescription(r.description || "");
        if (r.project_id) { setProjectId(r.project_id[0]); setProjectName(r.project_id[1]); }
        if (Array.isArray(r.user_ids) && r.user_ids.length) {
          const u = await readOdoo("res.users", [r.user_ids[0]], ["name"]);
          setAssigneeId(r.user_ids[0]);
          setAssigneeName(u?.[0]?.name || "");
        }
        if (r.date_deadline) setDeadline(String(r.date_deadline));
        if (r.stage_id) { setStageId(r.stage_id[0]); setStageName(r.stage_id[1]); }
        if (r.parent_id) { setIsSub(true); setParentId(r.parent_id[0]); setParentName(r.parent_id[1]); }
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  useEffect(() => {
    (async () => {
      try {
        if (!projectId) return setStagePick([]);
        const rows = await searchOdooDomain(
          "project.task.type",
          ["|", ["project_ids", "=", false], ["project_ids", "in", [projectId]]],
          ["id", "name"],
          50
        );
        setStagePick(rows || []);
      } catch {
        setStagePick([]);
      }
    })();
  }, [projectId]);

  async function save() {
    try {
      const values: any = {
        name: name || "Nova tarefa",
        description: description || "",
      };
      if (projectId) values.project_id = projectId;
      if (assigneeId) values.user_ids = [[6, 0, [assigneeId]]];
      if (deadline) values.date_deadline = deadline;
      if (stageId) values.stage_id = stageId;
      if (isSub && parentId) values.parent_id = parentId;

      let id = editId;

      if (mode === "edit") {
        await writeOdoo("project.task", id, values);
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      id = await createOdoo("project.task", values);

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "project.task",
        recordId: id,
        recordName: name || "",
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab2}>Título</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Título da tarefa" />
      </div>

      <TypeaheadPicker
        label="Projeto"
        placeholder="Pesquisar projeto…"
        model="project.project"
        pickedId={projectId}
        pickedName={projectName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setProjectId(id);
          setProjectName(id ? (it.display_name || it.name || `#${id}`) : "");
          if (!id) { setStageId(null); setStageName(""); }
        }}
      />

      <TypeaheadPicker
        label="Responsável"
        placeholder="Pesquisar utilizador…"
        model="res.users"
        fields={["id", "name", "display_name", "email"]}
        pickedId={assigneeId}
        pickedName={assigneeName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setAssigneeId(id);
          setAssigneeName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={S.grid2}>
        <PickerStatic
          label="Etapa"
          pickedId={stageId}
          pickedName={stageName}
          items={stagePick}
          onPick={(it: any) => { setStageId(it.id); setStageName(it.name || it.display_name || `#${it.id}`); }}
          placeholder={projectId ? "Escolher etapa..." : "Etapa (opcional)"}
        />

        <div style={S.row}>
          <label style={S.lab2}>Prazo</label>
          <input style={S.input} type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} />
        </div>
      </div>

      <div style={S.row}>
        <label style={S.lab2}>Subtarefa</label>
        <input
          type="checkbox"
          checked={isSub}
          onChange={(e) => {
            setIsSub(e.target.checked);
            if (!e.target.checked) { setParentId(null); setParentName(""); }
          }}
        />
      </div>

      {isSub ? (
        <TypeaheadPicker
          label="Parent task"
          placeholder={projectId ? "Pesquisar tarefa (filtra por projeto)..." : "Pesquisar tarefa (global)..."}
          model="project.task"
          fields={["id", "name", "display_name", "project_id"]}
          pickedId={parentId}
          pickedName={parentName}
          extraDomain={(q) => {
            const d: any[] = [];
            if (projectId) d.push(["project_id", "=", projectId]); // B: filtra se houver projeto
            if (q?.trim()) d.push(["name", "ilike", q.trim()]);
            return d;
          }}
          onPick={(it: any) => {
            const id = it?.id ?? null;
            setParentId(id);
            setParentName(id ? (it.display_name || it.name || `#${id}`) : "");
          }}
        />
      ) : null}

      <div style={{ marginTop: 10 }}>
        <label style={S.labBlock}>Descrição</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Descrição / notas..." />
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar alterações" : "Criar + Ligar ao email"}</button>
        <button style={S.btn2} onClick={() => closeDialog()}>Cancelar</button>
      </div>
    </div>
  );
}


function ProjectForm({ mode, ctx, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [partnerId, setPartnerId] = useState<number | null>(null);
  const [partnerName, setPartnerName] = useState("");
  const [managerId, setManagerId] = useState<number | null>(null);
  const [managerName, setManagerName] = useState("");
  const [description, setDescription] = useState("");

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        // description nem sempre existe, por isso fazemos fallback seguro
        let rows: any[] | null = null;
        try {
          rows = await readOdoo("project.project", [editId], ["name", "partner_id", "user_id", "description"]);
        } catch {
          rows = await readOdoo("project.project", [editId], ["name", "partner_id", "user_id"]);
        }
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        if (r.partner_id) {
          setPartnerId(r.partner_id[0]);
          setPartnerName(r.partner_id[1]);
        }
        if (r.user_id) {
          setManagerId(r.user_id[0]);
          setManagerName(r.user_id[1]);
        }
        if (r.description) setDescription(String(r.description));
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  async function save() {
    try {
      const values: any = { name: name || "Novo projeto" };
      if (partnerId) values.partner_id = partnerId;
      if (managerId) values.user_id = managerId;
      if (description) values.description = description;

      if (mode === "edit") {
        // tentativa com description; se falhar, remove e tenta novamente
        try {
          await writeOdoo("project.project", editId, values);
        } catch {
          const v2 = { ...values };
          delete v2.description;
          await writeOdoo("project.project", editId, v2);
        }
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      let id: number;
      try {
        id = await createOdoo("project.project", values);
      } catch {
        const v2 = { ...values };
        delete v2.description;
        id = await createOdoo("project.project", v2);
      }

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "project.project",
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab2} title="Nome do projeto no Odoo">Nome</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do projeto" />
      </div>

      <TypeaheadPicker
        label="Cliente"
        placeholder="Pesquisar contacto/empresa…"
        model="res.partner"
        pickedId={partnerId}
        pickedName={partnerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPartnerId(id);
          setPartnerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <TypeaheadPicker
        label="Gestor"
        placeholder="Pesquisar utilizador…"
        model="res.users"
        fields={["id", "name", "display_name", "email"]}
        pickedId={managerId}
        pickedName={managerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setManagerId(id);
          setManagerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={{ marginTop: 10 }}>
        <label style={S.labBlock} title="Descrição/Notas do projeto (se o teu Odoo suportar este campo)">Descrição</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Notas do projeto…" />
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={save} title={mode === "edit" ? "Guardar alterações no Odoo" : "Criar no Odoo e ligar ao email"}>
          {mode === "edit" ? "Guardar alterações" : "Criar + Ligar ao email"}
        </button>
        <button style={S.btn2} onClick={() => closeDialog()} title="Fechar sem guardar">
          Cancelar
        </button>
      </div>
    </div>
  );
}

function LeadForm({ mode, ctx, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [contactName, setContactName] = useState(ctx.fromName || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");
  const [phone, setPhone] = useState("");
  const [partnerId, setPartnerId] = useState<number | null>(null);
  const [partnerName, setPartnerName] = useState("");
  const [stageId, setStageId] = useState<number | null>(null);
  const [stageName, setStageName] = useState("");
  const [description, setDescription] = useState("");

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        let rows: any[] | null = null;
        try {
          rows = await readOdoo("crm.lead", [editId], ["name", "contact_name", "email_from", "phone", "partner_id", "stage_id", "description"]);
        } catch {
          rows = await readOdoo("crm.lead", [editId], ["name", "contact_name", "email_from", "phone", "partner_id", "stage_id"]);
        }
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setContactName(r.contact_name || "");
        setEmail(r.email_from || "");
        setPhone(r.phone || "");
        if (r.partner_id) { setPartnerId(r.partner_id[0]); setPartnerName(r.partner_id[1]); }
        if (r.stage_id) { setStageId(r.stage_id[0]); setStageName(r.stage_id[1]); }
        if (r.description) setDescription(String(r.description));
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  async function save() {
    try {
      const values: any = {
        name: name || `Lead: ${ctx.subject || "sem assunto"}`,
        contact_name: contactName || "",
        email_from: email || "",
      };
      if (phone) values.phone = phone;
      if (partnerId) values.partner_id = partnerId;
      if (stageId) values.stage_id = stageId;
      if (description) values.description = description;

      if (mode === "edit") {
        try {
          await writeOdoo("crm.lead", editId, values);
        } catch {
          const v2 = { ...values };
          delete v2.description;
          await writeOdoo("crm.lead", editId, v2);
        }
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      let id: number;
      try {
        id = await createOdoo("crm.lead", values);
      } catch {
        const v2 = { ...values };
        delete v2.description;
        id = await createOdoo("crm.lead", v2);
      }

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "crm.lead",
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab2} title="Título do lead no Odoo">Nome do lead</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do lead" />
      </div>

      <div style={S.row}>
        <label style={S.lab2} title="Nome da pessoa de contacto">Contacto</label>
        <input style={S.input} value={contactName} onChange={(e) => setContactName(e.target.value)} placeholder="Nome do contacto" />
      </div>

      <div style={S.row}>
        <label style={S.lab2} title="Email do lead">Email</label>
        <input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." />
      </div>

      <div style={S.row}>
        <label style={S.lab2} title="Telefone (opcional)">Telefone</label>
        <input style={S.input} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Telefone" />
      </div>

      <TypeaheadPicker
        label="Empresa/Contacto (Odoo)"
        placeholder="Pesquisar res.partner…"
        model="res.partner"
        pickedId={partnerId}
        pickedName={partnerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPartnerId(id);
          setPartnerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <TypeaheadPicker
        label="Etapa"
        placeholder="Pesquisar etapa do lead…"
        model="crm.stage"
        fields={["id", "name"]}
        pickedId={stageId}
        pickedName={stageName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setStageId(id);
          setStageName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={{ marginTop: 10 }}>
        <label style={S.labBlock} title="Notas do lead (se o teu Odoo suportar este campo)">Descrição</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Notas do lead…" />
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={save} title={mode === "edit" ? "Guardar alterações no Odoo" : "Criar no Odoo e ligar ao email"}>
          {mode === "edit" ? "Guardar alterações" : "Criar + Ligar ao email"}
        </button>
        <button style={S.btn2} onClick={() => closeDialog()} title="Fechar sem guardar">
          Cancelar
        </button>
      </div>
    </div>
  );
}

function ContactHubForm({ mode, ctx, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.fromName || ctx.subject || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");
  const [phone, setPhone] = useState("");

  const participants = useMemo(() => {
    const out: Array<{ role: string; name: string; email: string }> = [];
    if (ctx.fromEmail) out.push({ role: "De", name: ctx.fromName || "", email: ctx.fromEmail });
    ctx.toR?.forEach((r: any) => out.push({ role: "Para", name: r.name || "", email: r.email }));
    ctx.ccR?.forEach((r: any) => out.push({ role: "Cc", name: r.name || "", email: r.email }));
    // dedupe by email
    const seen = new Set<string>();
    return out.filter((p) => {
      const key = p.email.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }, [ctx]);

  const [match, setMatch] = useState<Record<string, { id: number; name: string; email?: string } | null>>({});

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("res.partner", [editId], ["name", "email", "phone"]);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setEmail(r.email || "");
        setPhone(r.phone || "");
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  useEffect(() => {
    (async () => {
      const emails = participants.map((p) => p.email).filter(Boolean);
      const next: Record<string, any> = {};
      for (const em of emails) next[em] = null;
      setMatch(next);

      // lookup em série (simplicidade > performance nesta fase)
      for (const em of emails) {
        try {
          const rows = await searchOdooDomain("res.partner", [["email", "ilike", em]], ["id", "name", "display_name", "email"], 5);
          const found = rows?.find((r: any) => String(r.email || "").toLowerCase() === em.toLowerCase()) || rows?.[0];
          setMatch((prev) => ({ ...prev, [em]: found ? { id: found.id, name: found.display_name || found.name || `#${found.id}`, email: found.email } : null }));
        } catch {
          // ignore lookup errors
        }
      }
    })();
  }, [participants]);

  async function saveMain() {
    try {
      if (mode === "edit") {
        await writeOdoo("res.partner", editId, { name: name || email || "Contacto", email, phone });
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      const id = await createOdoo("res.partner", { name: name || email || "Contacto", email, phone });
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "res.partner",
        recordId: id,
        recordName: name || email,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function linkToPartner(id: number, display: string) {
    try {
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "res.partner",
        recordId: id,
        recordName: display,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });
      onStatus(`Ligado a ${display} ✅`);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function createPartnerFrom(p: any) {
    try {
      const id = await createOdoo("res.partner", { name: p.name || p.email, email: p.email });
      await linkToPartner(id, p.name || p.email);
      setMatch((prev) => ({ ...prev, [p.email]: { id, name: p.name || p.email, email: p.email } }));
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab2} title="Nome do contacto no Odoo">Nome</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do contacto" />
      </div>

      <div style={S.row}>
        <label style={S.lab2} title="Email do contacto">Email</label>
        <input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." />
      </div>

      <div style={S.row}>
        <label style={S.lab2} title="Telefone (opcional)">Telefone</label>
        <input style={S.input} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Telefone" />
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={saveMain} title={mode === "edit" ? "Guardar alterações no Odoo" : "Criar no Odoo e ligar ao email"}>
          {mode === "edit" ? "Guardar alterações" : "Criar + Ligar ao email"}
        </button>
        <button style={S.btn2} onClick={() => closeDialog()} title="Fechar sem guardar">
          Cancelar
        </button>
      </div>

      <div style={{ marginTop: 16, borderTop: "1px solid #e9eefc", paddingTop: 12 }}>
        <div style={{ fontWeight: 900, marginBottom: 6 }} title="Inspirado no HubSpot: identifica participantes e permite ligar/criar contactos">
          Participantes no email
        </div>

        {participants.length === 0 ? (
          <div style={{ color: "#557" }}>Sem participantes disponíveis.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {participants.map((p) => {
              const m = match[p.email];
              return (
                <div key={p.email} style={S.partRow}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 800 }}>
                      <span style={S.badge} title="Origem do endereço">{p.role}</span> {p.name ? `${p.name} <${p.email}>` : p.email}
                    </div>
                    <div style={{ fontSize: 12, color: "#557" }}>
                      {m ? `Odoo: ${m.name} (#${m.id})` : "Odoo: não encontrado"}
                    </div>
                  </div>

                  {m ? (
                    <button style={S.btn3} onClick={() => linkToPartner(m.id, m.name)} title="Criar ligação oculta email ↔ contacto (esta conversa)">
                      Ligar
                    </button>
                  ) : (
                    <button style={S.btn3} onClick={() => createPartnerFrom(p)} title="Criar novo contacto no Odoo e ligar ao email">
                      Criar
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function GenericMiniForm({ mode, ctx, model, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const fields =
          model === "res.partner" ? ["name", "email"] :
          model === "crm.lead" ? ["name", "email_from"] :
          ["name"];
        const rows = await readOdoo(model, [editId], fields);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        if (model === "res.partner") setEmail(r.email || "");
        if (model === "crm.lead") setEmail(r.email_from || "");
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  async function save() {
    try {
      if (mode === "edit") {
        const values: any = { name: name || "Atualizado" };
        if (model === "res.partner") values.email = email;
        if (model === "crm.lead") values.email_from = email;
        await writeOdoo(model, editId, values);
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      const values: any =
        model === "res.partner" ? { name: name || email || "Novo contacto", email } :
        model === "crm.lead" ? { name: name || `Lead: ${ctx.subject || "sem assunto"}`, email_from: email } :
        { name: name || `Novo: ${ctx.subject || ""}` };

      const id = await createOdoo(model, values);

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model,
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab2}>{model === "crm.lead" ? "Nome do lead" : model === "res.partner" ? "Nome do contacto" : "Nome"}</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome" />
      </div>

      {(model === "crm.lead" || model === "res.partner") ? (
        <div style={S.row}>
          <label style={S.lab2}>Email</label>
          <input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." />
        </div>
      ) : null}

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar alterações" : "Criar + Ligar ao email"}</button>
        <button style={S.btn2} onClick={() => closeDialog()}>Cancelar</button>
      </div>
    </div>
  );
}

function PickerStatic({ label, pickedId, pickedName, items, onPick, placeholder }: any) {
  return (
    <div style={{ marginTop: 10 }}>
      <label style={S.labBlock}>{label}</label>
      <div style={{ display: "flex", gap: 8 }}>
        <input style={S.input} value={pickedId ? pickedName : ""} readOnly placeholder={placeholder} />
      </div>
      {items?.length ? (
        <div style={S.pickList}>
          {items.map((it: any) => (
            <button key={it.id} style={S.pickItem} onClick={() => onPick(it)}>
              <b>{it.display_name || it.name || `#${it.id}`}</b>
              <span style={{ color: "#777" }}>#{it.id}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  page: { fontFamily: "system-ui,Segoe UI,Arial", padding: 16, background: "#f7f9ff", minHeight: "100vh", color: "#0b3d91" },
  top: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 },
  h1: { fontWeight: 900, fontSize: 20, color: "#0b3d91" },
  h2: { color: "#557", marginTop: 2 },
  banner: { background: "#fff", border: "1px solid #d6def2", borderRadius: 12, padding: 12, marginBottom: 12 },
  card: { background: "#fff", border: "1px solid #d6def2", borderRadius: 12, padding: 12 },
  footer: { marginTop: 10, display: "flex", justifyContent: "space-between", color: "#557" },

  row: { display: "grid", gridTemplateColumns: "140px 1fr", gap: 10, alignItems: "center", marginTop: 10 },
  lab: { fontWeight: 800, color: "#0b3d91" },
  lab2: { fontWeight: 800, color: "#0b3d91" },
  labBlock: { display: "block", fontWeight: 800, marginBottom: 6, color: "#0b3d91" },

  sel: { padding: "8px 10px", border: "1px solid #d6def2", borderRadius: 10, color: "#0b3d91", background: "#fff" },
  input: { width: "100%", padding: "8px 10px", border: "1px solid #d6def2", borderRadius: 10, color: "#122", background: "#fff" },
  ta: { width: "100%", minHeight: 90, padding: "10px 10px", border: "1px solid #d6def2", borderRadius: 10, resize: "vertical", color: "#122" },

  btn: { padding: "10px 12px", borderRadius: 10, border: "1px solid #0b3d91", background: "#0b3d91", color: "#fff", fontWeight: 900, cursor: "pointer" },
  btn2: { padding: "10px 12px", borderRadius: 10, border: "1px solid #d6def2", background: "#fff", cursor: "pointer", color: "#0b3d91", fontWeight: 900 },
  btn3: { padding: "6px 10px", borderRadius: 10, border: "1px solid #d6def2", background: "#fff", cursor: "pointer", fontSize: 12, color: "#0b3d91", fontWeight: 800 },

  alert: { marginTop: 12, padding: 10, borderRadius: 10, border: "1px solid #f1d39a", background: "#fff7e6", color: "#623" },

  grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },

  pickList: {
    position: "absolute",
    left: 0,
    right: 0,
    top: "100%",
    marginTop: 6,
    background: "#fff",
    border: "1px solid #d6def2",
    borderRadius: 12,
    maxHeight: 240,
    overflow: "auto",
    zIndex: 999,
    boxShadow: "0 8px 24px rgba(0,0,0,0.08)",
  },
  pickItem: {
    width: "100%",
    textAlign: "left",
    padding: "10px 12px",
    border: "none",
    background: "transparent",
    cursor: "pointer",
    display: "flex",
    justifyContent: "space-between",
    gap: 10,
    color: "#122",
  },

  partRow: {
    display: "flex",
    gap: 10,
    alignItems: "center",
    padding: 10,
    borderRadius: 12,
    border: "1px solid #e9eefc",
    background: "#f7f9ff",
  },
  badge: {
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: 999,
    border: "1px solid #d6def2",
    background: "#fff",
    color: "#0b3d91",
    fontSize: 12,
    marginRight: 6,
  },
  threadToggle: {
    border: "1px solid var(--iccc-border, #d7dbeb)",
    background: "var(--iccc-card, #ffffff)",
    borderRadius: 999,
    padding: "2px 8px",
    fontSize: 11,
    cursor: "pointer",
    color: "var(--iccc-text, #0b2d6b)",
  },
};
