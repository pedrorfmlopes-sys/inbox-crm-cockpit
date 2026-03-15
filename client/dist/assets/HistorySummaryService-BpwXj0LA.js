import{a as n}from"./index-DvxmcHo8.js";async function a(o,e){const t=`
OUTLOOK RECENT HISTORY:
${o.outlookHistory||"No recent history."}

ODOO CHATTER/NOTES:
${o.odooChatter||"No internal notes."}

PROTECTION MOAT STATUS:
${o.protectionStatus||"Unknown"}
`.trim();try{const r=await n(t,[],e);if(r.ok)return r.summary;throw new Error("Briefing generation failed")}catch(r){return console.error("[HistorySummaryService] Error:",r),"Erro ao gerar resumo. Tente novamente."}}export{a as get30SecondBriefing};
