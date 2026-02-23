function l(e){return["BEGIN:VCARD","VERSION:3.0",`FN:${e.name}`,e.company?`ORG:${e.company}`:"",e.jobTitle?`TITLE:${e.jobTitle}`:"",e.email?`EMAIL;TYPE=INTERNET,WORK:${e.email}`:"",e.phone?`TEL;TYPE=WORK,VOICE:${e.phone}`:"",e.mobile?`TEL;TYPE=CELL,VOICE:${e.mobile}`:"",e.website?`URL:${e.website}`:"","END:VCARD"].filter(Boolean).join(`
`)}function d(e){const a=l(e),i=new Blob([a],{type:"text/vcard"}),t=URL.createObjectURL(i),o=document.createElement("a");o.href=t,o.download=`${e.name.replace(/\s+/g,"_")}.vcf`,document.body.appendChild(o),o.click(),document.body.removeChild(o),URL.revokeObjectURL(t)}function r(e){return`
<div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 13px; line-height: 1.2; color: #1e293b; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; max-width: 300px; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
    <div style="font-weight: 800; color: #0f172a; margin-bottom: 4px;">${e.name}</div>
    ${e.jobTitle?`<div style="color: #64748b; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.02em; margin-bottom: 8px;">${e.jobTitle}</div>`:""}
    <div style="display: grid; gap: 4px;">
        ${e.company?`<div style="display: flex; gap: 6px;"><b>Empresa:</b> ${e.company}</div>`:""}
        ${e.email?`<div style="display: flex; gap: 6px;"><b>Email:</b> ${e.email}</div>`:""}
        ${e.phone?`<div style="display: flex; gap: 6px;"><b>Tel:</b> ${e.phone}</div>`:""}
        ${e.mobile?`<div style="display: flex; gap: 6px;"><b>Móvel:</b> ${e.mobile}</div>`:""}
    </div>
</div>
`.trim()}async function p(e){const a=r(e),i=`${e.name}
${e.jobTitle||""}
${e.company||""}
${e.email||""}
${e.phone||""}
${e.mobile||""}`.trim();try{const t="text/html",o=new Blob([a],{type:t}),n=[new ClipboardItem({[t]:o,"text/plain":new Blob([i],{type:"text/plain"})})];return await navigator.clipboard.write(n),!0}catch(t){return console.error("Failed to copy visual card:",t),await navigator.clipboard.writeText(i),!1}}export{p as copyVisualCardToClipboard,d as downloadVCard,l as generateVCard,r as generateVisualCardHtml};
