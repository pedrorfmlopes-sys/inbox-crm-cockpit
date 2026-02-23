import { getOutlookContext, getCurrentItemToken } from "./office";

async function diagnoseSync() {
    console.log("--- Sync Diagnosis ---");
    const OfficeAny = (window as any).Office;
    console.log("Office object exists:", !!OfficeAny);
    if (OfficeAny) {
        console.log("Office context exists:", !!OfficeAny.context);
        if (OfficeAny.context) {
            console.log("Mailbox exists:", !!OfficeAny.context.mailbox);
            if (OfficeAny.context.mailbox) {
                console.log("Item exists:", !!OfficeAny.context.mailbox.item);
            }
        }
    }

    const syncToken = await getCurrentItemToken();
    const asyncCtx = await getOutlookContext();

    console.log("Sync Token (from sync properties):", syncToken);
    console.log("Async Context ID (from getOutlookContext):", asyncCtx.itemId || asyncCtx.conversationId);

    // Check if what's in window.location query params matches current item
    const qp = new URLSearchParams(window.location.search);
    console.log("Query Param conversationId:", qp.get("conversationId"));

    if (syncToken.includes(asyncCtx.itemId || "") || syncToken.includes(asyncCtx.conversationId || "")) {
        console.log("Result: Sync and Async match.");
    } else {
        console.warn("Result: MISMATCH! Sync properties might be stale.");
    }
}

// This is intended to be run in the browser console of the add-in
(window as any).diagnoseSync = diagnoseSync;
