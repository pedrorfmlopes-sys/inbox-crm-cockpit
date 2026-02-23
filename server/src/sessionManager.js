import { odooClientFromEnv } from "./odoo.js";

/**
 * Simple in-memory session manager for Odoo clients.
 */
class SessionManager {
    constructor() {
        this.sessions = new Map();
    }

    async createSession(credentials) {
        const client = await odooClientFromEnv(credentials);

        const token = Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
        this.sessions.set(token, { client, credentials });
        return { token, meta: client.meta };
    }

    getSession(token) {
        return this.sessions.get(token);
    }

    removeSession(token) {
        this.sessions.delete(token);
    }
}

export const sessionManager = new SessionManager();
