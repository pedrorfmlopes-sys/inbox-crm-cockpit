/**
 * client/src/modules/crm/triangulationService.ts
 * 3-Step Matching Algorithm for "The Moat"
 */

import { getAllProtectedProjects, ProtectedProject } from "./excelProvider";

export interface MatchResult {
    isProtected: boolean;
    confidence: number;
    matchedProject?: ProtectedProject;
    reason?: string;
}

export async function scanForProtection(anchors: {
    projectName: string;
    refArticles: string[];
    stakeholders: string[];
    location: string;
}): Promise<MatchResult> {
    const protectedProjects = await getAllProtectedProjects();

    if (protectedProjects.length === 0) return { isProtected: false, confidence: 0 };

    let bestMatch: { project: ProtectedProject, score: number, reason: string } | null = null;

    for (const p of protectedProjects) {
        let score = 0;
        let reasons: string[] = [];

        // Layer 1: Exact/Fuzzy Name Match (>90%)
        const nameSimilarity = calculateSimilarity(anchors.projectName, p.projectName);
        if (nameSimilarity > 0.9) {
            score = 1.0;
            reasons.push("Correspondência exata de nome");
        }

        // Layer 2: Anchor Match (Cross-reference)
        if (score < 1.0) {
            // Check Articles
            const articleMatch = anchors.refArticles.some(ref =>
                p.refArticles?.some(pRef => pRef.toLowerCase().includes(ref.toLowerCase()))
            );

            // Check Stakeholders / Construction Co
            const coMatch = anchors.stakeholders.some(s =>
                p.constructionCo && s.toLowerCase().includes(p.constructionCo.toLowerCase())
            );

            // Check Location
            const locMatch = anchors.location && p.location && (
                anchors.location.toLowerCase().includes(p.location.toLowerCase()) ||
                p.location.toLowerCase().includes(anchors.location.toLowerCase())
            );

            if (articleMatch && (coMatch || locMatch)) {
                score = 0.9;
                reasons.push("Âncora cruzada (Artigo + Local/Construtora)");
            } else if (coMatch && locMatch) {
                score = 0.85;
                reasons.push("Correspondência de Construtora + Localização");
            }
        }

        if (score > (bestMatch?.score || 0)) {
            bestMatch = { project: p, score, reason: reasons.join(", ") };
        }

        if (score === 1.0) break; // Perfect match found
    }

    // Layer 3: Semantic Match (Gemini integration would go here if score < 0.8 but > 0.4)
    // For MVP, we use the heuristic scores.

    return {
        isProtected: (bestMatch?.score || 0) > 0.8,
        confidence: bestMatch?.score || 0,
        matchedProject: bestMatch?.project,
        reason: bestMatch?.reason,
    };
}

function calculateSimilarity(s1: string, s2: string): number {
    if (!s1 || !s2) return 0;
    const longer = s1.length > s2.length ? s1 : s2;
    const shorter = s1.length > s2.length ? s2 : s1;
    if (longer.length === 0) return 1.0;
    return (longer.length - editDistance(longer, shorter)) / longer.length;
}

function editDistance(s1: string, s2: string): number {
    s1 = s1.toLowerCase();
    s2 = s2.toLowerCase();
    const costs = [];
    for (let i = 0; i <= s1.length; i++) {
        let lastValue = i;
        for (let j = 0; j <= s2.length; j++) {
            if (i === 0) costs[j] = j;
            else {
                if (j > 0) {
                    let newValue = costs[j - 1];
                    if (s1.charAt(i - 1) !== s2.charAt(j - 1))
                        newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
                    costs[j - 1] = lastValue;
                    lastValue = newValue;
                }
            }
        }
        if (i > 0) costs[s2.length] = lastValue;
    }
    return costs[s2.length];
}
