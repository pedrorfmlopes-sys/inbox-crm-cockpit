
import { buildPrompt } from './src/ai/promptTemplates.js';

const mockEmail = {
    subject: "Reunião de Projeto",
    from: "joao@exemplo.com",
    bodyText: "Olá Pedro, podemos confirmar a reunião para amanhã às 10h?"
};

const scenarios = [
    { name: "Morning Greeting", time: "2024-03-01T09:00:00Z", action: "reply" },
    { name: "Afternoon Greeting", time: "2024-03-01T15:00:00Z", action: "reply" },
    { name: "Evening Greeting", time: "2024-03-01T21:00:00Z", action: "reply" },
    { name: "Forward with Role Analysis", time: "2024-03-01T10:00:00Z", action: "forward", inputText: "Reenvia à Nerea" }
];

console.log("=== PROMPT VERIFICATION ===\n");

for (const s of scenarios) {
    console.log(`--- Scenario: ${s.name} ---`);
    const prompt = buildPrompt({
        action: s.action,
        email: mockEmail,
        currentTime: s.time,
        inputText: s.inputText || ""
    });

    // Check for greeting
    if (s.name.includes("Morning") && prompt.includes("Bom dia,")) console.log("✅ Morning greeting found");
    else if (s.name.includes("Afternoon") && prompt.includes("Boa tarde,")) console.log("✅ Afternoon greeting found");
    else if (s.name.includes("Evening") && prompt.includes("Boa noite,")) console.log("✅ Evening greeting found");

    // Check for structure rules
    if (prompt.includes("ESTRUTURA PROFISSIONAL")) console.log("✅ Professional structure rules found");
    if (prompt.includes("AGRADECIMENTO FINAL")) console.log("✅ Empathetic thanks rules found");

    if (s.action === "forward" && prompt.includes("INTELIGÊNCIA SOCIAL")) console.log("✅ Social intelligence rules found");

    console.log("\n");
}
