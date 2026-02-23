export interface TestScenario {
    id: string;
    name: string;
    description: string;
    context: {
        mode: "new" | "edit";
        entity: "project.task" | "project.project" | "crm.lead";
        editId?: string;
        conversationId: string;
        subject: string;
        fromEmail: string;
        fromName: string;
        receivedAtIso: string;
    };
    bodyText: string;
    attachments: Array<{ name: string; contentType: string; content: string }>;
    expectedResults: {
        aiTriggers: boolean;
        odooMemoryFound: boolean;
        queueableActions: boolean;
    };
}

export const SCENARIOS: TestScenario[] = [
    {
        id: "A",
        name: "Scenario A: New Client (Stress)",
        description: "Email with 500+ words, no partnerId, 2 attachments.",
        context: {
            mode: "new",
            entity: "project.task",
            conversationId: "conv_A_123",
            subject: "Proposta para novo empreendimento - Villa Sol",
            fromEmail: "alexandre.sa@divitek.pt",
            fromName: "Alexandre Sá",
            receivedAtIso: new Date().toISOString(),
        },
        bodyText: `Caro Pedro,\n\nEspero que estejas bem. Estou a contactar-te para dar seguimento ao nosso diálogo sobre o novo projeto Villa Sol. Como sabes, este é um empreendimento de grande escala e precisamos de garantir que todos os detalhes técnicos estão alinhados.\n\n` + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(60) + `\n\nEm anexo envio as plantas e o caderno de encargos.\n\nAbraço,\nAlexandre`,
        attachments: [
            { name: "plantas_villa_sol.pdf", contentType: "application/pdf", content: "MOCK_BASE64_PLANTAS" },
            { name: "caderno_encargos.pdf", contentType: "application/pdf", content: "MOCK_BASE64_CADERNO" },
        ],
        expectedResults: {
            aiTriggers: true,
            odooMemoryFound: false,
            queueableActions: true,
        },
    },
    {
        id: "B",
        name: "Scenario B: Existing Project",
        description: "Known contact with 3 open tasks in Odoo.",
        context: {
            mode: "edit",
            entity: "project.task",
            editId: "456",
            conversationId: "conv_B_456",
            subject: "RE: Projeto Central Park - Atualização",
            fromEmail: "marcos@empresa.com",
            fromName: "Marcos Silva",
            receivedAtIso: new Date().toISOString(),
        },
        bodyText: "Olá! Precisamos de agendar a reunião de obra para a próxima semana. Detetei também a necessidade de rever as medições do átrio.",
        attachments: [],
        expectedResults: {
            aiTriggers: true,
            odooMemoryFound: true,
            queueableActions: false, // edit mode creates instantly
        },
    },
    {
        id: "C",
        name: "Scenario C: Complex Task",
        description: "Email requiring multiple sub-tasks. Testing queueing.",
        context: {
            mode: "new",
            entity: "project.task",
            conversationId: "conv_C_789",
            subject: "Lista de pendentes - Hotel Funchal",
            fromEmail: "gestor@hotel.com",
            fromName: "Sara Mendes",
            receivedAtIso: new Date().toISOString(),
        },
        bodyText: "Bom dia. Para avançarmos com o Hotel Funchal preciso que: 1. Comprar o pão para o evento; 2. Verificar todas as faturas em atraso; 3. Ligar ao cliente para confirmar a data.",
        attachments: [],
        expectedResults: {
            aiTriggers: true,
            odooMemoryFound: false,
            queueableActions: true,
        },
    },
];
