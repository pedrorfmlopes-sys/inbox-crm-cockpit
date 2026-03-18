from pathlib import Path
import shutil

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import ListFlowable, ListItem, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "output" / "pdf"
PUBLIC_DIR = ROOT / "client" / "public" / "docs"
FILENAME = "inbox-cockpit-crm2-odoo-studio-setup.pdf"

LAYOUTS = [
    {
        "title": "Projetos",
        "model": "project.project",
        "description": "description",
        "fixed": "x_studio_iccc_project_brief",
        "history": "x_studio_iccc_project_history",
        "documents": "x_studio_iccc_project_documents",
    },
    {
        "title": "Leads",
        "model": "crm.lead",
        "description": "description",
        "fixed": "x_studio_iccc_lead_brief",
        "history": "x_studio_iccc_lead_history",
        "documents": "x_studio_iccc_lead_documents",
    },
    {
        "title": "Tarefas",
        "model": "project.task",
        "description": "description",
        "fixed": "x_studio_iccc_task_brief",
        "history": "x_studio_iccc_task_history",
        "documents": "x_studio_iccc_task_documents",
    },
    {
        "title": "Tickets",
        "model": "helpdesk.ticket",
        "description": "description",
        "fixed": "x_studio_iccc_ticket_brief",
        "history": "x_studio_iccc_ticket_history",
        "documents": "x_studio_iccc_ticket_documents",
    },
]


def build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="GuideTitle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=colors.HexColor("#12355B"),
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideIntro",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.5,
            leading=15,
            textColor=colors.HexColor("#334E68"),
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideSection",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13.5,
            leading=17,
            textColor=colors.HexColor("#0F4C81"),
            spaceBefore=10,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#243B53"),
            alignment=TA_LEFT,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideNote",
            parent=styles["BodyText"],
            fontName="Helvetica-Oblique",
            fontSize=9.5,
            leading=13,
            textColor=colors.HexColor("#486581"),
            spaceAfter=6,
        )
    )
    return styles


def build_table(data, col_widths):
    table = Table(data, colWidths=col_widths, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2FF")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#12355B")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LEADING", (0, 0), (-1, -1), 12),
                ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#243B53")),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FBFF")]),
                ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#BCCCDC")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9E2EC")),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return table


def bullet_list(styles, items):
    return ListFlowable(
        [ListItem(Paragraph(item, styles["GuideBody"]), leftIndent=6) for item in items],
        bulletType="bullet",
        start="circle",
        leftIndent=16,
        bulletFontName="Helvetica",
        bulletFontSize=8,
        bulletOffsetY=1,
    )


def add_layout_story(story, styles, layout):
    story.append(Paragraph(layout["title"], styles["GuideSection"]))
    story.append(
        Paragraph(
            f"O CRM2 pode usar este layout de forma independente no modelo <b>{layout['model']}</b>. "
            f"Se o modo desta entidade ficar em 'Descricao apenas', o CRM2 continua a usar apenas o campo base.",
            styles["GuideBody"],
        )
    )
    story.append(
        build_table(
            [
                ["Funcao", "Nome tecnico recomendado", "Tipo recomendado", "Obrigatorio"],
                ["Descricao base", layout["description"], "HTML ou Text", "Sim"],
                ["Informacao fixa", layout["fixed"], "HTML", "Opcional"],
                ["Historico", layout["history"], "HTML", "Opcional"],
                ["Documentos", layout["documents"], "HTML", "Opcional"],
            ],
            [42 * mm, 70 * mm, 35 * mm, 25 * mm],
        )
    )
    story.append(Spacer(1, 6))
    story.append(
        bullet_list(
            styles,
            [
                f"Abrir o Studio num registo de <b>{layout['title'][:-1] if layout['title'].endswith('s') else layout['title']}</b> e editar a vista de <b>{layout['model']}</b>.",
                f"Criar os campos <b>{layout['fixed']}</b>, <b>{layout['history']}</b> e <b>{layout['documents']}</b> se quiseres modo estruturado nesta entidade.",
                "Adicionar as abas Informacao fixa, Historico e Documentos na vista form e colocar um campo por aba.",
                "Guardar o Studio, voltar ao add-in, escolher esta entidade em Settings e correr a validacao.",
            ],
        )
    )


def build_story():
    styles = build_styles()
    story = []

    story.append(Paragraph("Inbox CRM Cockpit - Guia Odoo Studio para CRM2", styles["GuideTitle"]))
    story.append(
        Paragraph(
            "Este guia explica como preparar o Odoo Studio para o modo estruturado do CRM2. "
            "A configuracao e independente por entidade: projetos, leads, tarefas e tickets podem usar "
            "modos diferentes na mesma empresa.",
            styles["GuideIntro"],
        )
    )
    story.append(
        Paragraph(
            "O cockpit suporta dois modos por entidade: <b>Descricao apenas</b> e <b>Layout estruturado</b>. "
            "Podes ativar o modo estruturado apenas nos modelos que realmente precisam dele.",
            styles["GuideNote"],
        )
    )

    story.append(Paragraph("1. Estrategia recomendada", styles["GuideSection"]))
    story.append(
        bullet_list(
            styles,
            [
                "Comecar com 'Descricao apenas' em todas as entidades.",
                "Ativar o modo estruturado apenas onde houver ganho real de organizacao.",
                "Usar o validador do cockpit antes de ativar qualquer entidade em producao.",
                "Manter o fallback para descricao ligado nas primeiras instalacoes de cada cliente.",
            ],
        )
    )

    story.append(Paragraph("2. Resumo de modelos suportados", styles["GuideSection"]))
    story.append(
        build_table(
            [["Entidade", "Modelo", "Descricao base", "Campos extra opcionais"]]
            + [[layout["title"], layout["model"], layout["description"], "Informacao fixa / Historico / Documentos"] for layout in LAYOUTS],
            [34 * mm, 42 * mm, 40 * mm, 64 * mm],
        )
    )

    story.append(Paragraph("3. Settings do CRM2 / Odoo Layout", styles["GuideSection"]))
    story.append(
        Paragraph(
            "Nos Settings, a secao CRM2 / Odoo Layout permite escolher a entidade alvo e definir o modo dessa entidade. "
            "As opcoes de indice e 'Voltar ao topo' sao globais, mas a escolha entre 'Descricao apenas' e 'Layout estruturado' e feita separadamente em cada tipo.",
            styles["GuideBody"],
        )
    )
    story.append(
        build_table(
            [
                ["Opcao", "Escopo", "Objetivo"],
                ["Modo da entidade", "Independente por entidade", "Escolher descricao simples ou layout estruturado"],
                ["Fallback para descricao", "Independente por entidade", "Garantir continuidade se algum campo Studio falhar"],
                ["Indice de emails/posts", "Global", "Criar navegacao por ancora nos historicos"],
                ["Links Voltar ao topo", "Global", "Facilitar leitura de historicos longos"],
            ],
            [48 * mm, 48 * mm, 84 * mm],
        )
    )

    story.append(PageBreak())

    story.append(Paragraph("4. Setup Studio por entidade", styles["GuideSection"]))
    for layout in LAYOUTS:
        add_layout_story(story, styles, layout)
        story.append(Spacer(1, 4))

    story.append(PageBreak())

    story.append(Paragraph("5. Comportamento esperado no CRM2", styles["GuideSection"]))
    story.append(
        bullet_list(
            styles,
            [
                "Se a entidade estiver em 'Descricao apenas', o CRM2 escreve so no campo base e continua a funcionar como hoje.",
                "Se a entidade estiver em 'Layout estruturado', o CRM2 pode separar descricao base, informacao fixa, historico e documentos.",
                "Cada email ligado pode atualizar o historico e os documentos dessa mesma entidade sem interferir com as outras.",
                "A mesma empresa pode usar, por exemplo, projetos estruturados, leads simples, tickets estruturados e tarefas simples.",
            ],
        )
    )

    story.append(Paragraph("6. Troubleshooting rapido", styles["GuideSection"]))
    story.append(
        build_table(
            [
                ["Sintoma", "Causa provavel", "Acao recomendada"],
                ["Campo nao encontrado", "Nome tecnico errado ou campo nao criado", "Corrigir nome tecnico ou criar o campo no Studio"],
                ["Campo existe mas nao aparece na validacao", "Nao esta na vista form", "Adicionar o campo a uma aba ou pagina visivel"],
                ["Historico/documentos nao atualizam", "Entidade ainda em Descricao apenas", "Verificar o modo configurado para essa entidade"],
                ["Uma entidade funciona e outra nao", "Configuracao independente por tipo", "Validar apenas a entidade em falta e rever os campos dessa vista"],
            ],
            [48 * mm, 58 * mm, 74 * mm],
        )
    )

    story.append(
        Paragraph(
            "Recomendacao comercial: manter esta estrutura documentada por cliente e ativar o modo estruturado "
            "por fases, entidade a entidade, em vez de o ligar em todo o tenant de uma vez.",
            styles["GuideNote"],
        )
    )

    return story


def generate_pdf():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)

    output_path = OUTPUT_DIR / FILENAME
    public_path = PUBLIC_DIR / FILENAME

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=16 * mm,
        leftMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Inbox CRM Cockpit - Guia Odoo Studio para CRM2",
        author="OpenAI Codex",
    )
    doc.build(build_story())
    shutil.copyfile(output_path, public_path)
    return output_path, public_path


if __name__ == "__main__":
    output_path, public_path = generate_pdf()
    print(f"PDF gerado em: {output_path}")
    print(f"Copia publicada em: {public_path}")
