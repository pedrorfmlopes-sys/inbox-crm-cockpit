from pathlib import Path

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
        [
            ListItem(Paragraph(item, styles["GuideBody"]), leftIndent=6)
            for item in items
        ],
        bulletType="bullet",
        start="circle",
        leftIndent=16,
        bulletFontName="Helvetica",
        bulletFontSize=8,
        bulletOffsetY=1,
    )


def build_story():
    styles = build_styles()
    story = []

    story.append(Paragraph("Inbox CRM Cockpit - Guia Odoo Studio para CRM2", styles["GuideTitle"]))
    story.append(
        Paragraph(
            "Este guia explica como preparar o Odoo Studio para o modo estruturado do CRM2. "
            "O objetivo e separar descricao base, informacao fixa, historico e documentos do projeto "
            "sem perder o fallback para tenants que usam apenas a descricao standard.",
            styles["GuideIntro"],
        )
    )
    story.append(
        Paragraph(
            "Escopo desta versao: configuracao do modelo <b>project.project</b>. "
            "Os restantes modelos podem continuar a usar o fluxo tradicional enquanto o CRM2 evolui.",
            styles["GuideNote"],
        )
    )

    story.append(Paragraph("1. O que o CRM2 espera encontrar", styles["GuideSection"]))
    story.append(
        Paragraph(
            "O modo estruturado do CRM2 usa um campo base para a descricao e, opcionalmente, "
            "campos HTML separados para informacao fixa, historico e documentos. "
            "Quando os campos nao existem ou falham, a app pode regressar automaticamente ao campo de descricao.",
            styles["GuideBody"],
        )
    )
    story.append(
        build_table(
            [
                ["Funcao", "Nome tecnico recomendado", "Tipo recomendado", "Obrigatorio"],
                ["Descricao base", "description", "HTML ou Text", "Sim"],
                ["Informacao fixa", "x_studio_iccc_project_brief", "HTML", "Opcional"],
                ["Historico", "x_studio_iccc_project_history", "HTML", "Recomendado"],
                ["Documentos", "x_studio_iccc_project_documents", "HTML", "Recomendado"],
            ],
            [45 * mm, 67 * mm, 35 * mm, 25 * mm],
        )
    )

    story.append(Paragraph("2. Abas recomendadas na vista do projeto", styles["GuideSection"]))
    story.append(
        Paragraph(
            "No formulario do projeto, o CRM2 consegue validar se as abas existem. "
            "As etiquetas recomendadas sao estas:",
            styles["GuideBody"],
        )
    )
    story.append(
        build_table(
            [
                ["Aba", "Campo a colocar", "Observacao"],
                ["Informacao fixa", "x_studio_iccc_project_brief", "Contexto permanente do projeto"],
                ["Historico", "x_studio_iccc_project_history", "Emails/posts da conversa, geridos pelo CRM2"],
                ["Documentos", "x_studio_iccc_project_documents", "Blocos com anexos por conversa"],
            ],
            [40 * mm, 70 * mm, 65 * mm],
        )
    )

    story.append(Paragraph("3. Passo a passo no Odoo Studio", styles["GuideSection"]))
    story.append(
        bullet_list(
            styles,
            [
                "Abrir um registo de projeto no Odoo e clicar em <b>Studio</b>.",
                "Entrar no formulario de <b>project.project</b>.",
                "Criar os campos HTML recomendados se ainda nao existirem.",
                "Adicionar as abas <b>Informacao fixa</b>, <b>Historico</b> e <b>Documentos</b>.",
                "Colocar um campo por aba, com largura total, para evitar scroll horizontal.",
                "Guardar e publicar as alteracoes do Studio.",
                "Voltar ao add-in e abrir <b>Settings &gt; CRM2 / Odoo Layout</b>.",
                "Ativar o modo <b>Projeto com campos/abas proprias</b>.",
                "Executar <b>Validar configuracao Odoo</b> para confirmar campos, tipos e visibilidade na vista.",
            ],
        )
    )

    story.append(Paragraph("4. Como configurar no Settings do Inbox Cockpit", styles["GuideSection"]))
    story.append(
        Paragraph(
            "A secao <b>CRM2 / Odoo Layout</b> permite definir o modo de escrita e os nomes tecnicos. "
            "Se a empresa usar nomes diferentes dos recomendados, basta substituir os nomes tecnicos no cockpit "
            "e guardar as definicoes antes de validar.",
            styles["GuideBody"],
        )
    )
    story.append(
        build_table(
            [
                ["Opcao", "Quando usar"],
                ["Descricao apenas", "Tenants sem Studio ou fases piloto"],
                ["Projeto com campos/abas proprias", "Tenants com estrutura Odoo preparada"],
                ["Fallback automatico para descricao", "Recomendado para instalacoes comerciais"],
                ["Criar indice de emails/posts", "Ativa navegacao por ancora no historico"],
                ["Mostrar links Voltar ao topo", "Ajuda em historicos longos"],
            ],
            [75 * mm, 105 * mm],
        )
    )

    story.append(PageBreak())

    story.append(Paragraph("5. Comportamento esperado no CRM2", styles["GuideSection"]))
    story.append(
        bullet_list(
            styles,
            [
                "O editor <b>Descricao base</b> escreve no campo configurado como descricao principal.",
                "O editor <b>Informacao fixa</b> guarda texto permanente do projeto e nao deve ser reescrito pelo historico.",
                "Ao ligar um email ao projeto, o CRM2 pode gravar o resumo do email no campo de historico.",
                "Os anexos selecionados podem aparecer no campo de documentos, agrupados por conversa.",
                "Se algum campo falhar, o cockpit pode voltar ao campo base de descricao sem bloquear a operacao.",
            ],
        )
    )

    story.append(Paragraph("6. Recomendacoes para comercializacao multiempresa", styles["GuideSection"]))
    story.append(
        bullet_list(
            styles,
            [
                "Manter os nomes tecnicos recomendados sempre que possivel.",
                "Usar o validador do cockpit em cada novo cliente antes de ativar o modo estruturado.",
                "Nao desligar o fallback para descricao nas primeiras instalacoes.",
                "Documentar internamente quem gere o Studio em cada empresa cliente.",
                "Testar um projeto novo e um projeto existente antes de iniciar utilizacao produtiva.",
            ],
        )
    )

    story.append(Paragraph("7. Troubleshooting rapido", styles["GuideSection"]))
    story.append(
        build_table(
            [
                ["Sintoma", "Causa provavel", "Acao recomendada"],
                ["Campo nao encontrado", "Nome tecnico errado ou campo nao criado", "Corrigir nome tecnico ou criar o campo no Studio"],
                ["Campo existe mas nao aparece", "Campo fora da vista form", "Adicionar o campo a uma aba visivel"],
                ["Aba nao encontrada", "Etiqueta diferente da configurada", "Ajustar label no Studio ou no Settings"],
                ["CRM2 voltou para descricao", "Fallback ativo apos erro no campo estruturado", "Validar configuracao e rever permissao/tipo"],
            ],
            [42 * mm, 62 * mm, 76 * mm],
        )
    )

    story.append(Spacer(1, 8))
    story.append(
        Paragraph(
            "Documento gerado para acompanhar a configuracao do modo estruturado do CRM2 no Inbox CRM Cockpit.",
            styles["GuideNote"],
        )
    )

    return story


def add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#486581"))
    canvas.drawRightString(A4[0] - 18 * mm, 12 * mm, f"Pagina {doc.page}")
    canvas.restoreState()


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)

    output_pdf = OUTPUT_DIR / FILENAME
    public_pdf = PUBLIC_DIR / FILENAME

    doc = SimpleDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title="Inbox CRM Cockpit - Guia Odoo Studio para CRM2",
        author="OpenAI Codex",
        subject="Configuracao CRM2 / Odoo Studio",
    )
    doc.build(build_story(), onFirstPage=add_page_number, onLaterPages=add_page_number)

    public_pdf.write_bytes(output_pdf.read_bytes())
    print(output_pdf)
    print(public_pdf)


if __name__ == "__main__":
    main()
