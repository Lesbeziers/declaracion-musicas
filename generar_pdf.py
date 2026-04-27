from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

W, H = landscape(A4)

AZUL = colors.HexColor("#019DF4")
AZUL_OSCURO = colors.HexColor("#0B2739")
GRIS = colors.HexColor("#F5F5F5")
GRIS_TEXTO = colors.HexColor("#555555")
NARANJA = colors.HexColor("#FF6B00")
BLANCO = colors.white

# ── estilos ──────────────────────────────────────────────────────────────────
def estilos():
    titulo_portada = ParagraphStyle("tp", fontName="Helvetica-Bold", fontSize=36,
        textColor=BLANCO, alignment=TA_CENTER, leading=44)
    subtitulo_portada = ParagraphStyle("sp", fontName="Helvetica", fontSize=18,
        textColor=BLANCO, alignment=TA_CENTER, leading=26)
    titulo_slide = ParagraphStyle("ts", fontName="Helvetica-Bold", fontSize=24,
        textColor=AZUL_OSCURO, alignment=TA_LEFT, leading=30)
    cuerpo = ParagraphStyle("c", fontName="Helvetica", fontSize=13,
        textColor=AZUL_OSCURO, leading=20)
    bullet = ParagraphStyle("b", fontName="Helvetica", fontSize=12,
        textColor=AZUL_OSCURO, leading=19, leftIndent=16, bulletIndent=0,
        spaceBefore=2)
    bullet_bold = ParagraphStyle("bb", fontName="Helvetica-Bold", fontSize=12,
        textColor=AZUL_OSCURO, leading=19, leftIndent=16)
    label = ParagraphStyle("l", fontName="Helvetica-Bold", fontSize=11,
        textColor=AZUL, alignment=TA_LEFT)
    nota = ParagraphStyle("n", fontName="Helvetica-Oblique", fontSize=10,
        textColor=GRIS_TEXTO, alignment=TA_CENTER)
    tabla_cab = ParagraphStyle("tc", fontName="Helvetica-Bold", fontSize=11,
        textColor=BLANCO, alignment=TA_CENTER)
    tabla_cel = ParagraphStyle("tcc", fontName="Helvetica", fontSize=10,
        textColor=AZUL_OSCURO, alignment=TA_LEFT, leading=14)
    return dict(tp=titulo_portada, sp=subtitulo_portada, ts=titulo_slide,
                c=cuerpo, b=bullet, bb=bullet_bold, l=label, n=nota,
                tc=tabla_cab, tcc=tabla_cel)

S = estilos()

def sp(n=0.3): return Spacer(1, n*cm)
def hr(): return HRFlowable(width="100%", thickness=1.5, color=AZUL, spaceAfter=8, spaceBefore=4)

# ── diapositivas ─────────────────────────────────────────────────────────────

def slide_portada():
    # Fondo simulado con tabla de ancho total
    fondo = Table([
        [Paragraph("Declaración de Músicas", S["tp"])],
        [sp(0.5)],
        [Paragraph("Nueva herramienta web para la elaboración de Cue Sheets SGAE", S["sp"])],
        [sp(0.4)],
        [Paragraph("Movistar+ · 2025", S["sp"])],
    ], colWidths=[W - 4*cm])
    fondo.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_OSCURO),
        ("TOPPADDING", (0,0), (-1,-1), 20),
        ("BOTTOMPADDING", (0,0), (-1,-1), 20),
        ("LEFTPADDING", (0,0), (-1,-1), 40),
        ("RIGHTPADDING", (0,0), (-1,-1), 40),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    return [sp(3), fondo]


def slide_contexto():
    elems = [
        Paragraph("¿Qué es un Cue Sheet?", S["ts"]), hr(), sp(0.3),
        Paragraph("Un <b>Cue Sheet</b> es el documento oficial que toda producción audiovisual debe entregar a la <b>SGAE</b> para declarar las músicas utilizadas.", S["c"]),
        sp(0.3),
        Paragraph("Recoge para cada pieza musical:", S["c"]), sp(0.2),
        Paragraph("• Título, autor e intérprete", S["b"]),
        Paragraph("• Timecodes de entrada y salida en pantalla", S["b"]),
        Paragraph("• Duración de uso", S["b"]),
        Paragraph("• Modalidad de uso (Ambientación, Careta, Fondo…)", S["b"]),
        Paragraph("• Tipo de música (Librería, Comercial, Original)", S["b"]),
        sp(0.4),
        Paragraph("Es un <b>requisito legal</b> para la gestión de derechos de autor en España.", S["c"]),
    ]
    return elems


def slide_intro():
    elems = [
        Paragraph("La herramienta", S["ts"]), hr(), sp(0.3),
        Paragraph("Aplicación web interna accesible desde el navegador, sin instalación.", S["c"]),
        sp(0.3),
        Paragraph("• <b>Sin instalación</b> — funciona en cualquier navegador moderno", S["b"]),
        Paragraph("• <b>Sin servidores</b> — los datos no salen del equipo del usuario", S["b"]),
        Paragraph("• <b>Siempre actualizada</b> — todos los usuarios trabajan con la misma versión", S["b"]),
        Paragraph("• <b>Estilo corporativo</b> Movistar+", S["b"]),
        sp(0.5),
        Paragraph("Permite crear, editar, importar y exportar Cue Sheets en el formato oficial de SGAE.", S["c"]),
    ]
    return elems


def slide_funcionalidades():
    elems = [
        Paragraph("Funcionalidades principales", S["ts"]), hr(), sp(0.2),
    ]
    items = [
        ("1", "Editor tabular", "Tabla editable de hasta 250 filas con las 10 columnas de la declaración oficial."),
        ("2", "Timecodes y duración automática", "Calcula la duración al introducir TC In y TC Out. Normaliza cualquier formato de timecode."),
        ("3", "Validación antes de exportar", "Marca en naranja las celdas incompletas e impide exportar hasta corregirlas."),
        ("4", "Importación desde Excel", "Carga ficheros .xlsx/.xls existentes con detección automática de columnas."),
        ("5", "Exportación al formato SGAE", "Genera el Excel con la plantilla oficial v3 NUBE, lista para entregar."),
        ("6", "Historial deshacer / rehacer", "Ctrl+Z / Ctrl+Y con hasta 200 acciones. Edición avanzada: copiar, pegar, ordenar."),
    ]
    data = [[
        Paragraph(i, ParagraphStyle("n", fontName="Helvetica-Bold", fontSize=18, textColor=AZUL, alignment=TA_CENTER)),
        Paragraph(f"<b>{t}</b>", S["tcc"]),
        Paragraph(d, S["tcc"]),
    ] for i, t, d in items]
    t = Table(data, colWidths=[1.2*cm, 6.5*cm, 18*cm])
    t.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [GRIS, BLANCO]),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#DDDDDD")),
    ]))
    elems.append(t)
    return elems


def slide_editor():
    elems = [
        Paragraph("Editor tabular — Columnas de la declaración", S["ts"]), hr(), sp(0.2),
    ]
    cab = [Paragraph(c, S["tc"]) for c in ["Campo", "Tipo", "Obligatorio"]]
    filas = [
        ["Título", "Texto libre", "✔"],
        ["Autor", "Texto libre", "✔"],
        ["Intérprete", "Texto libre", "—"],
        ["TC In", "Timecode HH:MM:SS", "—"],
        ["TC Out", "Timecode HH:MM:SS", "—"],
        ["Duración", "Calculada o manual", "✔"],
        ["Modalidad", "Desplegable (5 opciones)", "✔"],
        ["Tipo de música", "Desplegable (3 opciones)", "✔"],
        ["Código librería", "Texto libre", "—"],
        ["Nombre librería", "Texto libre", "—"],
    ]
    data = [cab] + [[Paragraph(c, S["tcc"]) for c in f] for f in filas]
    t = Table(data, colWidths=[8*cm, 8*cm, 4*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), AZUL_OSCURO),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [GRIS, BLANCO]),
        ("ALIGN", (2,0), (2,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#DDDDDD")),
        ("FONTNAME", (2,1), (2,-1), "Helvetica-Bold"),
        ("TEXTCOLOR", (2,1), (2,-1), AZUL),
    ]))
    elems.append(t)
    elems += [sp(0.3), Paragraph("Los desplegables garantizan que sólo se usan valores reconocidos por SGAE.", S["n"])]
    return elems


def slide_validacion():
    elems = [
        Paragraph("Validación antes de exportar", S["ts"]), hr(), sp(0.3),
        Paragraph("Antes de generar el fichero, la herramienta comprueba que la declaración está completa:", S["c"]),
        sp(0.3),
        Paragraph("• Detecta celdas vacías en campos obligatorios", S["b"]),
        Paragraph("• Verifica el formato de los timecodes", S["b"]),
        Paragraph("• Comprueba que el título del programa esté informado", S["b"]),
        Paragraph("• Marca en <font color='#FF6B00'><b>naranja</b></font> cada celda con error y la etiqueta OBLIGATORIO", S["b"]),
        Paragraph("• <b>Bloquea la exportación</b> hasta que se corrijan todos los errores", S["b"]),
        sp(0.5),
        Paragraph("Resultado: ningún Cue Sheet incompleto llega a SGAE.", S["c"]),
    ]
    return elems


def slide_import_export():
    elems = [
        Paragraph("Importación y exportación de Excel", S["ts"]), hr(), sp(0.2),
    ]
    data = [
        [Paragraph("IMPORTACIÓN", S["tc"]), Paragraph("EXPORTACIÓN", S["tc"])],
        [
            Paragraph(
                "• Acepta .xlsx, .xls, .xlsm\n"
                "• Detecta automáticamente la fila de encabezados\n"
                "• Mapea columnas por nombre, aunque varíe la denominación\n"
                "  (ej: 'compositor', 'writer' → Autor)\n"
                "• Calcula duraciones si hay TC In y TC Out\n"
                "• Compatible con ficheros existentes", S["tcc"]),
            Paragraph(
                "• Genera la plantilla oficial v3 NUBE de SGAE\n"
                "• Nombre de fichero automático basado en el título\n"
                "• Elimina filas vacías para que NUBE no las procese\n"
                "• No requiere tener la plantilla guardada en el equipo\n"
                "• Descarga directa desde el navegador", S["tcc"]),
        ]
    ]
    t = Table(data, colWidths=[13.5*cm, 13.5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,0), AZUL),
        ("BACKGROUND", (1,0), (1,0), AZUL_OSCURO),
        ("BACKGROUND", (0,1), (0,1), colors.HexColor("#EAF6FE")),
        ("BACKGROUND", (1,1), (1,1), GRIS),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 12),
        ("BOTTOMPADDING", (0,0), (-1,-1), 12),
        ("LEFTPADDING", (0,0), (-1,-1), 14),
        ("RIGHTPADDING", (0,0), (-1,-1), 14),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC")),
    ]))
    elems.append(t)
    return elems


def slide_ventajas():
    elems = [
        Paragraph("Ventajas frente a la situación anterior", S["ts"]), hr(), sp(0.2),
    ]
    cab = [Paragraph(c, S["tc"]) for c in ["Aspecto", "Antes (Excels con macros)", "Ahora (herramienta web)"]]
    filas = [
        ["Acceso", "Macros bloqueadas por seguridad corporativa", "Funciona en cualquier navegador, sin macros"],
        ["Validación", "Sin validación — errores pasan desapercibidos", "Bloquea exportación si hay campos incompletos"],
        ["Duración", "Fórmula manual o cálculo a mano", "Cálculo automático al introducir timecodes"],
        ["Plantilla", "Cada usuario gestiona su propia copia", "Plantilla oficial siempre actualizada y centralizada"],
        ["Consistencia", "Versiones distintas entre departamentos", "Todos usan la misma versión en todo momento"],
        ["Historial", "Deshacer limitado de Excel", "Hasta 200 acciones deshacer/rehacer"],
        ["Datos", "Sin restricción en valores de campo", "Desplegables con valores oficiales SGAE"],
    ]
    data = [cab] + [[Paragraph(c, S["tcc"]) for c in f] for f in filas]
    colW = [6*cm, 10.5*cm, 10.5*cm]
    t = Table(data, colWidths=colW)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), AZUL_OSCURO),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [GRIS, BLANCO]),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#DDDDDD")),
        ("FONTNAME", (0,1), (0,-1), "Helvetica-Bold"),
        ("TEXTCOLOR", (1,1), (1,-1), colors.HexColor("#AA0000")),
        ("TEXTCOLOR", (2,1), (2,-1), colors.HexColor("#006600")),
    ]))
    elems.append(t)
    return elems


def slide_cierre():
    fondo = Table([
        [Paragraph("Declaración de Músicas", S["tp"])],
        [sp(0.4)],
        [Paragraph("Una herramienta sencilla, validada y siempre actualizada\npara cumplir con los requisitos de SGAE.", S["sp"])],
        [sp(0.6)],
        [Paragraph("Sin instalación · Sin macros · Sin errores en la entrega", ParagraphStyle(
            "tag", fontName="Helvetica-Bold", fontSize=14,
            textColor=AZUL, alignment=TA_CENTER))],
    ], colWidths=[W - 4*cm])
    fondo.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_OSCURO),
        ("TOPPADDING", (0,0), (-1,-1), 20),
        ("BOTTOMPADDING", (0,0), (-1,-1), 20),
        ("LEFTPADDING", (0,0), (-1,-1), 40),
        ("RIGHTPADDING", (0,0), (-1,-1), 40),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))
    return [sp(3), fondo]


# ── construcción del documento ────────────────────────────────────────────────

SLIDES = [
    slide_portada,
    slide_contexto,
    slide_intro,
    slide_funcionalidades,
    slide_editor,
    slide_validacion,
    slide_import_export,
    slide_ventajas,
    slide_cierre,
]

from reportlab.platypus import PageBreak

out = "/home/user/declaracion-musicas/Declaracion_Musicas_Presentacion.pdf"
doc = SimpleDocTemplate(out, pagesize=landscape(A4),
    leftMargin=2*cm, rightMargin=2*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)

story = []
for i, slide_fn in enumerate(SLIDES):
    story += slide_fn()
    if i < len(SLIDES) - 1:
        story.append(PageBreak())

doc.build(story)
print("PDF generado:", out)
