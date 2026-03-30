"""
PPT Generator — Backend Flask
Gera apresentações PowerPoint formatadas com identidade visual da empresa
usando Amazon Bedrock para estruturar o conteúdo.
"""
import os
import json
import uuid
import boto3
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import tempfile
import io

app = Flask(__name__, template_folder="../frontend", static_folder="../frontend/static")
CORS(app)

# ─────────────────────────────────────────────
# Configuração
# ─────────────────────────────────────────────
REGION         = os.environ.get("AWS_REGION", "us-east-1")
S3_BUCKET      = os.environ.get("S3_BUCKET", "ppt-generator-templates")
BEDROCK_MODEL  = "anthropic.claude-3-5-sonnet-20241022-v2:0"

s3      = boto3.client("s3",              region_name=REGION)
bedrock = boto3.client("bedrock-runtime", region_name=REGION)

UPLOAD_FOLDER  = "/tmp/uploads"
ALLOWED_EXTS   = {"pptx", "pdf", "docx", "txt", "md"}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ─────────────────────────────────────────────
# Rotas
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload-template", methods=["POST"])
def upload_template():
    """Faz upload do template da empresa para o S3."""
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Nome de arquivo inválido"}), 400

    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ALLOWED_EXTS:
        return jsonify({"error": f"Formato não suportado. Use: {', '.join(ALLOWED_EXTS)}"}), 400

    filename  = f"templates/{uuid.uuid4()}_{secure_filename(file.filename)}"
    file_data = file.read()

    s3.put_object(Bucket=S3_BUCKET, Key=filename, Body=file_data)

    return jsonify({
        "success":  True,
        "template_key": filename,
        "message":  f"Template '{file.filename}' enviado com sucesso"
    })


@app.route("/api/list-templates", methods=["GET"])
def list_templates():
    """Lista templates disponíveis no S3."""
    resp    = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="templates/")
    objects = resp.get("Contents", [])

    templates = []
    for obj in objects:
        key  = obj["Key"]
        name = key.split("/")[-1].split("_", 1)[-1]  # remove UUID prefix
        templates.append({
            "key":          key,
            "name":         name,
            "size_kb":      round(obj["Size"] / 1024, 1),
            "last_modified": str(obj["LastModified"])
        })

    return jsonify({"templates": templates})


@app.route("/api/generate", methods=["POST"])
def generate_presentation():
    """
    Gera uma apresentação PowerPoint preenchendo o template da empresa.
    Fluxo:
      1. Baixa o template PPTX do S3
      2. Usa Bedrock para estruturar o conteúdo em slides
      3. Preenche os slides do template com o conteúdo gerado
      4. Retorna URL de download
    """
    data = request.json
    if not data:
        return jsonify({"error": "Dados inválidos"}), 400

    doc_type     = data.get("doc_type", "proposta")
    briefing     = data.get("briefing", "")
    company_name = data.get("company_name", "Empresa")
    client_name  = data.get("client_name", "Cliente")
    template_key = data.get("template_key")
    num_slides   = data.get("num_slides", 10)

    if not briefing:
        return jsonify({"error": "Briefing é obrigatório"}), 400

    if not template_key:
        return jsonify({"error": "Selecione um template PPTX da empresa antes de gerar"}), 400

    # 1. Baixa o template PPTX do S3
    try:
        obj      = s3.get_object(Bucket=S3_BUCKET, Key=template_key)
        tpl_data = obj["Body"].read()
    except Exception as e:
        return jsonify({"error": f"Erro ao baixar template: {str(e)}"}), 500

    # 2. Abre o template e conta quantos slides tem
    tpl_buffer = io.BytesIO(tpl_data)
    prs        = Presentation(tpl_buffer)
    n_slides   = len(prs.slides)

    if n_slides == 0:
        return jsonify({"error": "O template não tem slides"}), 400

    # 3. Usa Bedrock para gerar conteúdo para cada slide do template
    slides_content = _generate_content_with_bedrock(
        doc_type, briefing, company_name, client_name, n_slides
    )

    # 4. Preenche os slides do template com o conteúdo
    pptx_buffer = _fill_template_slides(prs, slides_content, company_name, client_name)

    # 5. Salva no S3 e retorna URL
    output_key = f"outputs/{uuid.uuid4()}_{doc_type}_{client_name}.pptx"
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=output_key,
        Body=pptx_buffer.getvalue(),
        ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    download_url = s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": output_key},
        ExpiresIn=3600
    )

    return jsonify({
        "success":      True,
        "download_url": download_url,
        "slides_count": len(slides_content),
        "template_slides": n_slides,
        "output_key":   output_key
    })


# ─────────────────────────────────────────────
# Geração de conteúdo com Bedrock
# ─────────────────────────────────────────────

def _generate_content_with_bedrock(doc_type, briefing, company, client, num_slides):
    """Usa Claude para estruturar o conteúdo em slides."""

    type_instructions = {
        "proposta":    "Proposta comercial profissional com problema, solução, benefícios, investimento e próximos passos",
        "apresentacao": "Apresentação executiva com contexto, análise, recomendações e conclusão",
        "relatorio":   "Relatório de resultados com sumário executivo, métricas, análise e recomendações",
        "workshop":    "Material de workshop com agenda, conceitos, exercícios práticos e takeaways",
    }

    instruction = type_instructions.get(doc_type, type_instructions["proposta"])

    prompt = f"""Você é um especialista em criação de apresentações executivas.
Crie uma estrutura de {num_slides} slides para: {instruction}

Empresa apresentando: {company}
Cliente/Audiência: {client}
Briefing/Conteúdo: {briefing}

Retorne APENAS um JSON válido com esta estrutura exata:
{{
  "title": "Título principal da apresentação",
  "subtitle": "Subtítulo",
  "slides": [
    {{
      "type": "cover|section|content|bullets|metrics|closing",
      "title": "Título do slide",
      "subtitle": "Subtítulo opcional",
      "content": "Texto principal do slide (máximo 150 palavras)",
      "bullets": ["ponto 1", "ponto 2", "ponto 3"],
      "metrics": [{{"label": "Métrica", "value": "Valor", "unit": "unidade"}}],
      "notes": "Notas do apresentador"
    }}
  ]
}}

Tipos de slide:
- cover: capa (apenas 1, primeiro slide)
- section: divisor de seção (título grande, sem conteúdo)
- content: slide de conteúdo com texto
- bullets: slide com lista de pontos
- metrics: slide com números/métricas em destaque
- closing: encerramento com CTA (apenas 1, último slide)

Seja profissional, objetivo e persuasivo."""

    response = bedrock.invoke_model(
        modelId=BEDROCK_MODEL,
        body=json.dumps({
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": 4096,
            "messages": [{"role": "user", "content": prompt}]
        })
    )

    result = json.loads(response["body"].read())
    text   = result["content"][0]["text"].strip()

    # Remove markdown se vier com ```json
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    text = text.strip().rstrip("```")

    data = json.loads(text)
    return data.get("slides", [])


def _fill_template_slides(prs, slides_content, company, client):
    """
    Preenche os slides do template PPTX com o conteúdo gerado pelo Bedrock.
    Preserva 100% do design original: cores, fontes, logos, layouts.
    Substitui apenas o texto nos placeholders.
    """
    from pptx.util import Pt
    from pptx.enum.shapes import PP_PLACEHOLDER

    for i, slide in enumerate(prs.slides):
        if i >= len(slides_content):
            break

        content = slides_content[i]
        title   = content.get("title", "")
        subtitle = content.get("subtitle", "")
        body    = content.get("content", "")
        bullets = content.get("bullets", [])
        metrics = content.get("metrics", [])

        # Monta o texto do corpo
        if bullets:
            body_text = "\n".join(f"• {b}" for b in bullets)
        elif metrics:
            body_text = "\n".join(f"{m.get('label','')}: {m.get('value','')} {m.get('unit','')}".strip() for m in metrics)
        else:
            body_text = body

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            ph_type = None
            if shape.is_placeholder:
                ph_type = shape.placeholder_format.type

            # Detecta o tipo de placeholder pelo tipo ou pelo nome
            name_lower = shape.name.lower()

            is_title    = ph_type in (1, 13) or "title" in name_lower
            is_subtitle = ph_type in (2, 12) or "subtitle" in name_lower or "subtitulo" in name_lower
            is_body     = ph_type in (2, 7, 15) or "content" in name_lower or "body" in name_lower or "texto" in name_lower or "conteudo" in name_lower

            if is_title and title:
                _replace_text_preserving_format(shape.text_frame, title)

            elif is_subtitle and subtitle:
                _replace_text_preserving_format(shape.text_frame, subtitle)

            elif is_body and body_text:
                _replace_text_preserving_format(shape.text_frame, body_text)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


def _replace_text_preserving_format(text_frame, new_text):
    """
    Substitui o texto de um text_frame preservando a formatação original
    (fonte, tamanho, cor, negrito, alinhamento).
    """
    if not text_frame.paragraphs:
        return

    # Captura a formatação do primeiro run do primeiro parágrafo
    first_para = text_frame.paragraphs[0]
    ref_run    = first_para.runs[0] if first_para.runs else None

    # Limpa todos os parágrafos exceto o primeiro
    while len(text_frame.paragraphs) > 1:
        p = text_frame.paragraphs[-1]._p
        p.getparent().remove(p)

    # Divide o novo texto em linhas
    lines = new_text.split("\n")

    # Preenche o primeiro parágrafo
    para = text_frame.paragraphs[0]
    _set_para_text(para, lines[0], ref_run)

    # Adiciona parágrafos para as linhas restantes
    from pptx.oxml.ns import qn
    from copy import deepcopy
    for line in lines[1:]:
        # Clona o parágrafo de referência para manter formatação
        new_p = deepcopy(para._p)
        para._p.addnext(new_p)
        new_para = text_frame.paragraphs[-1]
        _set_para_text(new_para, line, ref_run)


def _set_para_text(para, text, ref_run=None):
    """Define o texto de um parágrafo preservando a formatação do run de referência."""
    from pptx.oxml.ns import qn

    # Remove runs existentes
    for r in para.runs:
        r._r.getparent().remove(r._r)

    # Adiciona novo run com o texto
    from pptx.oxml import parse_xml
    from lxml import etree

    if ref_run:
        # Clona o run de referência para manter formatação
        from copy import deepcopy
        new_r = deepcopy(ref_run._r)
        # Atualiza o texto
        t_elem = new_r.find(qn('a:t'))
        if t_elem is None:
            t_elem = etree.SubElement(new_r, qn('a:t'))
        t_elem.text = text
        para._p.append(new_r)
    else:
        # Cria run simples
        run = para.add_run()
        run.text = text


# ─────────────────────────────────────────────
# Funções de construção do PPTX (fallback sem template)
# ─────────────────────────────────────────────

def _hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )

def _build_pptx(slides_content, company, client, primary_hex, accent_hex, doc_type):
    """Constrói o arquivo PPTX com identidade visual."""
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    primary = _hex_to_rgb(primary_hex)
    accent  = _hex_to_rgb(accent_hex)
    white   = RGBColor(255, 255, 255)
    gray    = RGBColor(100, 100, 100)
    light   = RGBColor(245, 245, 245)

    blank_layout = prs.slide_layouts[6]  # layout em branco

    for i, slide_data in enumerate(slides_content):
        slide_type = slide_data.get("type", "content")
        slide      = prs.slides.add_slide(blank_layout)

        if slide_type == "cover":
            _build_cover_slide(slide, slide_data, company, client, primary, accent, white)
        elif slide_type == "section":
            _build_section_slide(slide, slide_data, primary, accent, white)
        elif slide_type == "metrics":
            _build_metrics_slide(slide, slide_data, primary, accent, white, gray)
        elif slide_type == "closing":
            _build_closing_slide(slide, slide_data, company, primary, accent, white)
        else:
            _build_content_slide(slide, slide_data, primary, accent, white, gray, light)

        # Número de slide (exceto capa)
        if slide_type != "cover":
            _add_slide_number(slide, i + 1, len(slides_content), gray)

        # Rodapé com nome da empresa
        _add_footer(slide, company, gray)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


def _add_shape_rect(slide, left, top, width, height, color, transparency=0):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(1, left, top, width, height)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_text(slide, text, left, top, width, height, font_size, color, bold=False, align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf    = txBox.text_frame
    tf.word_wrap = True
    p  = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size  = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold  = bold
    run.font.italic = italic
    return txBox


def _build_cover_slide(slide, data, company, client, primary, accent, white):
    W, H = Inches(13.33), Inches(7.5)

    # Fundo primário
    _add_shape_rect(slide, 0, 0, W, H, primary)

    # Barra de destaque lateral
    _add_shape_rect(slide, 0, 0, Inches(0.5), H, accent)

    # Linha decorativa
    _add_shape_rect(slide, Inches(0.8), Inches(3.2), Inches(8), Inches(0.04), accent)

    # Título
    _add_text(slide, data.get("title", "Apresentação"),
              Inches(0.9), Inches(1.5), Inches(10), Inches(1.5),
              40, white, bold=True, align=PP_ALIGN.LEFT)

    # Subtítulo
    _add_text(slide, data.get("subtitle", ""),
              Inches(0.9), Inches(3.4), Inches(9), Inches(0.8),
              20, RGBColor(200, 210, 220), align=PP_ALIGN.LEFT)

    # Para / Cliente
    _add_text(slide, f"Para: {client}",
              Inches(0.9), Inches(5.2), Inches(6), Inches(0.5),
              14, RGBColor(180, 190, 200), align=PP_ALIGN.LEFT)

    # Empresa
    _add_text(slide, company,
              Inches(0.9), Inches(5.8), Inches(6), Inches(0.5),
              14, accent, bold=True, align=PP_ALIGN.LEFT)


def _build_section_slide(slide, data, primary, accent, white):
    W, H = Inches(13.33), Inches(7.5)

    _add_shape_rect(slide, 0, 0, W, H, primary)
    _add_shape_rect(slide, 0, Inches(3.2), W, Inches(0.06), accent)

    _add_text(slide, data.get("title", ""),
              Inches(1), Inches(2.5), Inches(11), Inches(1.5),
              44, white, bold=True, align=PP_ALIGN.CENTER)

    if data.get("subtitle"):
        _add_text(slide, data["subtitle"],
                  Inches(1), Inches(4.2), Inches(11), Inches(0.8),
                  20, RGBColor(180, 190, 200), align=PP_ALIGN.CENTER)


def _build_content_slide(slide, data, primary, accent, white, gray, light):
    W, H = Inches(13.33), Inches(7.5)

    # Fundo branco
    _add_shape_rect(slide, 0, 0, W, H, white)

    # Barra superior
    _add_shape_rect(slide, 0, 0, W, Inches(1.1), primary)

    # Linha accent
    _add_shape_rect(slide, 0, Inches(1.1), Inches(3), Inches(0.05), accent)

    # Título
    _add_text(slide, data.get("title", ""),
              Inches(0.4), Inches(0.15), Inches(12), Inches(0.8),
              24, white, bold=True)

    # Conteúdo
    content = data.get("content", "")
    bullets = data.get("bullets", [])

    if bullets:
        for j, bullet in enumerate(bullets[:6]):
            y = Inches(1.4) + j * Inches(0.85)
            # Marcador
            _add_shape_rect(slide, Inches(0.5), y + Inches(0.15), Inches(0.12), Inches(0.12), accent)
            _add_text(slide, bullet, Inches(0.8), y, Inches(11.5), Inches(0.75), 16, gray)
    elif content:
        _add_text(slide, content, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.5), 16, gray)


def _build_metrics_slide(slide, data, primary, accent, white, gray):
    W, H = Inches(13.33), Inches(7.5)

    _add_shape_rect(slide, 0, 0, W, H, white)
    _add_shape_rect(slide, 0, 0, W, Inches(1.1), primary)
    _add_shape_rect(slide, 0, Inches(1.1), Inches(3), Inches(0.05), accent)

    _add_text(slide, data.get("title", ""),
              Inches(0.4), Inches(0.15), Inches(12), Inches(0.8),
              24, white, bold=True)

    metrics = data.get("metrics", [])
    if not metrics and data.get("bullets"):
        # Converte bullets em métricas se não tiver métricas definidas
        metrics = [{"label": b, "value": "", "unit": ""} for b in data["bullets"][:4]]

    cols = min(len(metrics), 4)
    if cols == 0:
        return

    card_w = Inches(12) / cols
    for j, metric in enumerate(metrics[:4]):
        x = Inches(0.5) + j * card_w

        # Card
        _add_shape_rect(slide, x, Inches(1.5), card_w - Inches(0.2), Inches(4.5),
                        RGBColor(245, 247, 250))

        # Valor
        _add_text(slide, metric.get("value", ""),
                  x + Inches(0.1), Inches(2.0), card_w - Inches(0.4), Inches(1.5),
                  42, accent, bold=True, align=PP_ALIGN.CENTER)

        # Unidade
        if metric.get("unit"):
            _add_text(slide, metric["unit"],
                      x + Inches(0.1), Inches(3.3), card_w - Inches(0.4), Inches(0.5),
                      14, gray, align=PP_ALIGN.CENTER)

        # Label
        _add_text(slide, metric.get("label", ""),
                  x + Inches(0.1), Inches(4.0), card_w - Inches(0.4), Inches(0.8),
                  14, RGBColor(80, 80, 80), align=PP_ALIGN.CENTER)


def _build_closing_slide(slide, data, company, primary, accent, white):
    W, H = Inches(13.33), Inches(7.5)

    _add_shape_rect(slide, 0, 0, W, H, primary)
    _add_shape_rect(slide, 0, 0, Inches(0.5), H, accent)
    _add_shape_rect(slide, Inches(0.8), Inches(3.8), Inches(8), Inches(0.04), accent)

    _add_text(slide, data.get("title", "Próximos Passos"),
              Inches(0.9), Inches(1.5), Inches(11), Inches(1.2),
              36, white, bold=True)

    content = data.get("content", "")
    if content:
        _add_text(slide, content,
                  Inches(0.9), Inches(3.0), Inches(10), Inches(1.5),
                  18, RGBColor(200, 210, 220))

    _add_text(slide, company,
              Inches(0.9), Inches(5.5), Inches(8), Inches(0.6),
              16, accent, bold=True)


def _add_slide_number(slide, current, total, gray):
    _add_text(slide, f"{current} / {total}",
              Inches(12.0), Inches(7.0), Inches(1.2), Inches(0.4),
              10, gray, align=PP_ALIGN.RIGHT)


def _add_footer(slide, company, gray):
    _add_text(slide, company,
              Inches(0.3), Inches(7.1), Inches(5), Inches(0.35),
              9, gray)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)

