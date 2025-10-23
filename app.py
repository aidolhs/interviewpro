# ====================================================================================
# app.py — 스트리밍 생성 + (Markdown → PDF/DOCX) 안정 변환 (prompt.txt 외부 프롬프트)
#  - PDF: 블록/인라인 수식 모두 라텍스 → 이미지로 삽입 (matplotlib mathtext)
#  - PDF: 그리스문자/수학기호/위·아래첨자 폴백(DejaVuSans) + 비율 보존
#  - DOCX: 제목 직후/쓰레기 블릿 제거, 빈 문단 제거, 인라인/블록 수식 이미지 삽입
#  - 프롬프트: 동일 폴더의 prompt.txt에서 읽어 {text_from_pdf}에 PDF추출 텍스트 삽입
# ====================================================================================

import os, io, re, pdfplumber
from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
from dotenv import load_dotenv
import google.generativeai as genai

# ===== ReportLab (PDF)
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm

# ===== python-docx (DOCX)
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

# ===== matplotlib for LaTeX
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
matplotlib.rcParams['mathtext.fontset'] = 'dejavusans'  # 수식 폰트셋 고정

# ------------------------------------------------------------------------------------
# 앱/모델 설정
# ------------------------------------------------------------------------------------
load_dotenv()
app = Flask(__name__)
CORS(app)

try:
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
    model = genai.GenerativeModel('gemini-2.5-flash')
except Exception as e:
    print(f"[WARN] Gemini 설정 실패: {e}")
    model = None

# ------------------------------------------------------------------------------------
# 유틸
# ------------------------------------------------------------------------------------
_ctrl_pattern = re.compile(r'[\x00-\x1F\x7F-\x9F]')

def sanitize_text(t: str) -> str:
    """제어문자/눈에 보이는 개행 화살표 제거"""
    return _ctrl_pattern.sub('', t).replace('↩', '').replace('\u2028', '\n')

def escape_html(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# 인라인 수식/블록 수식 탐지
_inline_math_re = re.compile(r'\$(.+?)\$')
_block_math_single_re = re.compile(r'^\s*\$\$(.+?)\$\$\s*$')
_block_math_open_re   = re.compile(r'^\s*\$\$(.+)$')
_block_math_close_re  = re.compile(r'^(.+)\$\$\s*$')

def render_latex_png(latex: str, dpi: int = 230, pad: float = 0.12) -> bytes:
    """LaTeX(수학 모드) → 투명 PNG"""
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.patch.set_alpha(0.0)
    ax = fig.add_subplot(111)
    ax.axis("off")
    ax.text(0.5, 0.5, f"${latex}$", fontsize=15, ha="center", va="center")
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=dpi, transparent=True,
                bbox_inches="tight", pad_inches=pad)
    plt.close(fig)
    buf.seek(0)
    return buf.read()

# 인라인 마크다운 → ReportLab 문단 태그(굵게/기울임/코드, 폴백 폰트 래핑)
_inline_code_re   = re.compile(r'`([^`]+)`')
_inline_bold_re   = re.compile(r'\*\*(.+?)\*\*|__(.+?)__')
_inline_italic_re = re.compile(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|_(.+?)_')
# Greek/위·아래첨자/수학기호 폴백 범위
_fallback_re = re.compile(r'([\u0370-\u03FF\u1F00-\u1FFF\u2070-\u209F\u2200-\u22FF]+)')

def md_inline_to_rl(text: str, fallback_font: str = 'DejaVuSans') -> str:
    code_spans = []
    def _save(m):
        code_spans.append(m.group(1))
        return f"\u0000CODE{len(code_spans)-1}\u0000"
    text = _inline_code_re.sub(_save, text)
    t = escape_html(text)
    t = _fallback_re.sub(lambda m: f"<font name='{fallback_font}'>{m.group(1)}</font>", t)
    t = _inline_bold_re.sub(lambda m: f"<b>{m.group(1) or m.group(2) or ''}</b>", t)
    t = _inline_italic_re.sub(lambda m: f"<i>{m.group(1) or m.group(2) or ''}</i>", t)
    t = re.sub(r'\u0000CODE(\d+)\u0000',
               lambda m: f"<font name='Courier'>{escape_html(code_spans[int(m.group(1))])}</font>", t)
    return t

def is_trash_line(s: str) -> bool:
    """빈 줄/빈 블릿/쓰레기 기호 줄 제거"""
    st = s.strip()
    if not st: return True
    if st in ('•', '-', '*', '+'): return True
    if re.fullmatch(r'[•\-\*\+]+', st): return True
    if re.fullmatch(r'(\-|\*|\+)\s*', st): return True
    return False

# ------------------------------------------------------------------------------------
# prompt.txt 읽기
# ------------------------------------------------------------------------------------
def load_prompt_template() -> str:
    """
    같은 폴더의 prompt.txt를 UTF-8로 읽음.
    - {text_from_pdf} 토큰이 있으면 그대로 치환
    - 없으면 파일 끝에 <학생부 내용>\n{text_from_pdf} 를 덧붙여 사용
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    prompt_path = os.path.join(script_dir, 'prompt.txt')
    try:
        with open(prompt_path, 'r', encoding='utf-8') as f:
            tpl = f.read()
            if '{text_from_pdf}' not in tpl:
                tpl = tpl.rstrip() + "\n\n<학생부 내용>\n{text_from_pdf}\n"
            return tpl
    except Exception as e:
        # 안전한 최소 프롬프트로 폴백
        print(f"[WARN] prompt.txt 읽기 실패: {e}")
        return ("다음 학생부 내용을 분석하여 맞춤형 면접 질문과 모범답변을 "
                "마크다운으로 생성하세요.\n\n<학생부 내용>\n{text_from_pdf}\n")

def build_prompt(extracted_text: str) -> str:
    tpl = load_prompt_template()
    # .format()은 중괄호가 많은 프롬프트에서 KeyError가 날 수 있으므로 안전하게 replace 사용
    return tpl.replace('{text_from_pdf}', extracted_text)

# ------------------------------------------------------------------------------------
# 생성 API (스트리밍)
# ------------------------------------------------------------------------------------
@app.route("/generate", methods=["POST"])
def generate_interview_questions_stream():
    if model is None:
        return Response("AI 모델이 초기화되지 않았습니다.", status=500, mimetype="text/plain")
    if "file" not in request.files:
        return Response("요청에 파일이 없습니다.", status=400, mimetype="text/plain")
    file = request.files["file"]
    if file.filename == "":
        return Response("파일이 선택되지 않았습니다.", status=400, mimetype="text/plain")

    try:
        with pdfplumber.open(file.stream) as pdf:
            extracted_text = "".join(page.extract_text() or "" for page in pdf.pages)
        if len(extracted_text.strip()) < 50:
            return Response("PDF에서 충분한 텍스트를 추출하지 못했습니다.", status=400, mimetype="text/plain")

        prompt = build_prompt(extracted_text)

        def stream():
            try:
                for chunk in model.generate_content(prompt, stream=True):
                    yield chunk.text or ""
            except Exception as e:
                yield f"\n[스트리밍 오류] {e}"
        return Response(stream(), mimetype="text/plain; charset=utf-8")
    except Exception as e:
        return Response(f"파일 처리 오류: {e}", status=500, mimetype="text/plain")

# ------------------------------------------------------------------------------------
# 다운로드 API
# ------------------------------------------------------------------------------------
@app.route("/download", methods=["POST"])
def download_file():
    fmt = request.args.get("format") or request.form.get("format")
    content = request.form.get("content")
    if not content:
        return jsonify({"error": "전송된 내용이 없습니다. (content 누락)"}), 400

    content = sanitize_text(content)
    if fmt == "pdf":
        buf, mime, name = create_pdf_from_markdown(content)
    elif fmt in ("docx", "doc"):
        buf, mime, name = create_docx_from_markdown(content)
    else:
        return jsonify({"error": "지원하지 않는 형식입니다. (pdf, docx)"}), 400

    return send_file(buf, as_attachment=True, download_name=name, mimetype=mime)

# ------------------------------------------------------------------------------------
# PDF 변환 (제목/목록/코드 + 블록/인라인 수식 이미지 삽입, 비율보존)
# ------------------------------------------------------------------------------------
def _append_text_with_inline_math_pdf(flow, text, style, fallback_font):
    pos = 0
    for m in _inline_math_re.finditer(text):
        left = text[pos:m.start()]
        if left.strip():
            flow.append(Paragraph(md_inline_to_rl(left, fallback_font), style))
        latex = m.group(1).strip()
        img_bytes = render_latex_png(latex)
        flow.append(Image(io.BytesIO(img_bytes), width=120, preserveAspectRatio=True))
        pos = m.end()
    tail = text[pos:]
    if tail.strip():
        flow.append(Paragraph(md_inline_to_rl(tail, fallback_font), style))

def create_pdf_from_markdown(md_text: str):
    buf = io.BytesIO()
    try:
        # 폰트 등록
        script_dir = os.path.dirname(os.path.abspath(__file__))
        nanum = os.path.join(script_dir, 'NanumGothic.ttf')
        dejavu = os.path.join(script_dir, 'DejaVuSans.ttf')
        if not os.path.exists(nanum):
            raise FileNotFoundError(f"폰트 파일 없음: {nanum}")
        pdfmetrics.registerFont(TTFont('NanumGothic', nanum))
        fallback_font = 'NanumGothic'
        if os.path.exists(dejavu):
            pdfmetrics.registerFont(TTFont('DejaVuSans', dejavu))
            fallback_font = 'DejaVuSans'

        # 스타일
        base = getSampleStyleSheet()
        body = ParagraphStyle('Body', parent=base['Normal'],
                              fontName='NanumGothic', fontSize=11, leading=16, spaceAfter=6)
        h1 = ParagraphStyle('H1', parent=body, fontSize=18, leading=24, spaceBefore=8, spaceAfter=10)
        h2 = ParagraphStyle('H2', parent=body, fontSize=15, leading=21, spaceBefore=6, spaceAfter=8)
        h3 = ParagraphStyle('H3', parent=body, fontSize=13, leading=19, spaceBefore=4, spaceAfter=6)
        code_style = ParagraphStyle('Code', parent=body, fontName='Courier',
                                    backColor='#f4f7fb', leftIndent=6, rightIndent=6,
                                    leading=15, spaceBefore=4, spaceAfter=8)

        flow = []
        lines = md_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')

        def flush_list(lst, numbered=False):
            if not lst: return
            lf = ListFlowable(
                [ListItem(Paragraph(md_inline_to_rl(it, fallback_font), body)) for it in lst],
                bulletType='1' if numbered else 'bullet', start='1', leftIndent=12
            )
            flow.append(lf)

        ul, ol = [], []
        in_code = False
        code_lines = []
        in_math = False
        math_buf = []
        skip_after_header = 0  # 제목 직후 쓰레기 블릿 방지

        for raw in lines:
            line = raw.rstrip()
            if is_trash_line(line):
                if skip_after_header > 0:
                    # 제목 직후 3줄 안의 쓰레기 라인은 무시
                    continue
                # 일반 쓰레기 라인도 무시
                continue

            # 코드 블록
            if line.strip().startswith('```'):
                if in_code:
                    code_html = "<br/>".join(escape_html(x) for x in code_lines) or " "
                    flow.append(Paragraph(code_html, code_style))
                    code_lines, in_code = [], False
                else:
                    flush_list(ul); flush_list(ol, numbered=True)
                    in_code = True
                continue
            if in_code:
                code_lines.append(line); continue

            # 수식 블록
            if not in_math and _block_math_single_re.match(line):
                latex = _block_math_single_re.match(line).group(1).strip()
                flow.append(Image(io.BytesIO(render_latex_png(latex)), width=150, preserveAspectRatio=True))
                continue
            if not in_math and _block_math_open_re.match(line):
                in_math = True
                math_buf.append(_block_math_open_re.match(line).group(1))
                continue
            if in_math:
                if _block_math_close_re.match(line):
                    math_buf.append(_block_math_close_re.match(line).group(1))
                    latex = '\n'.join(math_buf).strip()
                    flow.append(Image(io.BytesIO(render_latex_png(latex)), width=150, preserveAspectRatio=True))
                    math_buf, in_math = [], False
                else:
                    math_buf.append(line)
                continue

            # 빈 줄
            if line.strip() == '':
                flush_list(ul); flush_list(ol, numbered=True)
                flow.append(Spacer(1, 2*mm)); continue

            # 제목
            if line.startswith('### '):
                flush_list(ul); flush_list(ol, numbered=True)
                flow.append(Paragraph(md_inline_to_rl(line[4:].strip(), fallback_font), h3))
                skip_after_header = 3; continue
            if line.startswith('## '):
                flush_list(ul); flush_list(ol, numbered=True)
                flow.append(Paragraph(md_inline_to_rl(line[3:].strip(), fallback_font), h2))
                skip_after_header = 3; continue
            if line.startswith('# '):
                flush_list(ul); flush_list(ol, numbered=True)
                flow.append(Paragraph(md_inline_to_rl(line[2:].strip(), fallback_font), h1))
                skip_after_header = 3; continue

            # 목록
            if re.match(r'^(\-|\*|\+)\s+', line):
                ul.append(line[2:].strip()); continue
            if re.match(r'^\d+\.\s+', line):
                ol.append(re.sub(r'^\d+\.\s+', '', line).strip()); continue

            # 일반 문단 (인라인 수식 포함)
            flush_list(ul); flush_list(ol, numbered=True)
            if _inline_math_re.search(line):
                _append_text_with_inline_math_pdf(flow, line, body, fallback_font)
            else:
                flow.append(Paragraph(md_inline_to_rl(line, fallback_font), body))

            if skip_after_header > 0:
                skip_after_header -= 1

        flush_list(ul); flush_list(ol, numbered=True)
        if in_code and code_lines:
            code_html = "<br/>".join(escape_html(x) for x in code_lines) or " "
            flow.append(Paragraph(code_html, code_style))

        doc = SimpleDocTemplate(buf, pagesize=A4,
                                leftMargin=20*mm, rightMargin=20*mm, topMargin=17*mm, bottomMargin=17*mm)
        doc.build(flow)
        buf.seek(0)
        return buf, "application/pdf", "AI_면접_질문.pdf"
    except Exception as e:
        print(f"[ERR] PDF 생성 실패: {e}")
        raise

# ------------------------------------------------------------------------------------
# DOCX 변환 (블릿/빈 문단 제거 + 수식 이미지)
# ------------------------------------------------------------------------------------
_inline_token = re.compile(r'(\*\*.+?\*\*|__.+?__|`[^`]+`|\*(?!\*)(.+?)(?<!\*)\*|_.+?_)')

def _apply_run_font(run):
    run.font.size = Pt(11)
    run.font.name = '맑은 고딕'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

def _add_inline_runs(p, text):
    pos = 0
    for m in _inline_token.finditer(text):
        if m.start() > pos:
            r = p.add_run(text[pos:m.start()]); _apply_run_font(r)
        token = m.group(0)
        if token.startswith('**') or token.startswith('__'):
            r = p.add_run(token[2:-2]); r.bold = True; _apply_run_font(r)
        elif token.startswith('`') and token.endswith('`'):
            r = p.add_run(token[1:-1]); r.font.name = 'Consolas'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕'); r.font.size = Pt(11)
        else:
            r = p.add_run(token[1:-1]); r.italic = True; _apply_run_font(r)
        pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:]); _apply_run_font(r)

def _add_math_image_run(p, latex):
    # 길이에 따라 폭 동적 조정 (최대 3.5인치)
    width_inch = min(0.8 + 0.05 * len(latex), 3.5)
    img = render_latex_png(latex)
    p.add_run().add_picture(io.BytesIO(img), width=Inches(width_inch))

def _tune_paragraph(p):
    pf = p.paragraph_format
    pf.line_spacing = 1.3
    pf.space_before = Pt(2)
    pf.space_after = Pt(3)

def create_docx_from_markdown(md_text: str):
    buf = io.BytesIO()
    try:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = '맑은 고딕'
        style.font.size = Pt(11)
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

        lines = md_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
        in_code = False
        code_lines = []
        in_math = False
        math_buf = []
        skip_after_header = 0  # 제목 직후 쓰레기 블릿 방지

        for raw in lines:
            line = raw.rstrip()

            if is_trash_line(line):
                # 빈 줄/빈 블릿은 생성하지 않음 (특히 제목 직후)
                if skip_after_header > 0:
                    continue
                continue

            # 코드블록
            if line.strip().startswith('```'):
                if in_code:
                    p = doc.add_paragraph(); _tune_paragraph(p)
                    for cl in code_lines:
                        r = p.add_run(cl + '\n')
                        r.font.name = 'Consolas'
                        r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                        r.font.size = Pt(10)
                    code_lines, in_code = [], False
                else:
                    in_code = True
                continue
            if in_code:
                code_lines.append(line); continue

            # 수식 블록
            if not in_math and _block_math_single_re.match(line):
                p = doc.add_paragraph(); _tune_paragraph(p)
                _add_math_image_run(p, _block_math_single_re.match(line).group(1).strip())
                continue
            if not in_math and _block_math_open_re.match(line):
                in_math = True
                math_buf.append(_block_math_open_re.match(line).group(1))
                continue
            if in_math:
                if _block_math_close_re.match(line):
                    math_buf.append(_block_math_close_re.match(line).group(1))
                    p = doc.add_paragraph(); _tune_paragraph(p)
                    _add_math_image_run(p, '\n'.join(math_buf).strip())
                    math_buf, in_math = [], False
                else:
                    math_buf.append(line)
                continue

            # 제목
            if line.startswith('### '):
                p = doc.add_paragraph(line[4:].strip(), style='Heading 3'); _tune_paragraph(p)
                skip_after_header = 3; continue
            if line.startswith('## '):
                p = doc.add_paragraph(line[3:].strip(), style='Heading 2'); _tune_paragraph(p)
                skip_after_header = 3; continue
            if line.startswith('# '):
                p = doc.add_paragraph(line[2:].strip(), style='Heading 1'); _tune_paragraph(p)
                skip_after_header = 3; continue

            # 목록
            if re.match(r'^(\-|\*|\+)\s+', line):
                text = line[2:].strip()
                if text:
                    p = doc.add_paragraph(style='List Bullet'); _tune_paragraph(p)
                    cursor = 0
                    for m in _inline_math_re.finditer(text):
                        pre = text[cursor:m.start()]
                        if pre: _add_inline_runs(p, pre)
                        _add_math_image_run(p, m.group(1).strip())
                        cursor = m.end()
                    tail = text[cursor:]
                    if tail: _add_inline_runs(p, tail)
                continue
            if re.match(r'^\d+\.\s+', line):
                text = re.sub(r'^\d+\.\s+', '', line).strip()
                if text:
                    p = doc.add_paragraph(style='List Number'); _tune_paragraph(p)
                    cursor = 0
                    for m in _inline_math_re.finditer(text):
                        pre = text[cursor:m.start()]
                        if pre: _add_inline_runs(p, pre)
                        _add_math_image_run(p, m.group(1).strip())
                        cursor = m.end()
                    tail = text[cursor:]
                    if tail: _add_inline_runs(p, tail)
                continue

            # 일반 문단 (인라인 수식 포함)
            p = doc.add_paragraph(); _tune_paragraph(p)
            cursor = 0
            for m in _inline_math_re.finditer(line):
                pre = line[cursor:m.start()]
                if pre: _add_inline_runs(p, pre)
                _add_math_image_run(p, m.group(1).strip())
                cursor = m.end()
            tail = line[cursor:]
            if tail: _add_inline_runs(p, tail)

            if skip_after_header > 0:
                skip_after_header -= 1

        # 열린 코드블록 마무리
        if in_code and code_lines:
            p = doc.add_paragraph(); _tune_paragraph(p)
            for cl in code_lines:
                r = p.add_run(cl + '\n')
                r.font.name = 'Consolas'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                r.font.size = Pt(10)

        if len(doc.paragraphs) == 0:
            t = doc.add_paragraph('AI 생성 면접 질문 및 모범 답변', style='Heading 1'); _tune_paragraph(t)

        doc.save(buf)
        buf.seek(0)
        return buf, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "AI_면접_질문.docx"
    except Exception as e:
        print(f"[ERR] DOCX 생성 실패: {e}")
        raise

# ------------------------------------------------------------------------------------
@app.get("/healthz")
def healthz(): return "ok", 200

@app.get("/")
def root(): return "AI Interview Master backend running", 200

