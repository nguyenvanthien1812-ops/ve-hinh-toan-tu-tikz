"""
TikZ Web Service
================
- GET  /               : Trang web giao diện chính
- GET  /health         : Kiểm tra trạng thái pdflatex, poppler và các AI keys
- POST /generate       : Đề bài → AI (Gemini / Claude) → mã TikZ
- POST /fix            : TikZ lỗi + Log lỗi → AI sửa → TikZ mới
- POST /make-prompt    : Đề bài → sinh prompt chuẩn để gửi Claude.ai / ChatGPT
- POST /render         : Mã TikZ → pdflatex + pdftoppm → PNG (kèm cache & tùy chọn DPI)
- POST /render-pdf     : Mã TikZ → pdflatex → PDF bytes
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional

import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, Header, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

# Nạp biến môi trường từ file .env nếu có
load_dotenv()

# ── LaTeX Template ────────────────────────────────────────────────────────────

TEX_TEMPLATE = r"""\documentclass[tikz,border=12pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{tikz}
\usepackage{tkz-tab}
\usepackage{tkz-euclide}
\usepackage{pgfplots}
\pgfplotsset{compat=1.18}
\usetikzlibrary{calc,arrows.meta,decorations.pathreplacing,patterns,angles,quotes,intersections,positioning,3d,shapes}

\begin{document}
%%TIKZ_CONTENT%%
\end{document}
"""

DEFAULT_DPI = 300

# Bộ nhớ đệm (Cache) trong RAM: key = sha256(tikz + dpi + format), value = bytes
RENDER_CACHE: dict[str, bytes] = {}

# ── AI System Prompts ────────────────────────────────────────────────────────

MAKE_PROMPT_SYSTEM = """Bạn là chuyên gia kiến trúc Prompt TikZ hàng đầu cho giáo viên Toán học Việt Nam.

NHIỆM VỤ: Nhận đề bài hoặc mô tả hình vẽ của giáo viên (kể cả ảnh chụp đề bài nếu có), phân tích chuyên sâu các yếu tố hình học và viết lại thành một bản **PROMPT VẼ HÌNH TIKZ ĐẦY ĐỦ, CHUẨN XÁC, CHUYÊN NGHIỆP NHẤT** để gửi cho bất kỳ AI nào (Claude, ChatGPT, Gemini) sinh mã TikZ chuẩn sách giáo khoa ngay lần đầu.

BẢN PROMPT SINH RA PHẢI BAO GỒM ĐẦY ĐỦ CÁC MỤC SAU:

1. 🎯 LOẠI HÌNH & BỐ CỤC:
   - Xác định chính xác dạng toán: Hình học phẳng (2D) / Hình không gian (3D) / Đồ thị hàm số / Bảng biến thiên (tkz-tab)...
   - Quy định tỉ lệ khung hình cân đối, thoáng đãng, kích thước trực quan.

2. 📐 QUAN HỆ HÌNH HỌC CHUẨN XÁC & CÁCH DỰNG (RẤT QUAN TRỌNG):
   - Tính thẳng hàng: Nếu các điểm thẳng hàng (đường kính, cát tuyến qua tâm...), yêu cầu vẽ 1 đoạn thẳng duy nhất liền mạch \\draw (A) -- (B);, tuyệt đối không vẽ chắp vá làm gãy góc.
   - Tiếp tuyến đường tròn: Yêu cầu vuông góc 100% với bán kính tại tiếp điểm bằng cú pháp calc ($(P)!-1.2cm!90:(O)$) -- ($(P)!1.5cm!90:(O)$) hoặc tkz-euclide. Tiếp tuyến chung ngoài/trong phải tiếp xúc thực sự, không cắt vào trong đường tròn.
   - Giao điểm: Yêu cầu dùng name path và name intersections={of=... and ..., by={...}} hoặc tkz-euclide. TUYỆT ĐỐI CẤM tự giải phương trình tọa độ ra số thập phân nhiều chữ số.
   - Kéo dài đoạn thẳng: Dùng calc ($(A)!-0.3!(B)$) -- ($(A)!1.3!(B)$).

3. 🛡️ QUY TẮC CHỐNG ĐÈ NHÃN CHỮ LÊN ĐƯỜNG VẼ (BẮT BUỘC):
   - MỌI nhãn tên điểm, chữ cái, góc hay số liệu BẮT BUỘC có thuộc tính: fill=white, inner sep=1.2pt để tạo nền trắng che nhẹ nét vẽ bên dưới, chữ luôn sắc nét và không bao giờ bị đường thẳng cắt ngang.
   - Hướng nhãn thoát ra ngoài: Đặt vị trí nhãn lệch ra ngoài vùng hình vẽ (ví dụ: đỉnh trên dùng above=2pt, đáy trái dùng below left=2pt, giao điểm đường tròn thì quay nhãn ra phía ngoài đường tròn).
   - Thứ tự lớp vẽ: Toàn bộ lệnh chấm điểm \\fill và ghi nhãn \\node PHẢI ĐẶT Ở CUỐI CÙNG của khối tikzpicture để luôn hiển thị trên cùng (Top Z-index).

4. 🎨 PHÂN CẤP NÉT VẼ & KÝ HIỆU CHUẨN MỰC:
   - Nét chính (đường tròn, tam giác, hình chính): thick
   - Nét phụ, đường nối tâm, đường gióng: dashed, gray!80!black hoặc thin
   - Điểm mấu chốt: Chấm tròn đen đồng bộ \\fill (\\p) circle (1.5pt);
   - Ký hiệu góc vuông: Dùng \\draw pic[draw, angle radius=2mm] {right angle = ...}; hoặc tkzMarkRightAngle.
   - Tuyệt đối không vẽ mũi tên (->) hay đường cong to[bend] tùy tiện lên các đoạn thẳng.

5. 📦 ĐỊNH DẠNG ĐẦU RA MÃ TIKZ:
   Yêu cầu AI: "Hãy viết mã TikZ hoàn chỉnh bắt đầu bằng \\begin{tikzpicture} và kết thúc bằng \\end{tikzpicture}. Đảm bảo biên dịch thành công ngay với pdflatex (hỗ trợ tiếng Việt qua gói vietnam), không dùng fontspec, không kèm markdown giải thích."

QUY TẮC:
- Viết bằng tiếng Việt sư phạm, khúc chiết, chuẩn xác từng câu chữ.
- Trả về trực tiếp nội dung bản Prompt chuẩn mực, không thêm lời chào mở đầu hay kết thúc thừa thãi."""

TIKZ_SYSTEM = """Bạn là lập trình viên TikZ/LaTeX chuyên nghiệp cho tài liệu toán học Việt Nam.

NHIỆM VỤ: Nhận drawing description chi tiết, sinh ra mã TikZ biên dịch được ngay với pdflatex.

ĐẦU RA — chỉ trả về mã TikZ, KHÔNG có gì khác:
- Không markdown, không dấu ```, không giải thích
- Chỉ phần nằm BÊN TRONG \\begin{document}...\\end{document}
- KHÔNG có \\documentclass, \\usepackage, \\begin{document}, \\end{document}

THƯ VIỆN có sẵn (đã khai báo trong preamble):
  tikz, tkz-tab, tkz-euclide
  usetikzlibrary: calc, arrows.meta, decorations.pathreplacing, patterns, angles, quotes, intersections, positioning, 3d, shapes

QUY TẮC THẨM MỸ & KỸ THUẬT CHUẨN SÁCH GIÁO KHOA:
- ĐOẠN THẲNG HÀNG: Nếu 3 điểm thẳng hàng (đường kính, cát tuyến qua tâm...), phải vẽ 1 lệnh liền mạch \\draw (A) -- (B);, tuyệt đối không vẽ chắp vá làm lệch góc.
- TIẾP TUYẾN: Phải vuông góc 100% với bán kính tại tiếp điểm bằng cú pháp calc ($(P)!-1cm!90:(O)$) hoặc tkz-euclide.
- GIAO ĐIỂM TỰ ĐỘNG: Luôn dùng name intersections hoặc tkz-euclide, không tính nháp tay.
- CHỐNG ĐÈ NHÃN: Mọi node nhãn điểm cạnh nét vẽ phải có fill=white, inner sep=1.2pt để có nền trắng che nét vẽ, chữ luôn nổi rõ.
- CHẤM ĐIỂM: Mọi điểm mấu chốt phải có chấm tròn đen \\fill (X) circle (1.5pt);.
- KÝ HIỆU GÓC VUÔNG: Dùng \\draw pic[draw, angle radius=2mm] {right angle = ...}; hoặc tkzMarkRightAngle.
- PHÂN CẤP NÉT VẼ: Nét chính thick, nét phụ/nối tâm dashed, gray!80!black. Không vẽ mũi tên hay đường cong bend tùy tiện.
- Nhãn toán học dùng $...$ hoặc \\(...\\), không dùng text thuần.
- Bảng biến thiên/xét dấu: dùng tkz-tab với \\tkzTabInit, \\tkzTabLine, \\tkzTabVar.
- Đồ thị hàm số: dùng \\draw[domain=..., samples=80] plot (\\x, {biểu thức}).
- Standalone border mặc định 12pt — KHÔNG cần thêm margin thủ công.
- Nếu yêu cầu không phải hình vẽ toán, trả về: ERROR: Yêu cầu không phù hợp"""

FIX_SYSTEM = """Bạn là chuyên gia debug TikZ/LaTeX.

NHIỆM VỤ: Nhận mã TikZ bị lỗi và log lỗi từ pdflatex, phân tích và trả về mã đã sửa.

QUY TRÌNH:
1. Đọc kỹ dòng lỗi (dòng bắt đầu bằng "!" hoặc "Error")
2. Xác định chính xác nguyên nhân — KHÔNG đoán mò
3. Sửa tối thiểu, giữ nguyên phần còn lại
4. Kiểm tra lại tính nhất quán của toàn bộ code trước khi xuất

ĐẦU RA — chỉ mã TikZ đã sửa:
- Không markdown, không ```, không giải thích
- Chỉ phần BÊN TRONG \\begin{document}...\\end{document}
- KHÔNG có preamble

Nếu lỗi không thể sửa được, trả về: ERROR: <lý do cụ thể>"""

# ── App Init ──────────────────────────────────────────────────────────────────

app = FastAPI(
    title="TikZ to PNG / Web App",
    description="Biên dịch mã TikZ sang ảnh PNG chất lượng cao và tích hợp AI tạo hình vẽ toán học",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Pydantic Request Models ───────────────────────────────────────────────────

class GenerateRequest(BaseModel):
    prompt: Optional[str] = ""
    provider: Optional[str] = "auto"  # "gemini" | "anthropic" | "auto"
    model: Optional[str] = None
    image_base64: Optional[str] = None
    image_mime_type: Optional[str] = None

class FixRequest(BaseModel):
    tikz: str
    error: str
    provider: Optional[str] = "auto"
    model: Optional[str] = None

class MakePromptRequest(BaseModel):
    description: str
    provider: Optional[str] = "auto"
    model: Optional[str] = None

class RenderRequest(BaseModel):
    tikz: str
    dpi: Optional[int] = Field(default=DEFAULT_DPI, ge=72, le=1200)

class TeXLiveProxyRequest(BaseModel):
    document: str


# ── Utilities ─────────────────────────────────────────────────────────────────

def sanitize_tikz_code(code: str) -> str:
    """Làm sạch mã TikZ: loại bỏ markdown wrapper, bóc tách bên trong begin{document} nếu có."""
    code = code.strip()
    
    # 1. Loại bỏ markdown code fence blocks ```...```
    code = re.sub(r"^```[a-zA-Z0-9_-]*\s*", "", code)
    code = re.sub(r"\s*```$", "", code)
    code = code.strip("`").strip()

    # 2. Nếu người dùng dán cả file tex hoàn chỉnh, bóc tách phần trong \begin{document}...\end{document}
    doc_match = re.search(r"\\begin\{document\}(.*?)\\end\{document\}", code, re.DOTALL)
    if doc_match:
        code = doc_match.group(1).strip()

    # 3. Loại bỏ các dòng \documentclass hoặc \usepackage nếu còn sót
    cleaned_lines = []
    for line in code.splitlines():
        trimmed = line.strip()
        if trimmed.startswith(("\\documentclass", "\\usepackage", "\\usetikzlibrary", "\\begin{document}", "\\end{document}")):
            continue
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines).strip()


# ── AI API Helpers (Gemini & Anthropic) ─────────────────────────────────────────

async def call_gemini(
    system: str,
    user_content: str,
    custom_key: Optional[str] = None,
    model: Optional[str] = None,
    image_base64: Optional[str] = None,
    image_mime_type: Optional[str] = None,
) -> str:
    """Gọi Google Gemini API hỗ trợ dynamic model, Vision image và xoay vòng Key."""
    raw_keys = custom_key or os.environ.get("GEMINI_API_KEYS") or os.environ.get("GEMINI_API_KEY") or ""
    keys = [k.strip() for k in re.split(r"[\n,;]+", raw_keys) if len(k.strip()) > 5]

    if not keys:
        raise HTTPException(
            503,
            "Chưa cấu hình GEMINI_API_KEY. Vui lòng nhập API Key trong Cài đặt ⚙ trên web hoặc file .env.",
        )

    target_model = model or "gemini-3.6-flash"
    last_error = ""

    user_parts = []
    if image_base64:
        user_parts.append({
            "inline_data": {
                "mime_type": image_mime_type or "image/png",
                "data": image_base64
            }
        })
    if user_content:
        user_parts.append({"text": user_content})

    for i, api_key in enumerate(keys):
        payload = {
            "system_instruction": {"parts": [{"text": system}]},
            "contents": [{"role": "user", "parts": user_parts}],
            "generationConfig": {"maxOutputTokens": 4096, "temperature": 0.2},
        }

        url = f"https://generativelanguage.googleapis.com/v1beta/models/{target_model}:generateContent?key={api_key}"
        async with httpx.AsyncClient(timeout=60.0) as client:
            try:
                res = await client.post(url, json=payload)
            except Exception as e:
                last_error = f"Lỗi kết nối Gemini API (Key #{i+1}): {str(e)}"
                continue

        if res.status_code == 200:
            data = res.json()
            try:
                text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
                return sanitize_tikz_code(text)
            except (KeyError, IndexError):
                last_error = f"Gemini trả về định dạng không hợp lệ: {str(data)[:200]}"
                continue
        else:
            last_error = f"Gemini API lỗi [{res.status_code}] ở Key #{i+1}: {res.text[:200]}"
            # Nếu 429 hoặc 403 hoặc 400, tự động thử key tiếp theo
            continue

    raise HTTPException(502, f"Không thể gọi Gemini API: {last_error}")


async def call_claude(system: str, user_content: str, custom_key: Optional[str] = None) -> str:
    """Gọi Anthropic Claude API (claude-3-5-sonnet)."""
    api_key = custom_key or os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise HTTPException(503, "Chưa cấu hình ANTHROPIC_API_KEY. Vui lòng nhập API Key trong Cài đặt hoặc file .env.")

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": "claude-3-5-sonnet-20241022",
        "max_tokens": 2048,
        "temperature": 0.2,
        "system": system,
        "messages": [{"role": "user", "content": user_content}],
    }

    url = "https://api.anthropic.com/v1/messages"
    async with httpx.AsyncClient(timeout=60.0) as client:
        try:
            res = await client.post(url, headers=headers, json=payload)
        except Exception as e:
            raise HTTPException(502, f"Lỗi kết nối tới Claude API: {str(e)}")

    if res.status_code != 200:
        raise HTTPException(502, f"Claude API lỗi [{res.status_code}]: {res.text[:300]}")

    data = res.json()
    try:
        text = data["content"][0]["text"].strip()
    except (KeyError, IndexError):
        raise HTTPException(502, f"Claude trả về định dạng không hợp lệ: {str(data)[:200]}")

    return sanitize_tikz_code(text)


async def execute_ai_call(
    system: str,
    user_content: str,
    provider: Optional[str] = "auto",
    gemini_key: Optional[str] = None,
    anthropic_key: Optional[str] = None,
    model: Optional[str] = None,
    image_base64: Optional[str] = None,
    image_mime_type: Optional[str] = None,
) -> str:
    """Tự động điều phối gọi Gemini hoặc Claude tùy cấu hình và model."""
    raw_g = gemini_key or os.environ.get("GEMINI_API_KEYS") or os.environ.get("GEMINI_API_KEY")
    raw_a = anthropic_key or os.environ.get("ANTHROPIC_API_KEY")

    if provider == "anthropic" or (provider == "auto" and raw_a and not raw_g):
        return await call_claude(system, user_content, custom_key=anthropic_key)
    elif provider == "gemini" or (provider == "auto" and raw_g):
        return await call_gemini(
            system, user_content, custom_key=gemini_key, model=model,
            image_base64=image_base64, image_mime_type=image_mime_type
        )
    elif raw_a:
        return await call_claude(system, user_content, custom_key=anthropic_key)
    elif raw_g:
        return await call_gemini(
            system, user_content, custom_key=gemini_key, model=model,
            image_base64=image_base64, image_mime_type=image_mime_type
        )
    else:
        raise HTTPException(
            503,
            "Chưa cấu hình API Key nào (Gemini hoặc Claude). Vui lòng nhấn biểu tượng Cài đặt ⚙ trên web để nhập API Key.",
        )


# ── Health & Info Endpoint ────────────────────────────────────────────────────

@app.get("/health")
def health():
    tools = {k: shutil.which(k) for k in ("pdflatex", "pdftoppm")}
    missing = [k for k, v in tools.items() if v is None]
    has_gemini = bool(os.environ.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEYS"))
    has_anthropic = bool(os.environ.get("ANTHROPIC_API_KEY"))

    return {
        "status": "error" if missing else "ok",
        "missing_tools": missing,
        "tools_found": {k: v for k, v in tools.items() if v is not None},
        "providers": {
            "gemini": has_gemini,
            "anthropic": has_anthropic,
        },
    }


# ── Endpoints ─────────────────────────────────────────────────────────────────

@app.post("/generate")
async def generate(
    req: GenerateRequest,
    x_gemini_key: Optional[str] = Header(None),
    x_anthropic_key: Optional[str] = Header(None),
):
    """Sinh mã TikZ từ mô tả hoặc ảnh đính kèm."""
    prompt = (req.prompt or "").strip()
    if not prompt and not req.image_base64:
        raise HTTPException(400, "Vui lòng nhập đề bài hoặc cung cấp ảnh đề bài.")

    user_prompt = prompt or "Hãy quan sát đề bài toán / hình vẽ trong ảnh đính kèm và tạo mã TikZ hoàn chỉnh, chuẩn xác theo chương trình Toán học."

    # Sinh trực tiếp mã TikZ
    tikz = await execute_ai_call(
        TIKZ_SYSTEM,
        user_prompt,
        provider=req.provider,
        gemini_key=x_gemini_key,
        anthropic_key=x_anthropic_key,
        model=req.model,
        image_base64=req.image_base64,
        image_mime_type=req.image_mime_type,
    )

    if tikz.startswith("ERROR:"):
        raise HTTPException(422, tikz[6:].strip())

    return {"tikz": tikz, "description": user_prompt}


@app.post("/fix")
async def fix(
    req: FixRequest,
    x_gemini_key: Optional[str] = Header(None),
    x_anthropic_key: Optional[str] = Header(None),
):
    """TikZ lỗi + Thông báo lỗi từ LaTeX → AI sửa → TikZ mới."""
    if not req.tikz.strip() or not req.error.strip():
        raise HTTPException(400, "Thiếu mã TikZ hoặc thông báo lỗi.")

    user_msg = f"""Mã TikZ bị lỗi:
```latex
{req.tikz}
```

Thông báo lỗi từ pdflatex:
```
{req.error}
```

Hãy sửa lại mã TikZ để biên dịch được."""

    fixed = await execute_ai_call(
        FIX_SYSTEM,
        user_msg,
        provider=req.provider,
        gemini_key=x_gemini_key,
        anthropic_key=x_anthropic_key,
        model=req.model,
    )

    if fixed.startswith("ERROR:"):
        raise HTTPException(422, fixed[6:].strip())

    return {"tikz": fixed}


@app.post("/make-prompt")
async def make_prompt(
    req: MakePromptRequest,
    x_gemini_key: Optional[str] = Header(None),
    x_anthropic_key: Optional[str] = Header(None),
):
    """Mô tả ngắn → Sinh prompt chi tiết chuẩn để gửi Claude.ai / ChatGPT."""
    desc = req.description.strip()
    if not desc:
        raise HTTPException(400, "Mô tả không được để trống.")

    prompt_text = await execute_ai_call(
        MAKE_PROMPT_SYSTEM,
        desc,
        provider=req.provider,
        gemini_key=x_gemini_key,
        anthropic_key=x_anthropic_key,
        model=req.model,
    )
    return {"prompt": prompt_text}


@app.post("/render")
def render(req: RenderRequest):
    """Mã TikZ → pdflatex → pdftoppm → PNG image bytes."""
    tikz = sanitize_tikz_code(req.tikz)
    if not tikz:
        raise HTTPException(400, "Mã TikZ không được để trống.")

    dpi = req.dpi or DEFAULT_DPI

    # Kiểm tra Cache
    cache_key = hashlib.sha256(f"png:{dpi}:{tikz}".encode("utf-8")).hexdigest()
    if cache_key in RENDER_CACHE:
        return Response(content=RENDER_CACHE[cache_key], media_type="image/png")

    for tool in ("pdflatex", "pdftoppm"):
        if not shutil.which(tool):
            raise HTTPException(503, f"Công cụ {tool} chưa được cài đặt trên server.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        tex_content = TEX_TEMPLATE.replace("%%TIKZ_CONTENT%%", tikz)
        (tmp_path / "input.tex").write_text(tex_content, encoding="utf-8")

        # 1. Chạy pdflatex
        try:
            r1 = subprocess.run(
                ["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "input.tex"],
                cwd=tmp,
                capture_output=True,
                text=True,
                timeout=45,
                encoding="utf-8",
                errors="replace",
            )
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "Quá thời gian biên dịch LaTeX (45 giây).")

        if r1.returncode != 0:
            bad_lines = [l for l in r1.stdout.splitlines() if l.startswith("!") or "Error" in l or "l." in l]
            summary = "\n".join(bad_lines[:12]) or r1.stdout[-800:]
            raise HTTPException(422, f"Lỗi biên dịch LaTeX:\n\n{summary}")

        pdf_path = tmp_path / "input.pdf"
        if not pdf_path.exists():
            raise HTTPException(500, "pdflatex không tạo được file PDF.")

        # 2. Chạy pdftoppm chuyển PDF sang PNG
        try:
            r2 = subprocess.run(
                ["pdftoppm", "-png", "-r", str(dpi), "-singlefile", "input.pdf", "output"],
                cwd=tmp,
                capture_output=True,
                text=True,
                timeout=30,
                encoding="utf-8",
                errors="replace",
            )
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "Quá thời gian chuyển đổi PDF sang PNG.")

        if r2.returncode != 0:
            raise HTTPException(500, f"Lỗi pdftoppm: {r2.stderr}")

        png_path = tmp_path / "output.png"
        if not png_path.exists():
            raise HTTPException(500, "Không tìm thấy file ảnh PNG kết quả.")

        png_bytes = png_path.read_bytes()

        # Lưu vào cache (giới hạn kích thước cache tối đa 100 ảnh)
        if len(RENDER_CACHE) > 100:
            RENDER_CACHE.pop(next(iter(RENDER_CACHE)))
        RENDER_CACHE[cache_key] = png_bytes

        return Response(content=png_bytes, media_type="image/png")


@app.post("/render-pdf")
def render_pdf(req: RenderRequest):
    """Mã TikZ → pdflatex → PDF bytes để tải về."""
    tikz = sanitize_tikz_code(req.tikz)
    if not tikz:
        raise HTTPException(400, "Mã TikZ không được để trống.")

    cache_key = hashlib.sha256(f"pdf:{tikz}".encode("utf-8")).hexdigest()
    if cache_key in RENDER_CACHE:
        return Response(
            content=RENDER_CACHE[cache_key],
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=tikz_drawing.pdf"},
        )

    if not shutil.which("pdflatex"):
        raise HTTPException(503, "pdflatex chưa được cài đặt trên server.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        tex_content = TEX_TEMPLATE.replace("%%TIKZ_CONTENT%%", tikz)
        (tmp_path / "input.tex").write_text(tex_content, encoding="utf-8")

        try:
            r1 = subprocess.run(
                ["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "input.tex"],
                cwd=tmp,
                capture_output=True,
                text=True,
                timeout=45,
                encoding="utf-8",
                errors="replace",
            )
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "Quá thời gian biên dịch LaTeX.")

        if r1.returncode != 0:
            bad_lines = [l for l in r1.stdout.splitlines() if l.startswith("!") or "Error" in l]
            summary = "\n".join(bad_lines[:12]) or r1.stdout[-800:]
            raise HTTPException(422, f"Lỗi biên dịch LaTeX:\n\n{summary}")

        pdf_path = tmp_path / "input.pdf"
        if not pdf_path.exists():
            raise HTTPException(500, "pdflatex không tạo được file PDF.")

        pdf_bytes = pdf_path.read_bytes()
        RENDER_CACHE[cache_key] = pdf_bytes

        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=tikz_drawing.pdf"},
        )


@app.post("/texlive-proxy")
async def texlive_proxy(req: TeXLiveProxyRequest):
    """Proxy gửi request lên https://texlive.net/cgi-bin/latexcgi nhằm tránh lỗi CORS của trình duyệt."""
    if not req.document.strip():
        raise HTTPException(400, "Nội dung tài liệu TeX không được để trống.")

    files = {
        "filecontents[]": ("document.tex", req.document.encode("utf-8"), "text/plain"),
    }
    data = {
        "filename[]": "document.tex",
        "engine": "pdflatex",
        "return": "pdf",
    }

    try:
        async with httpx.AsyncClient(timeout=45.0, follow_redirects=True) as client:
            resp = await client.post("https://texlive.net/cgi-bin/latexcgi", files=files, data=data)
    except Exception as e:
        raise HTTPException(502, f"Không thể kết nối máy chủ TeXLive.net: {str(e)}")

    content_type = resp.headers.get("content-type", "")
    if "pdf" in content_type:
        return Response(content=resp.content, media_type="application/pdf")
    else:
        return Response(content=resp.content, media_type="text/plain", status_code=422)


# ── Frontend Static Hosting ───────────────────────────────────────────────────

# Tìm kiếm file index.html ở các vị trí khả dĩ
def find_frontend_index() -> Optional[Path]:
    base_dir = Path(__file__).resolve().parent
    candidates = [
        base_dir / "index.html",
        base_dir / "frontend" / "index.html",
        base_dir.parent / "frontend" / "index.html",
        base_dir / "static" / "index.html",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None

@app.get("/")
def serve_index():
    index_file = find_frontend_index()
    if index_file and index_file.exists():
        return FileResponse(str(index_file))
    return {
        "message": "TikZ Backend API is running.",
        "docs": "/docs",
        "health": "/health",
    }


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    print("🚀 Đang khởi động TikZ Web Service tại http://localhost:8000 ...")
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
