"""
TikZ Word -> PNG Debugger
=========================

Công cụ debug quá trình lấy mã TikZ từ Word (đang bôi đen) và biên dịch
thành ảnh PNG bằng pdflatex + pdftoppm (Poppler), có giao diện Tkinter.

Đây là công cụ DEBUG nên nguyên tắc bắt buộc:
    - Không dùng `except: pass` ở bất cứ đâu.
    - Mọi exception phải được log đầy đủ (kèm traceback) và hiển thị trên GUI.
    - Mọi lệnh subprocess phải log: command line, working directory,
      stdout, stderr, exit code.
    - Không được để giao diện bị treo -> biên dịch chạy trong thread riêng.

Yêu cầu cài đặt:
    pip install pywin32 pillow

Chạy:
    python main.py
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from PIL import Image, ImageTk

try:
    import win32com.client  # type: ignore
    import pywintypes  # type: ignore
    WIN32_AVAILABLE = True
except ImportError:
    # Vẫn cho phép chương trình chạy (ví dụ để test giao diện) nhưng sẽ báo lỗi
    # rõ ràng khi người dùng bấm "Lấy từ Word".
    WIN32_AVAILABLE = False


# =============================================================================
# Cấu hình / hằng số
# =============================================================================

TEX_TEMPLATE: str = r"""\documentclass[tikz,border=5pt]{standalone}

\usepackage{tikz}
\usepackage{tkz-tab}
\usepackage{tkz-euclide}
\usetikzlibrary{calc}
\usetikzlibrary{arrows.meta}
\usetikzlibrary{decorations.pathreplacing}
\usetikzlibrary{patterns}
\usetikzlibrary{angles}
\usetikzlibrary{quotes}
\usetikzlibrary{intersections}
\usetikzlibrary{positioning}

\begin{document}

%%TIKZ_CONTENT%%

\end{document}
"""

TEX_FILENAME: str = "test.tex"
PDF_FILENAME: str = "test.pdf"
PNG_PREFIX: str = "output"          # pdftoppm sẽ sinh ra output-1.png
PNG_FILENAME: str = "output-1.png"  # tên file PNG trang đầu tiên
PDF_DPI: int = 300

LOG_TAG_INFO = "INFO"
LOG_TAG_SUCCESS = "SUCCESS"
LOG_TAG_WARNING = "WARNING"
LOG_TAG_ERROR = "ERROR"

LOG_COLORS = {
    LOG_TAG_INFO: "#1a1a1a",     # đen
    LOG_TAG_SUCCESS: "#0a8f2b",  # xanh
    LOG_TAG_WARNING: "#e07b00",  # cam
    LOG_TAG_ERROR: "#c62828",    # đỏ
}


# =============================================================================
# Kết quả trả về của một lệnh subprocess (để log đầy đủ, không rút gọn)
# =============================================================================

@dataclass
class CommandResult:
    """Lưu lại toàn bộ thông tin của một lần chạy subprocess để debug."""
    command: list[str]
    cwd: str
    returncode: int
    stdout: str
    stderr: str


# =============================================================================
# Lớp kiểm tra môi trường (pdflatex / pdftoppm / magick)
# =============================================================================

class EnvironmentChecker:
    """Kiểm tra các công cụ dòng lệnh cần thiết có tồn tại trong PATH hay không."""

    REQUIRED_TOOLS: dict[str, str] = {
        "pdflatex": "pdflatex",
        "pdftoppm": "Poppler (pdftoppm)",
        "magick": "ImageMagick (magick)",
    }

    @staticmethod
    def check_all() -> dict[str, Optional[str]]:
        """
        Trả về dict: tên_lệnh -> đường dẫn tìm thấy (hoặc None nếu không tìm thấy).
        Không raise exception ở đây, chỉ trả kết quả để nơi gọi tự quyết định log.
        """
        result: dict[str, Optional[str]] = {}
        for exe_name in EnvironmentChecker.REQUIRED_TOOLS:
            result[exe_name] = shutil.which(exe_name)
        return result


# =============================================================================
# Lớp kết nối Word qua COM
# =============================================================================

class WordConnector:
    """Kết nối tới instance Word đang mở (GetActiveObject) và đọc vùng bôi đen."""

    @staticmethod
    def connect() -> "win32com.client.CDispatch":
        """
        Kết nối tới instance Word đang mở.

        Raises:
            RuntimeError: nếu pywin32 chưa được cài đặt.
            Exception: nguyên bản exception từ COM (không bị nuốt/rút gọn),
                       để tầng gọi log đầy đủ traceback.
        """
        if not WIN32_AVAILABLE:
            raise RuntimeError(
                "Chưa cài đặt pywin32 (win32com.client). "
                "Hãy chạy: pip install pywin32"
            )

        # Không bọc try/except ở đây theo kiểu nuốt lỗi -> để nguyên exception
        # bay lên cho tầng gọi (GUI) log đầy đủ traceback.
        word_app = win32com.client.GetActiveObject("Word.Application")
        return word_app

    @staticmethod
    def get_selected_text() -> str:
        """Giữ lại để tương thích; chỉ lấy text, không lấy Range."""
        word_app = WordConnector.connect()
        text, _range = WordConnector.get_selection_text_and_range(word_app)
        return text

    @staticmethod
    def get_selection_text_and_range(word_app: "win32com.client.CDispatch"):
        """
        Lấy đồng thời:
            - text: nội dung đang bôi đen (đã chuẩn hóa \\r -> \\n)
            - word_range: đối tượng Range COM trỏ đúng vị trí bôi đen,
              dùng để sau này chèn ảnh trở lại đúng chỗ đó.

        Không nuốt exception: mọi lỗi COM bay nguyên vẹn lên tầng gọi.
        """
        selection = word_app.Selection
        text: str = selection.Text if selection is not None else ""

        if text is None:
            text = ""

        text = text.replace("\r", "\n").strip()

        # Selection.Range trả về 1 Range COM object độc lập, vẫn giữ đúng
        # vị trí (Start/End) trong document ngay cả khi Selection di chuyển
        # sau đó (ví dụ người dùng bấm chuột chỗ khác trong Word).
        word_range = selection.Range

        return text, word_range

    @staticmethod
    def replace_range_with_image(word_range: "win32com.client.CDispatch",
                                  image_path: Path,
                                  dpi: int) -> None:
        """
        Xóa nội dung tại `word_range` (đoạn TikZ gốc) và chèn ảnh PNG
        `image_path` vào đúng vị trí đó, dưới dạng InlineShape.

        Kích thước ảnh trong Word được tính lại từ kích thước pixel và DPI
        gốc (300 DPI khi xuất PNG) để hiển thị đúng tỉ lệ thực tế trên trang
        giấy, không dùng kích thước pixel thô (sẽ bị quá to).

        Không nuốt exception: mọi lỗi COM / lỗi đọc ảnh bay nguyên vẹn lên
        tầng gọi để log đầy đủ.
        """
        image = Image.open(image_path)
        image.load()  # buộc đọc ngay để phát hiện file ảnh hỏng
        width_px, height_px = image.size

        width_points = width_px / dpi * 72.0
        height_points = height_px / dpi * 72.0

        # Xóa nội dung TikZ cũ đang nằm trong Range này
        word_range.Text = ""

        # Chèn ảnh tại vị trí Range (đã bị xóa text, giờ là 1 điểm chèn)
        inline_shape = word_range.InlineShapes.AddPicture(FileName=str(image_path))

        # Khóa tỉ lệ và đặt kích thước theo DPI thực tế
        inline_shape.LockAspectRatio = True
        inline_shape.Width = width_points
        inline_shape.Height = height_points


# =============================================================================
# Lớp biên dịch TikZ -> PDF -> PNG
# =============================================================================

class LatexCompiler:
    """
    Chịu trách nhiệm:
        1. Tạo thư mục temp
        2. Sinh file test.tex từ template
        3. Chạy pdflatex
        4. Chạy pdftoppm để chuyển PDF -> PNG
    Mọi bước đều gọi callback log_func(message, tag) để GUI hiển thị.
    """

    def __init__(self, base_dir: Path, log_func: Callable[[str, str], None]) -> None:
        self.base_dir: Path = base_dir
        self.temp_dir: Path = base_dir / "temp"
        self.log_func: Callable[[str, str], None] = log_func

    # -------------------------------------------------------------------
    def log(self, message: str, tag: str = LOG_TAG_INFO) -> None:
        self.log_func(message, tag)

    # -------------------------------------------------------------------
    def prepare_temp_dir(self) -> Path:
        """Bước 1: Tạo thư mục temp/ (không xóa nội dung cũ để tiện xem lại)."""
        self.log("===== STEP 1 =====", LOG_TAG_INFO)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.log(f"Đường dẫn temp: {self.temp_dir}", LOG_TAG_INFO)
        self.log(f"Thư mục temp tồn tại: {self.temp_dir.exists()}", LOG_TAG_INFO)
        self.log("Đã tạo/xác nhận thư mục temp", LOG_TAG_SUCCESS)
        return self.temp_dir

    # -------------------------------------------------------------------
    def write_tex_file(self, tikz_content: str) -> Path:
        """Bước 2: Sinh file test.tex từ template + nội dung TikZ người dùng nhập."""
        self.log("===== STEP 2 =====", LOG_TAG_INFO)
        tex_path = self.temp_dir / TEX_FILENAME

        full_content = TEX_TEMPLATE.replace("%%TIKZ_CONTENT%%", tikz_content)

        # Không dùng except: pass -> nếu ghi file lỗi, để exception bay lên
        tex_path.write_text(full_content, encoding="utf-8")

        self.log(f"Đường dẫn test.tex: {tex_path}", LOG_TAG_INFO)
        self.log(f"test.tex tồn tại: {tex_path.exists()}", LOG_TAG_INFO)
        self.log("Đã tạo test.tex", LOG_TAG_SUCCESS)
        return tex_path

    # -------------------------------------------------------------------
    def run_command(self, command: list[str], cwd: Path) -> CommandResult:
        """
        Chạy 1 lệnh subprocess và log TOÀN BỘ: command, cwd, stdout, stderr, exit code.
        Không bắt exception ở đây để lỗi FileNotFoundError (thiếu exe) bay nguyên
        vẹn lên tầng gọi.
        """
        self.log(f"Đang chạy: {' '.join(command)}", LOG_TAG_INFO)
        self.log(f"Working directory: {cwd}", LOG_TAG_INFO)

        completed = subprocess.run(
            command,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        self.log("===== OUTPUT (stdout) =====", LOG_TAG_INFO)
        self.log(completed.stdout if completed.stdout else "(trống)", LOG_TAG_INFO)
        self.log("===== OUTPUT (stderr) =====", LOG_TAG_INFO)
        self.log(completed.stderr if completed.stderr else "(trống)", LOG_TAG_INFO)
        self.log(f"Exit code: {completed.returncode}", LOG_TAG_INFO)

        return CommandResult(
            command=command,
            cwd=str(cwd),
            returncode=completed.returncode,
            stdout=completed.stdout,
            stderr=completed.stderr,
        )

    # -------------------------------------------------------------------
    def compile_pdf(self, tex_path: Path) -> Path:
        """Bước 3: Chạy pdflatex để tạo PDF."""
        self.log("===== STEP 3 =====", LOG_TAG_INFO)

        pdflatex_path = shutil.which("pdflatex")
        if pdflatex_path is None:
            # Ném đúng loại lỗi mà pdflatex sẽ gây ra nếu không có trong PATH
            raise FileNotFoundError(
                "Không tìm thấy pdflatex trong PATH. "
                "Hệ thống không thể chạy lệnh pdflatex."
            )

        command = [
            "pdflatex",
            "-interaction=nonstopmode",
            "-halt-on-error",
            TEX_FILENAME,
        ]

        result = self.run_command(command, cwd=self.temp_dir)

        if result.returncode != 0:
            raise subprocess.CalledProcessError(
                returncode=result.returncode,
                cmd=result.command,
                output=result.stdout,
                stderr=result.stderr,
            )

        pdf_path = self.temp_dir / PDF_FILENAME
        pdf_exists = pdf_path.exists()

        self.log("===== STEP 4 =====", LOG_TAG_INFO)
        self.log(f"test.pdf tồn tại: {pdf_exists}", LOG_TAG_INFO)

        if not pdf_exists:
            raise FileNotFoundError(
                f"pdflatex báo thành công (exit code 0) nhưng không tìm thấy "
                f"file PDF tại: {pdf_path}"
            )

        self.log("Đã tạo PDF", LOG_TAG_SUCCESS)
        return pdf_path

    # -------------------------------------------------------------------
    def convert_to_png(self, pdf_path: Path) -> Path:
        """Bước 5: Chuyển PDF -> PNG bằng pdftoppm, độ phân giải 300 DPI."""
        pdftoppm_path = shutil.which("pdftoppm")
        if pdftoppm_path is None:
            raise FileNotFoundError(
                "Không tìm thấy pdftoppm trong PATH (Poppler chưa được cài đặt "
                "hoặc chưa thêm vào PATH)."
            )

        command = [
            "pdftoppm",
            "-png",
            "-r",
            str(PDF_DPI),
            PDF_FILENAME,
            PNG_PREFIX,
        ]

        result = self.run_command(command, cwd=self.temp_dir)

        if result.returncode != 0:
            raise subprocess.CalledProcessError(
                returncode=result.returncode,
                cmd=result.command,
                output=result.stdout,
                stderr=result.stderr,
            )

        png_path = self.temp_dir / PNG_FILENAME
        png_exists = png_path.exists()

        self.log("===== STEP 5 =====", LOG_TAG_INFO)
        self.log(f"output-1.png tồn tại: {png_exists}", LOG_TAG_INFO)

        if not png_exists:
            raise FileNotFoundError(
                f"pdftoppm báo thành công (exit code 0) nhưng không tìm thấy "
                f"file PNG tại: {png_path}. Có thể PDF có tên trang khác "
                f"(ví dụ output-2.png nếu nhiều trang)."
            )

        self.log("Đã tạo PNG", LOG_TAG_SUCCESS)
        return png_path


# =============================================================================
# GUI chính
# =============================================================================

class TikZDebuggerApp(tk.Tk):
    """Cửa sổ chính của công cụ debug TikZ -> PNG."""

    def __init__(self) -> None:
        super().__init__()

        self.title("TikZ Word -> PNG Debugger")
        self.geometry("1300x850")
        self.minsize(1000, 650)

        # Thư mục làm việc: cùng thư mục với script này
        self.base_dir: Path = Path(__file__).resolve().parent
        self.temp_dir: Path = self.base_dir / "temp"

        # Ảnh PNG hiện tại đang preview (giữ tham chiếu để tránh bị garbage-collect)
        self._current_preview_image: Optional[ImageTk.PhotoImage] = None
        self._current_png_path: Optional[Path] = None

        # Cờ để tránh bấm Compile nhiều lần cùng lúc
        self._is_compiling: bool = False

        # Đối tượng Word Application + Range đã lấy được lần "Lấy từ Word" gần nhất.
        # Range dùng để chèn ảnh trở lại ĐÚNG vị trí đã bôi đen ban đầu.
        # QUAN TRỌNG: các đối tượng COM này chỉ được tạo/truy cập ở main thread
        # (thread chạy mainloop của Tkinter) để tránh lỗi apartment-threading của COM.
        self._word_app: Optional["win32com.client.CDispatch"] = None
        self._word_range: Optional["win32com.client.CDispatch"] = None

        # Tùy chọn: tự động chèn ảnh vào Word ngay sau khi Compile thành công
        self.auto_insert_var: tk.BooleanVar = tk.BooleanVar(value=False)

        self._build_style()
        self._build_widgets()
        self._check_environment_on_startup()

    # -------------------------------------------------------------------
    # Xây dựng giao diện
    # -------------------------------------------------------------------
    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            # Nếu theme "clam" không có sẵn trên hệ thống, giữ theme mặc định.
            pass

        style.configure("TButton", padding=8, font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"))

    def _build_widgets(self) -> None:
        # ---- Thanh nút phía trên ----
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(
            toolbar, text="Lấy từ Word", command=self.on_click_fetch_from_word
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            toolbar, text="Compile", command=self.on_click_compile
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            toolbar, text="Chèn vào Word", command=self.on_click_insert_to_word
        ).pack(side=tk.LEFT, padx=4)

        ttk.Checkbutton(
            toolbar,
            text="Tự động chèn sau khi Compile",
            variable=self.auto_insert_var,
        ).pack(side=tk.LEFT, padx=(4, 12))

        ttk.Button(
            toolbar, text="Mở thư mục", command=self.on_click_open_folder
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            toolbar, text="Xem PDF", command=self.on_click_view_pdf
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            toolbar, text="Xem PNG", command=self.on_click_view_png
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            toolbar, text="Xóa Log", command=self.on_click_clear_log
        ).pack(side=tk.LEFT, padx=4)

        self.status_label = ttk.Label(toolbar, text="", foreground="#555555")
        self.status_label.pack(side=tk.RIGHT, padx=8)

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X)

        # ---- Vùng giữa: TikZ (trái) + Preview PNG (phải) ----
        middle_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        middle_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=6, pady=6)

        # Bên trái: TextBox TikZ
        left_frame = ttk.Frame(middle_pane, padding=4)
        ttk.Label(left_frame, text="Mã TikZ (có thể chỉnh sửa)", style="Header.TLabel").pack(
            anchor=tk.W, pady=(0, 4)
        )
        self.tikz_text = scrolledtext.ScrolledText(
            left_frame, wrap=tk.NONE, font=("Consolas", 11), undo=True
        )
        self.tikz_text.pack(fill=tk.BOTH, expand=True)
        middle_pane.add(left_frame, weight=1)

        # Bên phải: Preview PNG
        right_frame = ttk.Frame(middle_pane, padding=4)
        ttk.Label(right_frame, text="Xem trước PNG", style="Header.TLabel").pack(
            anchor=tk.W, pady=(0, 4)
        )
        self.preview_canvas = tk.Canvas(right_frame, background="#f2f2f2")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        middle_pane.add(right_frame, weight=1)

        # ---- Vùng dưới: Log ----
        log_frame = ttk.Frame(self, padding=(6, 0, 6, 6))
        log_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=False)

        ttk.Label(log_frame, text="Log", style="Header.TLabel").pack(anchor=tk.W)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=18, font=("Consolas", 10), wrap=tk.WORD, state=tk.NORMAL
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Tạo tag màu cho log
        for tag_name, color in LOG_COLORS.items():
            self.log_text.tag_config(tag_name, foreground=color)

    # -------------------------------------------------------------------
    # Kiểm tra môi trường lúc khởi động
    # -------------------------------------------------------------------
    def _check_environment_on_startup(self) -> None:
        self.append_log("===== KIỂM TRA MÔI TRƯỜNG =====", LOG_TAG_INFO)

        tool_paths = EnvironmentChecker.check_all()

        display_names = {
            "pdflatex": "pdflatex",
            "pdftoppm": "Poppler",
            "magick": "ImageMagick",
        }

        any_missing = False
        for exe_name, path_found in tool_paths.items():
            label = display_names.get(exe_name, exe_name)
            if path_found is None:
                self.append_log(f"Không tìm thấy {label}", LOG_TAG_ERROR)
                any_missing = True
            else:
                self.append_log(f"Đã tìm thấy {label} tại: {path_found}", LOG_TAG_SUCCESS)

        if not WIN32_AVAILABLE:
            self.append_log(
                "Không tìm thấy thư viện pywin32 (win32com.client). "
                "Chức năng 'Lấy từ Word' sẽ không hoạt động cho đến khi cài đặt.",
                LOG_TAG_ERROR,
            )
            any_missing = True

        if any_missing:
            self.append_log(
                "Một số công cụ/thư viện còn thiếu, các chức năng liên quan sẽ báo lỗi khi sử dụng.",
                LOG_TAG_WARNING,
            )
        else:
            self.append_log("Môi trường đầy đủ, sẵn sàng sử dụng.", LOG_TAG_SUCCESS)

    # -------------------------------------------------------------------
    # Ghi log (an toàn khi gọi từ thread khác nhờ self.after)
    # -------------------------------------------------------------------
    def append_log(self, message: str, tag: str = LOG_TAG_INFO) -> None:
        """
        Ghi 1 dòng log vào ScrolledText, có màu theo tag.
        Có thể được gọi từ thread nền -> luôn schedule qua self.after(0, ...)
        để đảm bảo an toàn với Tkinter (chỉ main thread được vẽ UI).
        """

        def _do_append() -> None:
            self.log_text.insert(tk.END, message + "\n", tag)
            self.log_text.see(tk.END)

        self.after(0, _do_append)

    def log_from_worker(self, message: str, tag: str) -> None:
        """Callback log dùng cho LatexCompiler (chạy trong thread nền)."""
        self.append_log(message, tag)

    # -------------------------------------------------------------------
    # Xử lý sự kiện: Lấy từ Word
    # -------------------------------------------------------------------
    def on_click_fetch_from_word(self) -> None:
        self.append_log("Đang kết nối tới Word...", LOG_TAG_INFO)

        try:
            word_app = WordConnector.connect()
            selected_text, word_range = WordConnector.get_selection_text_and_range(word_app)
        except FileNotFoundError:
            # GetActiveObject thường không raise FileNotFoundError, giữ lại để
            # phòng trường hợp hệ thống trả lỗi dạng này.
            self._show_exception("Không tìm thấy Word đang chạy.")
            return
        except Exception as exc:  # noqa: BLE001 - phải bắt để hiện lỗi trên GUI
            # Xử lý riêng lỗi COM "Word chưa mở" cho rõ ràng, nhưng vẫn hiện
            # nguyên traceback đầy đủ bên dưới.
            if WIN32_AVAILABLE and isinstance(exc, pywintypes.com_error):
                self.append_log(
                    "Không thể kết nối Word. Word có thể chưa được mở.",
                    LOG_TAG_ERROR,
                )
            self._show_exception("Lỗi khi kết nối/đọc dữ liệu từ Word")
            return

        if not selected_text:
            self.append_log(
                "Chưa chọn (bôi đen) nội dung nào trong Word, hoặc vùng chọn rỗng.",
                LOG_TAG_WARNING,
            )
            messagebox.showwarning(
                "Chưa chọn nội dung", "Vui lòng bôi đen đoạn TikZ trong Word trước."
            )
            return

        # Lưu lại app + range để "Chèn vào Word" sau này biết chèn ảnh vào đâu,
        # kể cả khi người dùng đã bấm chuột sang chỗ khác trong Word.
        self._word_app = word_app
        self._word_range = word_range

        self.tikz_text.delete("1.0", tk.END)
        self.tikz_text.insert("1.0", selected_text)
        self.append_log("Đã kết nối Word", LOG_TAG_SUCCESS)
        self.append_log(f"Đã lấy {len(selected_text)} ký tự từ vùng bôi đen.", LOG_TAG_INFO)
        self.append_log(
            "Đã lưu vị trí (Range) trong Word để có thể chèn ảnh trở lại đúng chỗ này.",
            LOG_TAG_INFO,
        )

    # -------------------------------------------------------------------
    # Xử lý sự kiện: Compile
    # -------------------------------------------------------------------
    def on_click_compile(self) -> None:
        if self._is_compiling:
            messagebox.showinfo("Đang biên dịch", "Vui lòng đợi quá trình biên dịch hiện tại hoàn tất.")
            return

        tikz_content = self.tikz_text.get("1.0", tk.END).strip()
        if not tikz_content:
            messagebox.showwarning("Trống", "Chưa có nội dung TikZ để biên dịch.")
            return

        self._is_compiling = True
        self.status_label.config(text="Đang biên dịch...")
        self.append_log("========================================", LOG_TAG_INFO)
        self.append_log("BẮT ĐẦU QUÁ TRÌNH COMPILE", LOG_TAG_INFO)

        # Chạy compile trong thread riêng để không treo giao diện.
        worker_thread = threading.Thread(
            target=self._compile_worker, args=(tikz_content,), daemon=True
        )
        worker_thread.start()

    def _compile_worker(self, tikz_content: str) -> None:
        """
        Hàm chạy trong thread nền, thực hiện toàn bộ pipeline compile.
        Mọi exception đều được bắt ở lớp ngoài cùng này để hiển thị đầy đủ
        trên GUI (không nuốt lỗi), sau đó cập nhật lại trạng thái UI.
        """
        compiler = LatexCompiler(base_dir=self.base_dir, log_func=self.log_from_worker)

        try:
            self.append_log(f"Đường dẫn temp (dự kiến): {self.temp_dir}", LOG_TAG_INFO)

            compiler.prepare_temp_dir()
            tex_path = compiler.write_tex_file(tikz_content)

            self.append_log(f"Đường dẫn test.tex: {tex_path}", LOG_TAG_INFO)
            self.append_log(f"Có tồn tại không: {tex_path.exists()}", LOG_TAG_INFO)

            pdf_path = compiler.compile_pdf(tex_path)
            png_path = compiler.convert_to_png(pdf_path)

            self._current_png_path = png_path

            # Toàn bộ thao tác tiếp theo (hiển thị preview, và có thể là chèn
            # ảnh vào Word qua COM) PHẢI chạy trên main thread -> schedule qua after().
            self.after(0, lambda: self._on_compile_success(png_path))

            self.append_log("========================================", LOG_TAG_SUCCESS)
            self.append_log("COMPILE HOÀN TẤT THÀNH CÔNG", LOG_TAG_SUCCESS)

        except FileNotFoundError as exc:
            self._log_full_exception(exc, "FileNotFoundError")
        except subprocess.CalledProcessError as exc:
            self._log_full_exception(exc, "CalledProcessError")
        except Exception as exc:  # noqa: BLE001 - phải bắt để hiện lỗi trên GUI, không được nuốt
            self._log_full_exception(exc, type(exc).__name__)
        finally:
            self._is_compiling = False
            self.after(0, lambda: self.status_label.config(text=""))

    def _on_compile_success(self, png_path: Path) -> None:
        """
        Chạy trên MAIN THREAD (được gọi qua self.after từ worker thread).
        Hiển thị PNG lên preview, và nếu người dùng bật "Tự động chèn sau khi
        Compile" thì chèn luôn ảnh vào Word tại vị trí đã lấy TikZ trước đó.
        """
        self._display_png(png_path)

        if self.auto_insert_var.get():
            self.append_log(
                "Tự động chèn vào Word đang bật -> tiến hành chèn ảnh...",
                LOG_TAG_INFO,
            )
            self._insert_image_to_word(png_path)

    def _log_full_exception(self, exc: Exception, exc_type_name: str) -> None:
        """Ghi log đầy đủ, KHÔNG rút gọn, cho mọi loại exception xảy ra khi compile."""
        full_traceback = traceback.format_exc()

        self.append_log("========================================", LOG_TAG_ERROR)
        self.append_log(f"LỖI: {exc_type_name}", LOG_TAG_ERROR)
        self.append_log(str(exc), LOG_TAG_ERROR)

        # Nếu là CalledProcessError, hiện thêm stdout/stderr gốc của lệnh lỗi
        if isinstance(exc, subprocess.CalledProcessError):
            self.append_log(f"Command: {exc.cmd}", LOG_TAG_ERROR)
            self.append_log(f"Return code: {exc.returncode}", LOG_TAG_ERROR)
            if exc.output:
                self.append_log("----- stdout (đầy đủ) -----", LOG_TAG_ERROR)
                self.append_log(str(exc.output), LOG_TAG_ERROR)
            if exc.stderr:
                self.append_log("----- stderr (đầy đủ) -----", LOG_TAG_ERROR)
                self.append_log(str(exc.stderr), LOG_TAG_ERROR)

        self.append_log("----- Traceback đầy đủ -----", LOG_TAG_ERROR)
        self.append_log(full_traceback, LOG_TAG_ERROR)

    # -------------------------------------------------------------------
    # Hiển thị PNG lên canvas preview
    # -------------------------------------------------------------------
    def _display_png(self, png_path: Path) -> None:
        """
        Mở file PNG bằng Pillow và hiển thị lên canvas preview.
        Nếu lỗi, log đầy đủ traceback (không nuốt lỗi).
        """
        try:
            image = Image.open(png_path)
            image.load()  # buộc đọc dữ liệu ngay để phát hiện lỗi file hỏng

            # Co giãn ảnh vừa với canvas hiện tại, giữ tỉ lệ
            canvas_width = max(self.preview_canvas.winfo_width(), 100)
            canvas_height = max(self.preview_canvas.winfo_height(), 100)

            image_ratio = image.width / image.height
            canvas_ratio = canvas_width / canvas_height

            if image_ratio > canvas_ratio:
                new_width = canvas_width
                new_height = int(canvas_width / image_ratio)
            else:
                new_height = canvas_height
                new_width = int(canvas_height * image_ratio)

            new_width = max(new_width, 1)
            new_height = max(new_height, 1)

            resized_image = image.resize((new_width, new_height), Image.LANCZOS)
            self._current_preview_image = ImageTk.PhotoImage(resized_image)

            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(
                canvas_width // 2,
                canvas_height // 2,
                image=self._current_preview_image,
                anchor=tk.CENTER,
            )

            self.append_log(f"Đã hiển thị PNG: {png_path}", LOG_TAG_SUCCESS)

        except Exception as exc:  # noqa: BLE001 - phải hiện lỗi, không được nuốt
            self._log_full_exception(exc, type(exc).__name__)
            self._show_exception("Lỗi khi hiển thị PNG bằng Pillow")

    # -------------------------------------------------------------------
    # Các nút phụ: Mở thư mục / Xem PDF / Xem PNG / Xóa Log
    # -------------------------------------------------------------------
    def on_click_open_folder(self) -> None:
        try:
            self.temp_dir.mkdir(parents=True, exist_ok=True)
            self.append_log(f"Mở thư mục: {self.temp_dir}", LOG_TAG_INFO)
            os.startfile(str(self.temp_dir))  # type: ignore[attr-defined]
        except AttributeError:
            # os.startfile chỉ có trên Windows; nếu không có, thử cách khác
            self._show_exception(
                "os.startfile không khả dụng trên hệ điều hành này "
                "(chức năng này được thiết kế cho Windows)."
            )
        except Exception as exc:  # noqa: BLE001
            self._log_full_exception(exc, type(exc).__name__)
            self._show_exception("Lỗi khi mở thư mục")

    def on_click_view_pdf(self) -> None:
        pdf_path = self.temp_dir / PDF_FILENAME
        self.append_log(f"Đường dẫn test.pdf: {pdf_path}", LOG_TAG_INFO)
        self.append_log(f"test.pdf tồn tại: {pdf_path.exists()}", LOG_TAG_INFO)

        if not pdf_path.exists():
            self.append_log("Chưa có file PDF nào, hãy Compile trước.", LOG_TAG_WARNING)
            messagebox.showwarning("Chưa có PDF", "Chưa có file test.pdf, hãy bấm Compile trước.")
            return

        try:
            os.startfile(str(pdf_path))  # type: ignore[attr-defined]
            self.append_log("Đã mở test.pdf bằng trình xem mặc định.", LOG_TAG_SUCCESS)
        except Exception as exc:  # noqa: BLE001
            self._log_full_exception(exc, type(exc).__name__)
            self._show_exception("Lỗi khi mở file PDF")

    def on_click_view_png(self) -> None:
        png_path = self.temp_dir / PNG_FILENAME
        self.append_log(f"Đường dẫn output-1.png: {png_path}", LOG_TAG_INFO)
        self.append_log(f"output-1.png tồn tại: {png_path.exists()}", LOG_TAG_INFO)

        if not png_path.exists():
            self.append_log("Chưa có file PNG nào, hãy Compile trước.", LOG_TAG_WARNING)
            messagebox.showwarning("Chưa có PNG", "Chưa có file output-1.png, hãy bấm Compile trước.")
            return

        self._display_png(png_path)

    def on_click_clear_log(self) -> None:
        self.log_text.delete("1.0", tk.END)

    # -------------------------------------------------------------------
    # Chèn ảnh PNG trở lại đúng vị trí đã lấy TikZ trong Word
    # -------------------------------------------------------------------
    def on_click_insert_to_word(self) -> None:
        """Nút 'Chèn vào Word': chèn ảnh PNG hiện có vào đúng vị trí đã bôi đen ban đầu."""
        self._insert_image_to_word(self._current_png_path)

    def _insert_image_to_word(self, png_path: Optional[Path]) -> None:
        """
        Chèn file PNG tại `png_path` vào Word, thay thế đúng đoạn TikZ đã lấy
        trước đó (dựa vào self._word_range đã lưu khi bấm "Lấy từ Word").

        Luôn chạy trên main thread. Không nuốt exception: mọi lỗi COM đều
        được log đầy đủ + hiện traceback.
        """
        self.append_log("===== CHÈN ẢNH VÀO WORD =====", LOG_TAG_INFO)

        if png_path is None or not png_path.exists():
            self.append_log(
                "Chưa có file PNG nào để chèn (hãy Compile trước).", LOG_TAG_WARNING
            )
            messagebox.showwarning(
                "Chưa có PNG", "Chưa có file PNG để chèn, hãy bấm Compile trước."
            )
            return

        if self._word_range is None or self._word_app is None:
            self.append_log(
                "Chưa có vị trí (Range) trong Word. Hãy bấm 'Lấy từ Word' trước "
                "để công cụ biết cần chèn ảnh vào đâu.",
                LOG_TAG_WARNING,
            )
            messagebox.showwarning(
                "Chưa xác định vị trí",
                "Hãy bấm 'Lấy từ Word' trước khi chèn, để công cụ biết chèn ảnh vào đâu.",
            )
            return

        self.append_log(f"Đường dẫn ảnh sẽ chèn: {png_path}", LOG_TAG_INFO)

        try:
            WordConnector.replace_range_with_image(
                self._word_range, png_path, dpi=PDF_DPI
            )
        except Exception as exc:  # noqa: BLE001 - phải hiện lỗi, không được nuốt
            # Trường hợp thường gặp: document/Range đã không còn hợp lệ vì
            # người dùng đã đóng file Word, undo nhiều lần, v.v.
            if WIN32_AVAILABLE and isinstance(exc, pywintypes.com_error):
                self.append_log(
                    "Lỗi COM khi thao tác với Range trong Word. Có thể tài liệu "
                    "đã bị đóng/thay đổi kể từ lúc 'Lấy từ Word'.",
                    LOG_TAG_ERROR,
                )
            self._log_full_exception(exc, type(exc).__name__)
            self._show_exception("Lỗi khi chèn ảnh vào Word")
            return

        self.append_log(
            f"Đã chèn '{png_path.name}' vào Word tại đúng vị trí đã lấy TikZ.",
            LOG_TAG_SUCCESS,
        )

        # Range đã bị "tiêu thụ" (đoạn text đã bị xóa và thay bằng ảnh) ->
        # xóa tham chiếu để tránh chèn nhầm lần nữa vào cùng vị trí cũ nếu
        # người dùng lỡ bấm "Chèn vào Word" thêm lần nữa mà chưa "Lấy từ Word" mới.
        self._word_range = None

    # -------------------------------------------------------------------
    # Tiện ích hiển thị lỗi ra messagebox (song song với log, không thay thế log)
    # -------------------------------------------------------------------
    def _show_exception(self, context_message: str) -> None:
        full_traceback = traceback.format_exc()
        messagebox.showerror(
            context_message,
            f"{context_message}\n\nChi tiết đầy đủ đã được ghi vào Log.\n\n{full_traceback[-1200:]}",
        )


# =============================================================================
# Entry point
# =============================================================================

def main() -> None:
    app = TikZDebuggerApp()
    app.mainloop()


if __name__ == "__main__":
    main()