# TikZ → Ảnh (Web App Vẽ Hình Toán Học)

Ứng dụng web hiện đại giúp giáo viên và học sinh vẽ hình toán học bằng mã TikZ / LaTeX và AI, tự động chuyển đổi sang ảnh PNG sắc nét (300/600 DPI) và PDF vector để dán trực tiếp vào Microsoft Word, PowerPoint hoặc đề thi.

---

## ✨ Tính năng nổi bật

1. **✨ Tạo hình tự động bằng AI**:
   - Nhập mô tả bài toán bằng tiếng Việt (đồ thị hàm số, bảng biến thiên, hình học phẳng, hình không gian,...).
   - Hỗ trợ **Google Gemini** và **Anthropic Claude**.
   - Pipeline 2 bước: Phân tích drawing description chi tiết → Sinh mã TikZ chuẩn → Tự động biên dịch & tự sửa lỗi cú pháp nếu có.
2. **📋 Tạo Prompt chuẩn hóa**:
   - Tự động sinh prompt chất lượng cao để copy gửi sang [Claude.ai](https://claude.ai) hoặc [ChatGPT](https://chatgpt.com).
3. **{ } Soạn thảo & Biên dịch TikZ trực tiếp**:
   - Hỗ trợ đầy đủ các gói: `tikz`, `tkz-tab`, `tkz-euclide`, `calc`, `arrows.meta`, `patterns`, `angles`,...
   - Tích hợp sẵn **thư viện hình mẫu**: Bảng biến thiên (bậc 3, phân thức), bảng xét dấu, tam giác vuông, đường tròn ngoại tiếp, hình chóp $S.ABCD$, lăng trụ tam giác, đồ thị hàm số, vòng tròn lượng giác, sơ đồ Venn.
4. **📋 Sao chép 1-Click vào Word (Clipboard API)**:
   - Nhấn nút **"📋 Sao chép ảnh"** → Mở Word/PowerPoint và nhấn **`Ctrl + V`** để dán ngay ảnh chất lượng cao 300 DPI mà không cần tải về máy.
5. **🔍 Xem trước tương tác**:
   - Phóng to/thu nhỏ (Zoom), chuyển đổi nền (Lưới ô vuông, Trắng, Tối, Trong suốt).
   - Tải về định dạng PNG (150, 300, 600 DPI) hoặc PDF vector.

---

## 🚀 Hướng dẫn chạy trên máy tính (Local)

### 1. Cài đặt Python Dependencies
```bash
pip install -r requirements.txt
```

### 2. Cài đặt LaTeX & Poppler (Nếu chưa có)
- **Windows**: Cài đặt [MiKTeX](https://miktex.org/download) (hoặc qua Scoop: `scoop install miktex`).
- **Ubuntu/Debian**:
  ```bash
  sudo apt update && sudo apt install -y texlive-latex-extra texlive-science texlive-pictures poppler-utils
  ```

### 3. Cấu hình API Key (Tùy chọn)
Bạn có thể cấu hình API Key theo 2 cách:
- **Cách 1 (Khuyên dùng)**: Mở giao diện web → Nhấn nút **⚙ Cài đặt** ở góc trên bên phải → Nhập Gemini Key / Anthropic Key (được lưu an toàn trong trình duyệt).
- **Cách 2**: Tạo file `.env` từ `env.example`:
  ```env
  GEMINI_API_KEY=AIzaSy...
  ANTHROPIC_API_KEY=sk-ant-...
  ```

### 4. Khởi động Web App
Chạy một trong các lệnh sau trong thư mục gốc:

```bash
# Cách 1: Chạy trực tiếp với Python
python main.py

# Cách 2: Chạy với Uvicorn
uvicorn main:app --reload --port 8000
```

Mở trình duyệt tại: **`http://localhost:8000`**

---

## 🐳 Triển khai với Docker / Cloud (Render.com, VPS)

### 1. Build & Run Docker cục bộ
```bash
docker build -t tikz-web-app .
docker run -p 8000:8000 -e GEMINI_API_KEY=your_key tikz-web-app
```

### 2. Deploy lên Render.com
1. Đẩy code lên GitHub repository.
2. Truy cập [render.com](https://render.com) → **New Web Service** → Chọn repo.
3. Thiết lập:
   - **Runtime**: `Docker`
   - **Environment Variables**: Thêm `GEMINI_API_KEY` hoặc `ANTHROPIC_API_KEY`.
4. Bấm **Deploy** và nhận link truy cập trực tuyến.
