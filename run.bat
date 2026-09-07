@echo off
chcp 65001 > nul
title TikZ -> Web App Launcher
cls

echo ================================================================
echo       TIKZ TO PNG / WEB APP - TRÌNH KHỞI ĐỘNG ỨNG DỤNG
echo ================================================================
echo.

:: 1. Kiểm tra Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [LỖI] Không tìm thấy Python trên máy tính của bạn!
    echo Vui lòng cài đặt Python từ https://www.python.org/downloads/
    echo (Nhớ tích chọn "Add Python to PATH" khi cài đặt).
    echo.
    pause
    exit /b 1
)

:: 2. Kiểm tra & cài đặt dependencies nếu cần
echo [1/3] Đang kiểm tra thư viện Python...
python -c "import fastapi, uvicorn, httpx, dotenv" >nul 2>nul
if %errorlevel% neq 0 (
    echo [*] Đang tự động cài đặt các thư viện cần thiết...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [CẢNH BÁO] Cài đặt thư viện gặp lỗi, thử tiếp tục...
    )
) else (
    echo [OK] Thư viện Python đã sẵn sàng.
)

:: 3. Kiểm tra pdflatex & pdftoppm
echo [2/3] Đang kiểm tra công cụ biên dịch LaTeX...
where pdflatex >nul 2>nul
if %errorlevel% neq 0 (
    echo [CẢNH BÁO] Không tìm thấy pdflatex trong PATH!
    echo Nếu đã cài MiKTeX hoặc TeXLive, hãy đảm bảo đã thêm vào PATH.
) else (
    echo [OK] Đã tìm thấy pdflatex.
)

:: 4. Mở trình duyệt sau 2 giây trong luồng nền
echo [3/3] Đang khởi động Web Server tại http://localhost:8000 ...
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8000"

echo.
echo ================================================================
echo  Ứng dụng đang chạy tại: http://localhost:8000
echo  Trình duyệt web sẽ tự động mở trong giây lát...
echo  Để DỪNG ứng dụng: Nhấn Ctrl + C hoặc đóng cửa sổ này.
echo ================================================================
echo.

:: 5. Chạy server chính
python main.py

pause
