from flask import render_template, request, redirect, url_for, send_file
from app import app
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import zipfile

@app.route('/home', methods=['GET'])
def index():
    """Trang chính: hiển thị form upload."""
    return render_template('index.html')


@app.route('/post', methods=['POST'])
def generate_certificates():
    """Xử lý upload Excel + Template, sinh ảnh, nén ZIP."""
    excel_file = request.files.get('excelFile')
    template_image = request.files.get('templateImage')

    if not excel_file or not template_image:
        return "Bạn cần upload cả file Excel và ảnh template."

    output_dir = os.path.join(os.getcwd(), 'output')
    os.makedirs(output_dir, exist_ok=True)

    # Lưu file tạm
    template_path = os.path.join(output_dir, template_image.filename)
    template_image.save(template_path)

    excel_path = os.path.join(output_dir, excel_file.filename)
    excel_file.save(excel_path)

    try:
        df = pd.read_excel(excel_path)
        if df.empty:
            return "Excel file is empty."

        ext = os.path.splitext(template_image.filename)[1].lower()
        result_files = []

        # Danh sách font fallback
        font_candidates = [
            os.path.join('fonts', 'DancingScript-VariableFont_wght.ttf'),
            'DancingScript-VariableFont_wght.ttf',
            os.path.join('fonts', 'Roboto-Regular.ttf'),
            os.path.join('fonts', 'DejaVuSans.ttf'),
            'Roboto-Regular.ttf',
            'DejaVuSans.ttf'
        ]

        def load_font(size):
            for font_path in font_candidates:
                try:
                    return ImageFont.truetype(font_path, size)
                except:
                    continue
            return ImageFont.load_default()

        # --- Vẽ từng học sinh ---
        for idx, row in df.iterrows():
            image = Image.open(template_path)
            image = image.convert('RGB') if ext in ['.jpg', '.jpeg'] else image.convert('RGBA')
            draw = ImageDraw.Draw(image)
            W, H = image.size

            # --- Tiêu đề ---
            font_title_main = load_font(60)
            font_subtitle = load_font(48)
            month_year = "09/2025"
            title_lines = [
                "Tiếng Anh cô Hằng",
                "THÔNG BÁO HỌC PHÍ",
                f"THÁNG {month_year}"
            ]
            y_title = 180
            for line in title_lines:
                current_font = font_subtitle if "Tiếng Anh" in line else font_title_main
                bbox = current_font.getbbox(line)
                w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                x = (W - w) // 2
                color = (153, 0, 0, 255) if "Tiếng Anh" in line else (0, 0, 0, 255)
                draw.text((x, y_title), line, font=current_font, fill=color)
                y_title += h + 10

            # --- Dữ liệu ---
            def format_vnd(val):
                try:
                    n = float(val)
                    if n.is_integer():
                        n = int(n)
                    return f"{n:,}".replace(",", ".") + " VNĐ"
                except:
                    return str(val)

            ten = str(row.get('Họ và tên', row.get('Unnamed: 0', '')))
            lop = str(row.get('Lớp', row.get('Unnamed: 1', '')))
            ngay_hoc = str(row.get('Ngày học', row.get('Unnamed: 2', '')))
            so_buoi = str(row.get('Số buổi học', row.get('Unnamed: 3', '')))
            so_tien_buoi = format_vnd(row.get('Số tiền/buổi', row.get('Unnamed: 4', '')))
            hoc_phi = format_vnd(row.get('Học phí', row.get('Unnamed: 5', '')))
            sach = format_vnd(row.get('Sách', row.get('Unnamed: 6', '')))
            hoc_phi_ton = format_vnd(row.get('Học phí tồn', row.get('Unnamed: 7', '')))
            tong_hoc_phi = format_vnd(row.get('Tổng học phí', row.get('Unnamed: 8', '')))

            content = (
                f"Học sinh: {ten}\n"
                f"Lớp {lop}\n\n"
                f"Ngày học: {ngay_hoc}\n\n"
                f"Số buổi: {so_buoi}    Số tiền/ buổi: {so_tien_buoi}\n\n"
                f"Tổng học phí: {hoc_phi} + {sach} (sách) = {tong_hoc_phi}"
            )

            font_body = load_font(42)
            y_text = 460
            for line in content.split('\n'):
                bbox = font_body.getbbox(line)
                w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                x = (W - w) // 2
                draw.text((x, y_text), line, font=font_body, fill=(0, 0, 0, 255))
                y_text += h + 8

            # --- Lưu ảnh ---
            result_filename = f"result_{idx+1}_{template_image.filename}"
            result_path = os.path.join(output_dir, result_filename)
            image.save(result_path)
            result_files.append(result_path)

        # --- Nén ZIP ---
        zip_filename = 'certificates.zip'
        zip_path = os.path.join(output_dir, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in result_files:
                zipf.write(file, os.path.basename(file))

        # --- Trả link tải ---
        return redirect(url_for('download_file', filename=zip_filename))

    except Exception as e:
        return f"Lỗi xử lý file: {e}"


@app.route('/download/<filename>')
def download_file(filename):
    output_dir = os.path.join(os.getcwd(), 'output')
    file_path = os.path.join(output_dir, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found", 404
