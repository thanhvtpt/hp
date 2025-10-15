from flask import render_template, request, redirect, url_for, send_file, flash
from app import app
import os

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files.get('excelFile')
        template_image = request.files.get('templateImage')
        import pandas as pd
        from PIL import Image, ImageDraw, ImageFont
        import zipfile
        import datetime

        output_dir = os.path.join(os.getcwd(), 'output')
        os.makedirs(output_dir, exist_ok=True)

        template_path = None
        if template_image and template_image.filename:
            template_path = os.path.join(output_dir, template_image.filename)
            template_image.save(template_path)

        excel_path = None
        if excel_file and excel_file.filename:
            excel_path = os.path.join(output_dir, excel_file.filename)
            excel_file.save(excel_path)

        if template_path and excel_path:
            try:
                df = pd.read_excel(excel_path)
                if df.empty:
                    return "Excel file is empty."
                ext = os.path.splitext(template_image.filename)[1].lower()
                result_files = []

                for idx, row in df.iterrows():
                    image = Image.open(template_path)
                    if ext in ['.jpg', '.jpeg']:
                        image = image.convert('RGB')
                    else:
                        image = image.convert('RGBA')
                    draw = ImageDraw.Draw(image)

                    # Font Unicode hỗ trợ tiếng Việt (ưu tiên Dancing Script, sau đó Roboto hoặc DejaVuSans)
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

                    W, H = image.size

                    # --- THÊM TIÊU ĐỀ ---
                    font_title_main = load_font(60)
                    font_subtitle = load_font(48)

                    # Lấy tháng/năm hiện tại (hoặc thay bằng tháng cố định)
                    month_year = "09/2025"
                    title_lines = [
                        "Tiếng Anh cô Hằng",
                        "THÔNG BÁO HỌC PHÍ",
                        f"THÁNG {month_year}"
                    ]

                    # Vẽ tiêu đề căn giữa phía trên
                    y_title = 180
                    for line in title_lines:
                        current_font = font_subtitle if "Tiếng Anh" in line else font_title_main
                        bbox = current_font.getbbox(line)
                        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                        x = (W - w) // 2
                        color = (153, 0, 0, 255) if "Tiếng Anh" in line else (0, 0, 0, 255)
                        draw.text((x, y_title), line, font=current_font, fill=color)
                        y_title += h + 10

                    # --- ĐỌC DỮ LIỆU EXCEL ---
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

                    if not any([ten, lop, ngay_hoc, so_buoi, so_tien_buoi, hoc_phi, sach, hoc_phi_ton, tong_hoc_phi]):
                        debug_cols = ', '.join(df.columns)
                        content = f"Không lấy được dữ liệu từ Excel. Các cột hiện có: {debug_cols}"
                    else:
                        content = (
                            f"Học sinh: {ten}\n"
                            f"Lớp {lop}\n\n"
                            f"Ngày học: {ngay_hoc}\n\n"
                            f"Số buổi: {so_buoi}    Số tiền/ buổi: {so_tien_buoi}\n\n"
                            f"Tổng học phí: {hoc_phi} + {sach} (sách) = {tong_hoc_phi}"
                        )

                    # --- VẼ NỘI DUNG CHÍNH ---
                    font_body = load_font(42)
                    lines = content.split('\n')
                    y_text = 460  # vị trí bắt đầu nội dung (có thể tăng lên hoặc giảm xuống)
                    for line in lines:
                        bbox = font_body.getbbox(line)
                        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                        x = (W - w) // 2
                        draw.text((x, y_text), line, font=font_body, fill=(0, 0, 0, 255))
                        y_text += h + 8

                    # --- GHI CHÚ DƯỚI CÙNG ---
                    # note = (
                    #     "Phụ huynh chuyển khoản vui lòng không ghi nội dung, "
                    #     "chụp lại màn hình gửi lại giúp cô và hoàn thiện học phí trước 25 hàng tháng!\n"
                    #     "Cô Hằng cảm ơn cả nhà ạ!"
                    # )
                    # font_note = load_font(22)
                    # note_lines = note.split('\n')
                    # total_note_height = sum([font_note.getbbox(line)[3] - font_note.getbbox(line)[1] + 2 for line in note_lines])
                    # y_note = H - total_note_height - 50
                    # for line in note_lines:
                    #     bbox = font_note.getbbox(line)
                    #     w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                    #     x = (W - w) // 2
                    #     draw.text((x, y_note), line, font=font_note, fill=(102, 51, 153, 255))
                    #     y_note += h + 4

                    # --- LƯU ẢNH ---
                    result_filename = f"result_{idx+1}_{template_image.filename}"
                    result_path = os.path.join(output_dir, result_filename)
                    if ext in ['.jpg', '.jpeg']:
                        image.save(result_path, format='JPEG')
                    else:
                        image.save(result_path)
                    result_files.append(result_path)

                # --- NÉN ZIP ---
                zip_filename = 'certificates.zip'
                zip_path = os.path.join(output_dir, zip_filename)
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file in result_files:
                        zipf.write(file, os.path.basename(file))

                download_link = url_for('download_file', filename=zip_filename)
                return f"Đã tạo ảnh cho tất cả dòng trong Excel! <a href='{download_link}'>Tải file ZIP</a>"

            except Exception as e:
                return f"Lỗi xử lý file: {e}"

        return "Bạn cần upload cả file Excel và ảnh template."

    return render_template('index.html')


# Route tải file ZIP
@app.route('/download/<filename>')
def download_file(filename):
    output_dir = os.path.join(os.getcwd(), 'output')
    file_path = os.path.join(output_dir, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found", 404
