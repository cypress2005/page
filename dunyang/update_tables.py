import os
import io
import glob
try:
    from xlsx2html import xlsx2html
    from bs4 import BeautifulSoup
    import pypdfium2 as pdfium
    from PIL import Image
except ImportError:
    print("請先安裝 requirements.txt 中的套件 (pip install -r requirements.txt)")
    import sys
    sys.exit(1)

def update_html():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    def get_latest_file(*patterns):
        files = []
        for pattern in patterns:
            files.extend(glob.glob(pattern))
        if not files:
            return None
        return max(files, key=os.path.getmtime)

    # 尋找最新的 Excel, PDF, 或是圖檔
    file_building = get_latest_file(
        os.path.join(base_dir, "file", "*建物選配表*.xlsx"),
        os.path.join(base_dir, "file", "*建物選配表*.pdf"),
        os.path.join(base_dir, "file", "*建物選配表*.png"),
        os.path.join(base_dir, "file", "*建物選配表*.jpg"),
        os.path.join(base_dir, "file", "*建物選配表*.jpeg")
    )
    file_parking = get_latest_file(
        os.path.join(base_dir, "file", "*車位選配表*.xlsx"),
        os.path.join(base_dir, "file", "*車位選配表*.pdf"),
        os.path.join(base_dir, "file", "*車位選配表*.png"),
        os.path.join(base_dir, "file", "*車位選配表*.jpg"),
        os.path.join(base_dir, "file", "*車位選配表*.jpeg")
    )

    if not file_building or not file_parking:
        print("錯誤：在 file 資料夾中找不到建物選配表或車位選配表 (支援 Excel、PDF 或圖檔)！請確認檔名包含這些字眼。")
        return

    target_building_html = os.path.join(base_dir, "choose_building.html")
    target_parking_html = os.path.join(base_dir, "choose_parking.html")
    
    if not os.path.exists(target_building_html) or not os.path.exists(target_parking_html):
        print(f"錯誤：找不到網頁主檔 (choose_building.html 或 choose_parking.html)！")
        return

    def get_table_html(excel_path):
        out_stream = io.StringIO()
        # 將 excel 轉成 HTML 結構並存入記憶體
        xlsx2html(excel_path, out_stream)
        out_stream.seek(0)
        soup = BeautifulSoup(out_stream.read(), "html.parser")
        return soup.find("table")

    def clean_table(table, last_main_col_letter):
        from string import ascii_uppercase
        if not table:
            return table
        
        cols = table.find_all("col")
        main_count = ascii_uppercase.index(last_main_col_letter) + 1
        for i, col in enumerate(cols):
            if i >= main_count:
                col.decompose()
                
        for cell in table.find_all(["td", "th"]):
            cell_id = cell.get("id", "")
            if "!" in cell_id:
                col_address = ''.join([c for c in cell_id.split('!')[-1] if c.isalpha()]).upper()
                if len(col_address) > 1 or col_address > last_main_col_letter:
                    cell.decompose()
        return table

    def convert_pdf_to_images(pdf_path):
        # 使用 pypdfium2 將 PDF 轉出為高畫質圖片
        pdf = pdfium.PdfDocument(pdf_path)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        soup = BeautifulSoup("", "html.parser")
        wrapper = soup.new_tag("div", style="display: flex; flex-direction: column; gap: 20px; align-items: center;")
        
        for i in range(len(pdf)):
            page = pdf[i]
            # scale=3 提高產出圖片的解析度
            image = page.render(scale=3).to_pil()
            img_filename = f"{base_name}_第{i+1}頁.png"
            img_path = os.path.join(os.path.dirname(pdf_path), img_filename)
            image.save(img_path)
            
            img_tag = soup.new_tag("img")
            img_tag['src'] = f"file/{img_filename}"
            img_tag['style'] = "max-width: 100%; height: auto; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);"
            wrapper.append(img_tag)
            
        return wrapper

    def create_image_element(img_path):
        filename = os.path.basename(img_path)
        soup = BeautifulSoup("", "html.parser")
        wrapper = soup.new_tag("div", style="display: flex; justify-content: center;")
        img_tag = soup.new_tag("img")
        img_tag['src'] = f"file/{filename}"
        img_tag['style'] = "max-width: 100%; height: auto; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);"
        wrapper.append(img_tag)
        return wrapper

    def get_content_element(file_path, last_main_col_letter):
        if not file_path:
            return None
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.xlsx':
            print(f"  [Excel] 正在解析: {os.path.basename(file_path)}")
            table = get_table_html(file_path)
            return clean_table(table, last_main_col_letter)
        elif ext == '.pdf':
            print(f"  [PDF] 正在轉換並匯出圖片: {os.path.basename(file_path)}")
            return convert_pdf_to_images(file_path)
        elif ext in ['.png', '.jpg', '.jpeg']:
            print(f"  [圖片] 正在載入圖檔: {os.path.basename(file_path)}")
            return create_image_element(file_path)
        return None

    print(f"正在處理建物資料 (最新檔案: {os.path.basename(file_building)}) ...")
    b_table = get_content_element(file_building, "I")
    
    print(f"\n正在處理車位資料 (最新檔案: {os.path.basename(file_parking)}) ...")
    p_table = get_content_element(file_parking, "H")

    print(f"\n正在更新網頁檔: choose_building.html ...")
    with open(target_building_html, "r", encoding="utf-8") as f:
        soup_b = BeautifulSoup(f.read(), "html.parser")
    b_container = soup_b.find("div", id="building-container")
    if b_container and b_table:
        b_container.clear()
        b_container.append(b_table)
    with open(target_building_html, "w", encoding="utf-8") as f:
        f.write(str(soup_b))

    print(f"正在更新網頁檔: choose_parking.html ...")
    with open(target_parking_html, "r", encoding="utf-8") as f:
        soup_p = BeautifulSoup(f.read(), "html.parser")
    p_container = soup_p.find("div", id="parking-container")
    if p_container and p_table:
        p_container.clear()
        p_container.append(p_table)
    with open(target_parking_html, "w", encoding="utf-8") as f:
        f.write(str(soup_p))
        
    print("\n✅ 更新成功！請分別開啟 choose_building.html 與 choose_parking.html 來查看。")

if __name__ == "__main__":
    update_html()
