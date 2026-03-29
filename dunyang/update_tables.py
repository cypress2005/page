import os
import io
import glob
try:
    from xlsx2html import xlsx2html
    from bs4 import BeautifulSoup
except ImportError:
    print("請先安裝 requirements.txt 中的套件 (pip install -r requirements.txt)")
    import sys
    sys.exit(1)

def update_html():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 尋找最新的 Excel 檔案
    building_files = glob.glob(os.path.join(base_dir, "file", "*建物選配表*.xlsx"))
    parking_files = glob.glob(os.path.join(base_dir, "file", "*車位選配表*.xlsx"))
    
    excel_building = building_files[0] if building_files else None
    excel_parking = parking_files[0] if parking_files else None

    if not excel_building or not excel_parking:
        print("錯誤：在 file 資料夾中找不到建物選配表或車位選配表！請確認檔名包含這些字眼。")
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

    print(f"正在讀取建物資料: {os.path.basename(excel_building)} ...")
    b_table = get_table_html(excel_building)
    b_table = clean_table(b_table, "I")
    
    print(f"正在讀取車位資料: {os.path.basename(excel_parking)} ...")
    p_table = get_table_html(excel_parking)
    p_table = clean_table(p_table, "H")

    print(f"正在更新網頁檔: choose_building.html ...")
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
        
    print("✅ 更新成功！請分別開啟 choose_building.html 與 choose_parking.html 來查看。")

if __name__ == "__main__":
    update_html()
