import requests
from bs4 import BeautifulSoup
import openpyxl
import pandas as pd
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def get_html_doc(url):
    options = Options()
    # options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(url)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'td.w2')))
    html = driver.page_source
    driver.quit()
    return BeautifulSoup(html, 'html.parser')

def encode(s):
    return s.lower().strip().replace(' ', '+').replace("'", '')

def extract_text(td):
    """tdから内部のテキストを取得"""
    if td is None:
        return ""
    a_tag = td.find('a')
    if a_tag:
        return a_tag.get_text().strip()
    return td.get_text().strip()

def add_race_result(dict_, horse_name, race_data):
    """馬名をキーに、レース結果を追加"""
    encoded_name = encode(horse_name)
    if encoded_name not in dict_:
        dict_[encoded_name] = {}
    # race_dataは{"year": {...}}の形式で、キーは年号
    dict_[encoded_name].update(race_data)

def parse_race_page(race_url, race_name, horse_dict):
    try:
        doc = get_html_doc(race_url)
    except:
        return
    
    # レース結果テーブルを探す（border='1', cellspacing='0の属性を持つテーブル）
    table = doc.find('table', {'border': '1', 'cellpadding': '2'})
    if not table:
        return
    
    rows = table.find_all('tr')
    
    # ヘッダー行を特定（最初のセルが"Year"のクラスを持つ）
    header_index = None
    for i, tr in enumerate(rows):
        tds = tr.find_all('td')
        if tds and len(tds) > 0:
            first_cell = tds[0].get_text().strip()
            if first_cell == "Year":
                header_index = i
                break
    
    if header_index is None:
        return
    
    # ヘッダーの直後のデータ行のみを処理
    for tr in rows[header_index + 1:]:
        tds = tr.find_all('td')
        if len(tds) < 13:
            continue
        
        # インデックス: 0=Year, 1=Winner, 2=Sire, 3=Dam, 4=Trainer, 5=Fam., 6=Track, 7=Dist., 8=Grade, 9=Surf., 10=2nd, 11=3rd, 12=Time, 13+=Comment
        yr = extract_text(tds[0]).strip()
        
        # 年号が数字であることを確認（4桁の数字 or "Race Not Run"）
        if not yr or (not yr.isdigit() and "Race Not Run" not in yr):
            continue
        
        # "Race Not Run"の場合はスキップ
        if "Race Not Run" in tds[1].get_text().strip():
            continue
        
        # データ行を確認（Winnerが実際のデータ行）
        winner = extract_text(tds[1]).strip()
        if not winner or winner in ["Winner", ""]:
            continue
        
        sire = extract_text(tds[2]).strip()
        dam = extract_text(tds[3]).strip()
        trainer = extract_text(tds[4]).strip()
        fam = extract_text(tds[5]).strip()
        track = extract_text(tds[6]).strip()
        dist = extract_text(tds[7]).strip()
        grade = extract_text(tds[8]).strip()
        surf = extract_text(tds[9]).strip()
        second = extract_text(tds[10]).strip()
        third = extract_text(tds[11]).strip()
        time = extract_text(tds[12]).strip()
        comment = extract_text(tds[13]).strip() if len(tds) > 13 else ""
        
        if not yr:
            continue
        
        # 勝者のレース結果
        if winner and winner not in ["Winner"]:
            race_data = {
                yr: {
                    "race": race_name,
                    "sire": sire,
                    "dam": dam,
                    "trainer": trainer,
                    "fam": fam,
                    "track": track,
                    "dist": dist,
                    "grade": grade,
                    "surf": surf,
                    "1st": winner,
                    "2nd": second,
                    "3rd": third,
                    "time": time,
                    "comment": comment,
                    "position": "1着"
                }
            }
            add_race_result(horse_dict, winner, race_data)
        
        # 2着のレース結果
        if second and second not in ["2nd"]:
            race_data = {
                yr: {
                    "race": race_name,
                    "sire": "",
                    "dam": "",
                    "trainer": "",
                    "fam": "",
                    "track": track,
                    "dist": dist,
                    "grade": grade,
                    "surf": surf,
                    "1st": winner,
                    "2nd": second,
                    "3rd": third,
                    "time": time,
                    "comment": comment,
                    "position": "2着"
                }
            }
            add_race_result(horse_dict, second, race_data)
        
        # 3着のレース結果
        if third and third not in ["3rd"]:
            race_data = {
                yr: {
                    "race": race_name,
                    "sire": "",
                    "dam": "",
                    "trainer": "",
                    "fam": "",
                    "track": track,
                    "dist": dist,
                    "grade": grade,
                    "surf": surf,
                    "1st": winner,
                    "2nd": second,
                    "3rd": third,
                    "time": time,
                    "comment": comment,
                    "position": "3着"
                }
            }
            add_race_result(horse_dict, third, race_data)

def dump_result(h_dict, base_url, output_file='result.csv', use_columns=False):
    data = []
    
    if use_columns:
        # カラム化して出力（年ごとに行を分ける）
        for k, race_dict in h_dict.items():
            for year, race_info in race_dict.items():
                row = {
                    'horse': k,
                    'profile_url': f"{base_url}/{k}",
                    'year': year,
                    'race': race_info.get('race', ''),
                    'sire': race_info.get('sire', ''),
                    'dam': race_info.get('dam', ''),
                    'trainer': race_info.get('trainer', ''),
                    'fam': race_info.get('fam', ''),
                    'track': race_info.get('track', ''),
                    'dist': race_info.get('dist', ''),
                    'grade': race_info.get('grade', ''),
                    'surf': race_info.get('surf', ''),
                    '1st': race_info.get('1st', ''),
                    '2nd': race_info.get('2nd', ''),
                    '3rd': race_info.get('3rd', ''),
                    'time': race_info.get('time', ''),
                    'comment': race_info.get('comment', ''),
                    'position': race_info.get('position', '')
                }
                data.append(row)
    else:
        # JSON形式で出力
        for k, race_dict in h_dict.items():
            data.append({
                'horse': k,
                'profile_url': f"{base_url}/{k}",
                'notes': json.dumps(race_dict, ensure_ascii=False, indent=2)
            })
    
    df = pd.DataFrame(data)
    df.to_csv(output_file, index=False, quoting=1)  # quoting=1 forces quote_all

def scrape_pedigree_query(list_url=None, excel_file='main.xlsx', sheet='MAIN', cell='A1', use_columns=False):
    if list_url is None:
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb[sheet]
            list_url = ws[cell].value
        except:
            # デフォルトURL
            list_url = "https://www.pedigreequery.com/index.php?query_type=stakes&search_bar=stakes&field=country&h=japan"
    base_url = "https://www.pedigreequery.com"

    # 1. 一覧ページを取得
    doc_list = get_html_doc(list_url)
    print("Title:", doc_list.title.string if doc_list.title else "No title")

    # 2. bodyのHTMLを取り出し、再パース
    raw_html = str(doc_list.body) if doc_list.body else str(doc_list)
    doc_tbl = BeautifulSoup(raw_html, 'html.parser')

    # 3. レースURL収集
    race_dict = {}
    for td in doc_tbl.find_all('td', class_='w2'):
        a_tag = td.find('a')
        if a_tag:
            href = a_tag.get('href')
            if href.startswith('/'):
                href = base_url + href
            href = href.split('#')[0]
            title = a_tag.get_text().strip().replace('\n', '').replace('\r', '')
            if not title:
                title = "(no title)"
            if href not in race_dict:
                race_dict[href] = title

    print(f"レース URL 件数: {len(race_dict)}")

    if not race_dict:
        print("レース URL を取得できませんでした。")
        return

    # 4. 各レース詳細
    horse_dict = {}
    for race_url, race_name in race_dict.items():
        parse_race_page(race_url, race_name, horse_dict)

    # 5. 出力
    dump_result(horse_dict, base_url, use_columns=use_columns)

    print(f"完了！ レース {len(race_dict)} 件 ／ 馬 {len(horse_dict)} 頭")

if __name__ == "__main__":
    import sys
    use_cols = False
    url = None
    
    # コマンドライン引数を処理
    for arg in sys.argv[1:]:
        if arg.lower() == 'true':
            use_cols = True
        else:
            url = arg
    
    if url:
        scrape_pedigree_query(url, use_columns=use_cols)
    else:
        scrape_pedigree_query(use_columns=use_cols)
