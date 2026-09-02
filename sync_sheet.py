import gspread
from google.oauth2.service_account import Credentials
import json
import os

SPREADSHEET_ID = '1yPOv4d9XSQIdyTKcc1T8A2PnTP3oExwDaZrMyZUtW0A'
SHEET_NAME = '오사카_교토' 

def main():
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_values()

    headers = data[0]
    rows = data[1:]

    processed_data = {}
    current_day = ""

    for row in rows:
        while len(row) < 8:
            row.append("")
            
        day_val = row[0].strip()
        if day_val != "":
            current_day = day_val
            if current_day not in processed_data:
                processed_data[current_day] = []
        
        # 날짜를 아직 못 찾았거나 완전히 빈 줄이면 건너뛰기 (에러 방지)
        if not any(row) or current_day == "":
            continue
            
        processed_data[current_day].append(row)

    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>가족 여행 일정</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 10px; margin: 0; background-color: #f9f9f9; }
        .header { text-align: center; margin-bottom: 20px; color: #333; }
        .tab { overflow-x: auto; white-space: nowrap; background-color: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 15px; -webkit-overflow-scrolling: touch; }
        .tab button { background-color: inherit; border: none; outline: none; cursor: pointer; padding: 14px 20px; font-size: 15px; color: #666; font-weight: bold; }
        .tab button.active { color: #007aff; border-bottom: 3px solid #007aff; }
        .tabcontent { display: none; background-color: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 15px; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 600px; }
        th, td { border-bottom: 1px solid #eee; padding: 10px; text-align: left; }
        th { background-color: #f4f4f4; color: #555; }
        a { color: #007aff; text-decoration: none; word-break: break-all; }
    </style>
    </head>
    <body>
    <div class="header"><h2>✈️ 오사카/교토 가족 여행</h2></div>
    <div class="tab">
    """

    for i, day in enumerate(processed_data.keys()):
        active_class = ' active' if i == 0 else ''
        html_content += f'<button class="tablinks{active_class}" onclick="openTab(event, \'tab{i}\')">{day}</button>\n'

    html_content += "</div>\n"

    for i, (day, day_rows) in enumerate(processed_data.items()):
        display = 'block' if i == 0 else 'none'
        html_content += f'<div id="tab{i}" class="tabcontent" style="display:{display}">\n'
        html_content += "<table>\n<tr>"
        
        for h in headers[1:]:
            html_content += f"<th>{h}</th>"
        html_content += "</tr>\n"
        
        for r in day_rows:
            html_content += "<tr>"
            for j, cell in enumerate(r[1:]):
                if cell.startswith('http'):
                    html_content += f'<td><a href="{cell}" target="_blank">링크 연결</a></td>'
                else:
                    html_content += f"<td>{cell}</td>"
            html_content += "</tr>\n"
        
        html_content += "</table>\n</div>\n"

    html_content += """
    <script>
    function openTab(evt, tabName) {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        document.getElementById(tabName).style.display = "block";
        evt.currentTarget.className += " active";
    }
    </script>
    </body>
    </html>
    """

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)

if __name__ == "__main__":
    main()
