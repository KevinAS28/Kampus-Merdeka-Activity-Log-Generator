# %%
import os, json, subprocess, json, requests
from dotenv import dotenv_values

config = dotenv_values('config.env')


# %%


headers = {
    'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Referer': 'https://kampusmerdeka.kemdikbud.go.id/',
    'sec-ch-ua-mobile': '?0',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'sec-ch-ua-platform': '"Windows"',
}

json_data = {
    'email': config['EMAIL'],
    'password': config['PASSWORD'],
}

credential = json.loads(requests.post('https://api.kampusmerdeka.kemdikbud.go.id/user/auth/login/mbkm', headers=headers, json=json_data).text)
print(credential)


# %%
week_data = dict()

for i in range(1, int(config['WEEK_COUNT'])+1):
  headers = {
      'authority': 'api.kampusmerdeka.kemdikbud.go.id',
      'accept': '*/*',
      'accept-language': 'en-US,en;q=0.9',
      'authorization': 'Bearer '+credential['data']['access_token'],
      'origin': 'https://kampusmerdeka.kemdikbud.go.id',
      'referer': 'https://kampusmerdeka.kemdikbud.go.id/',
      'sec-ch-ua': '"Google Chrome";v="111", "Not(A:Brand";v="8", "Chromium";v="111"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"Windows"',
      'sec-fetch-dest': 'empty',
      'sec-fetch-mode': 'cors',
      'sec-fetch-site': 'same-site',
      'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36',
  }

  response = requests.get(
      f'https://api.kampusmerdeka.kemdikbud.go.id/magang/report/perweek/{config["ACTIVITY_ID"]}/{i}',
      headers=headers,
  )

  data = json.loads(response.text)

  week_data[i] = data

print(json.dumps(week_data[1], indent=4))

# %% [markdown]
# 

# %%
from datetime import datetime, timedelta, datetime
import docx

OUTPUT = f'presences_{config["ACTIVITY_ID"]}.docx'

try:
  if os.path.isfile(OUTPUT):
      os.remove(OUTPUT)
except PermissionError:
   print(f'File {OUTPUT} sedang dipakai atau script ini tidak memiliki akses write')
   exit(1)

document: docx.Document = docx.Document()

WEEKDAYS = 'Senin Selasa Rabu Kamis Jumat Sabtu Minggu'.split(' ')

table = document.add_table(rows=1, cols=3)
table.style = 'Table Grid'

for week, week_report in week_data.items():
    try:
      for day_index, daily_report in enumerate(week_report['data']['daily_report']):
          the_date = datetime.strptime(daily_report["report_date"], "%Y-%m-%dT%H:%M:%SZ").date()          
          row_cells = table.add_row().cells
          row_data = [
              f'Minggu ke-{week} / {WEEKDAYS[the_date.weekday()]} - {the_date}',
              daily_report['report'],
              '' if day_index<(len(week_report['data']['daily_report'])-1) else week_report['data']['learned_weekly'],
          ]

          for i, cell in enumerate(row_data):
              if cell is None:
                  continue
              row_cells[i].text = str(cell)    
    except KeyError:
       print(f'Week {week} data cannot be read')


document.save(OUTPUT)

print('DONE! File:', OUTPUT)

# %%



