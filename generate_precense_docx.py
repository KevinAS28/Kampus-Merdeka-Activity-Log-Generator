# %%
import os, json, subprocess, json, requests
from dotenv import dotenv_values

config = dotenv_values('config.env')

print(config)

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

if os.path.isfile(OUTPUT):
    os.remove(OUTPUT)


document: docx.Document = docx.Document()


table = document.add_table(rows=1, cols=3)
table.style = 'Table Grid'

for week, week_report in week_data.items():
    for day_index, daily_report in enumerate(week_report['data']['daily_report']):
        row_cells = table.add_row().cells
        row_data = [
            f'{week}/{datetime.strptime(daily_report["report_date"], "%Y-%m-%dT%H:%M:%SZ").date()}',
            daily_report['report'],
            # None, #image
            '' if day_index<(len(week_report['data']['daily_report'])-1) else week_report['data']['learned_weekly'],
            # None #image
        ]

        for i, cell in enumerate(row_data):
            if cell is None:
                continue
            row_cells[i].text = str(cell)    

        # paragraph = row_cells[3].paragraphs[0]
        # run = paragraph.add_run()
        # run.add_picture('kevin.png', width = 1400000, height = 1400000)

        # paragraph = row_cells[5].paragraphs[0]
        # run = paragraph.add_run()
        # run.add_picture('kak pam.png', width = 1400000, height = 1400000)        

document.save(OUTPUT)

print('DONE! File:', OUTPUT)

# %%



