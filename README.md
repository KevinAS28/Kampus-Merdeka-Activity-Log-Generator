install the requirements first:
python -m pip install -r requirements.txt

fill the config.env file

for the activity id, go to kampusmedeka website -> login -> kegiatanku -> kegiatan aktif -> your role. Then see your url browser. in my case, it is: https://kampusmerdeka.kemdikbud.go.id/activity/active/1231234 , so the ACTIVITY_ID = 1231234

then run the python file:
python generate_precense_docx.py

open the generated docx file

open to donation :)