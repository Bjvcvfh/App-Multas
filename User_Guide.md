(CLEAR ENVIRONMENT)

python -m venv venv
venv\Scripts\activate

pip install pyinstaller pandas pdfplumber docxtpl pillow pypdf docx2pdf pywin32 ttkbootstrap
python app_tk.py

logs_multas.csv generated on \App-Multas\output\logs_multas.csv
