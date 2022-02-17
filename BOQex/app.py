from flask import Flask, request, redirect, url_for, render_template, send_file
from werkzeug.utils import secure_filename
from PyPDF2 import PdfFileWriter
import os
import pandas as pd
import numpy as np
from pickle import TRUE
import win32com.client
from win32com.client import constants as c
import pythoncom

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'
ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}

app = Flask(__name__, static_url_path="/static")
DIR_PATH = os.path.dirname(os.path.realpath(__file__))
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
# limit upload size upto 8mb
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('No file attached in request')
            return redirect(request.url)
        file = request.files['file']
        file.save(os.path.join(app.config['UPLOAD_FOLDER'],file.filename))
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            # filename = secure_filename(file.filename)
            BOQ(os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'],file.filename)))
            return redirect(url_for('uploaded_file', filename='readme.txt'))
    return render_template('index.html')  

def BOQ(excel_file):
    df = pd.read_excel(excel_file)
    for j in df.columns.values:
        for i in df.index:
            if df[j][i] == "Item Description":
                x=i
                y=j
            if df[j][i] == "Quantity":
                p=i
                q=j
            if df[j][i] == "Units":
                u=i
                v=j
            if df[j][i] == "Sl.\nNo." or df[j][i] == "Sl.No.":
                c=j
    lines = []
    for i in df.index:
        st = str(df[y][i])
        check1 = st.find("Tmt bar")
        check2 = st.find("HYSD bar")
        check3 = st.find("Steel reinforcement")
        check4 = st.find("cutting, bending, placing")
        check6 = st.find("Sail")
        check7 = st.find("Rinl")
       
        if check1 != -1 or check2 != -1 or check3 != -1 or check4 != -1 or check6 != -1 or check7 != -1:
            if np.isnan(df[q][i]):
                lines.append("steel")
    #             print(df[c][i])
                r=i+1
                while(df[c][r]!=df[c][i]+1):
                    if not np.isnan(df[q][r]):
                        lines.append(df[y][r])
                        lines.append(df[q][r])
                        lines.append(df[v][r])
                    r=r+1
            else:
                lines.append("steel")
                lines.append(df[q][i])
                lines.append(df[v][i])
    with open('readme.txt', 'w') as f:
        for line in lines:
            f.write(str(line))
            f.write('\n')


@app.route('/<filename>')
def uploaded_file(filename):
   return send_file( os.path.abspath('readme.txt'), filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)