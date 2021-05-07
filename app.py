# -*- coding: utf-8 -*-
"""
Created on Thu May  6 23:00:24 2021

@author: César Cortez
"""

from flask import Flask, render_template, request, redirect, url_for, session, Markup, flash, send_file, make_response
import pandas as pd
from datetime import date
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import os, io, zipfile, time
#import plotly_express as px
import requests
from utils import main
import pandas as pd
import docx
import random
import os
from os.path import basename
import string

#from utils import utils_data_wrangling, utils_plotly, utils_validations#, utils_google

app = Flask(__name__)


path_output = "D:\Proyectos\Free Code\multi_report\data\output"
# session_id = str(random.randint(0,1000))

session_id = str("")
@app.route("/")
def route():
    # print("SESSION ID",session_id)
    return render_template("Home_2.html")
    
@app.route("/generator",methods=["POST","GET"])
def report_generator():
    global session_id
    # session_id = str(random.randint(0,1000))
    S=20
    session_id = "".join(random.choices(string.ascii_uppercase + string.digits, k = S))


    
    print("SESSION ID",session_id)
    if request.method =="POST":
        
        # feature_selection = request.form.get("form_fields[field_040b2df][]")form_fields[field_040b2df][]
        try:
            word_template = docx.Document(request.files.get('form_fields[field_9445d12]'))
            print(request.files.get('excel_generator'))
            df_generator = pd.read_excel(request.files.get('excel_generator'))
            
            document_name = request.form.get("form_fields[message]")
            document_name = str(document_name).strip()
            
            feature_selection = request.form.getlist("form_fields[field_040b2df][]")
            
            print("FEATURE",feature_selection)
        except TypeError as Err:
            print(Err)
            return render_template(("fail_file_format.html"))
        
        try:
            os.makedirs('./data/output/{new_folder}'.format(new_folder=session_id))
            main.multi_report_n(df_generator,word_template,"reporte",folder_name = session_id,feature_list=feature_selection,column_namer=str(document_name))
        except TypeError as Err:
            print(Err)
            render_template(("fail_file_format.html"))
            
    return render_template("Generator_html.html")

@app.route("/zip_report",methods=["POST"])
def download():
    global session_id
    try:
        file_path = path_output + "/{folder_id}/".format(folder_id=session_id) # local_path
        
        # timestr = time.strftime("%Y%m%d-%H%M%S")
        fileName = "report_collections.zip"
        memory_file = io.BytesIO()
    
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(file_path):
                for file in files:
                    # zipf.write(os.path.join(root, file))
                    print("ROOT",root)
                    print("FILE",file)
                    zipf.write(os.path.join(root,file),basename(file_path+file))
        memory_file.seek(0)
    except:
        return render_template("fail_file_format.html")
    
    return send_file(memory_file,attachment_filename=fileName,as_attachment=True)
    

if __name__=="__main__":
    port = int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)

