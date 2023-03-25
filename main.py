import os
import csv
import shutil
import openpyxl
import json
from database_connection import database_connection
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
import gspread
from datetime import date
from google.oauth2.service_account import Credentials
import pandas as pd
from gspread_formatting import *



app = Flask(__name__)




def get_service_account(key_file_path):
    service_account = gspread.service_account(key_file_path)
    return service_account


def get_worksheet(service_account, sheet_name):
    worksheet = service_account.open(sheet_name).sheet1

    return worksheet


def add_new_col_and_update_date(worksheet):
    today = get_date()
    worksheet.add_cols(1)
    last_col = worksheet.col_count + 1
    worksheet.update_cell(1, last_col, today)
    return last_col


def get_date():
    return str(date.today())


def get_all_keywords(worksheet):
    row_keyword = (worksheet.col_values(1))
    return row_keyword[1:]


def get_keyword_row(worksheet, keyword):
    row = int(str(worksheet.find(keyword)).split(" ")[1].removeprefix("R").split("C")[0])
    return row


def get_rank_dict():
   
    ROW_SHEET = row_sheet()

    keyword_dict = {}
    df = pd.read_excel(ROW_SHEET)

    keywords = df[df.columns[0]].to_list()
    rank = df[df.columns[1]].to_list()
    urls = df[df.columns[3]].to_list()

    for i in range(len(keywords)):

        url = urls[i]
        if type(url) is float:
            url = " "
        if rank[i] == "-":

            keyword_dict[keywords[i]] = [0, urls[i]]
        else:

            keyword_dict[keywords[i]] = [rank[i], urls[i]]

    return keyword_dict


def update_data(worksheet, row, col, rank, url):
    if type(url) is float:
        url = " "
    print(f"Updating: Row: {row},Col: {col},Rank: {rank}, Url: {url}")
    worksheet.update_cell(row=row, col=2, value=url)
    worksheet.update_cell(row=row, col=col, value=rank)


def delete_sheet_1(key_file_path, sheet_id):
    service_account = get_service_account(key_file_path)
    workbook = service_account.open_by_key(sheet_id)
    sheet = workbook.sheet1
    workbook.del_worksheet(sheet)
    sheet = workbook.sheet1
    sheet.update_title("Keywords with Ranking")


def copy_row_sheet_to_new_sheet(new_sheet_id, key_file_path):
    row_sheet = "row_copy"
    
    service_account = get_service_account(key_file_path)
    worksheet = get_worksheet(service_account, row_sheet)
    new_sheet = worksheet.copy_to(spreadsheet_id=new_sheet_id)
    delete_sheet_1(key_file_path, new_sheet_id)


def create_new_sheet(key_file_path, project_name):
    file_name = f"{project_name}_result"
    folder_id = "1g6A2JG9dm7hChWQ2BfcV13tnQqDMZuty"
    service = get_service_account(key_file_path)
    new_sheet = service.create(title=file_name, folder_id=folder_id)
    new_sheet_id = (new_sheet.id)
    copy_row_sheet_to_new_sheet(new_sheet_id, key_file_path)
    return new_sheet.url, new_sheet_id


def update_result_sheet(key_file_path, sheet_name):
    print(sheet_name)
    service_account = get_service_account(key_file_path)
    worksheet = get_worksheet(service_account, sheet_name)

    col = add_new_col_and_update_date(worksheet)
    keyword_list = get_all_keywords(worksheet)

    rank_keyword_dict = get_rank_dict()

    for keyword in keyword_list:
        try:
            row = get_keyword_row(worksheet, keyword)
            rank = int(rank_keyword_dict[keyword][0])
            url = rank_keyword_dict[keyword][1]
            update_data(worksheet, row, col, rank, url)
        except Exception as e:
            print(f"{e} Keyword Details Not Found")




def add_project_to_db(project_name, project_details):
    project_details = json.dumps(project_details)
    query = f"insert into new_project (project_name,project_info) values ('{project_name}','{project_details}')"
    commit_query_executer(query)


def convert_csv_to_xlsx(csv_file):
    csv_data = []
    with open(csv_file) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)

    wb = openpyxl.Workbook()
    sheet = wb.active
    for row in csv_data:
        sheet.append(row)

    wb.save('static/upload_files/row_data.xlsx')


def commit_query_executer(query):
    db_connecton, db_cursor = database_connection()
    db_cursor.execute(query)
    db_connecton.commit()
    db_cursor.close()


def fetch_query_executer(query):
    db_connecton, db_cursor = database_connection()
    db_cursor.execute(query)
    fetch_data = db_cursor.fetchall()
    db_cursor.close()
    return fetch_data

def delete_unwanted_rows(row_file):
    global workbook
    global worksheet
    ROW_SHEET = row_file
    workbook = openpyxl.load_workbook(ROW_SHEET)
    
    worksheet = workbook.active
    worksheet.delete_rows(idx = 1, amount =6)
    workbook.save(ROW_SHEET)

def row_sheet():
    file_name = (os.listdir("static/upload_files")[0])
    return f"static/upload_files/{file_name}"

def get_project_name_and_rank_sheet_url():
    return_dict = {}
    query = "select * from new_project"
    all_projects = fetch_query_executer(query)

    for project in all_projects:
        
        project_details = json.loads(project[2])
        share_link = project_details['share_link']
        
        return_dict[project[1]] = share_link
    return return_dict


def delete_row(row_file):
    os.remove(row_file)


def get_project_name():
    project_list = []
    query = f"select project_name from new_project"
    project_list_fetch = fetch_query_executer(query)

    for project in project_list_fetch:
        project_list.append(project[0])
    return project_list


def get_project_details(project_name):
    query = f"select project_info from new_project where project_name = '{project_name}'"

    project_name = fetch_query_executer(query)

    return project_name[0][0]


@app.route('/')
def home():
    project_list = get_project_name()
    
    return render_template('home.html', project_list=project_list)



@app.route('/upload', methods=['POST'])
def upload_file():
    form_dict = request.form.to_dict()
    project_details = get_project_details(form_dict["project_name"])
    project_details = json.loads(project_details)
    print(form_dict)
    print(project_details)
    file = request.files['upload_row_file']
    file.save('static/upload_files/' + file.filename)
    csv_file = row_sheet()
    convert_csv_to_xlsx(csv_file)
    delete_row(csv_file)
    row_file = "static/upload_files/row_data.xlsx"
    delete_unwanted_rows(row_file)
   

    key_file_location = "static/secret_file/secret_file.json"
    sheet_name = f"{project_details['project_name']}_result"
    
    update_result_sheet(key_file_location, sheet_name)
    delete_row(row_file)
    
    return render_template('success.html')



@app.route('/add_project')
def add_project():
    return render_template('add_project.html')


@app.route('/rank_sheet')
def rank_sheet():
    project_dict = get_project_name_and_rank_sheet_url()
    print(project_dict)
    print(type(project_dict))
    return render_template('rank_sheet.html', project_dict=project_dict)


@app.route('/add_success', methods=['POST'])
def add_success():
   
    project_details = request.form.to_dict()
    
    key_file_location = "static/secret_file/secret_file.json"

    
    file_name = f"{project_details['project_name']}_result"
  
    share_link, sheet_id = create_new_sheet(key_file_location, project_details["project_name"])



    project_details['sheet_id'] = sheet_id
    project_details['share_link'] = share_link

    add_project_to_db(project_details['project_name'], project_details)

    return render_template('add_success.html')


if __name__ == "__main__":
    cwd = (os.getcwd())
    uploader_folder = os.path.join(cwd, "/static/upload_files/")
    app.config['UPLOAD_FOLDER'] = uploader_folder
    app.run(debug=False,host='0.0.0.0')
