import docx
import docx2txt
import sqlalchemy
import pandas as pd
from flask import Flask, render_template, request
import os

engine = sqlalchemy.create_engine(os.environ.get("POSTGRES"))
engine.execute(sqlalchemy.text("CREATE TABLE IF NOT EXISTS companies (company_id SERIAL PRIMARY KEY, company_name VARCHAR (100) NOT NULL)"))
engine.execute(sqlalchemy.text("CREATE TABLE IF NOT EXISTS questions (question_id SERIAL PRIMARY KEY, question VARCHAR (1000) NOT NULL, answer VARCHAR (2000) NOT NULL, section VARCHAR (500) NOT NULL, number INT NOT NULL, company_name VARCHAR (100), company_id INT NOT NULL, version VARCHAR(20), FOREIGN KEY (company_id) REFERENCES companies (company_id))"))

app = Flask(__name__)

def upload_question_to_database(question, answer, section, company, version, number, sql_engine):
    companies_df = pd.read_sql(sqlalchemy.text("SELECT * FROM companies WHERE company_name = '" + company + "'"), sql_engine)
    if companies_df.shape[0] == 0:
        sql_engine.execute(sqlalchemy.text(f"INSERT INTO companies(company_name) VALUES ('{company}')"))
    questions_df = pd.read_sql(sqlalchemy.text("SELECT * FROM questions WHERE question = '" + question + "' AND answer = '" + answer + "'"), sql_engine)
    if questions_df.shape[0] == 0:
        company_id = pd.read_sql(sqlalchemy.text("SELECT company_id FROM companies WHERE company_name = '" + company + "'"), sql_engine)["company_id"][0]
        sql_engine.execute(sqlalchemy.text(f"INSERT INTO questions (question, answer, section, number, company_name, company_id, version) VALUES ('{question}', '{answer}', '{section}', {number}, '{company}', {company_id}, '{version}')"))

def read_questions(company, version, current_section, i, sql_engine, all_lines, questions, sections):
    question_number = 1
    while i < len(all_lines):
        current_question = all_lines[i]
        try:
            next_ten_questions = questions[questions.index(current_question) + 1: questions.index(current_question) + 11]
        except:
            i += 1
            continue
        current_answer = all_lines[i+1]
        i += 2
        while i < len(all_lines):
            if all_lines[i] in next_ten_questions:
                break
            if all_lines[i] == sections[min(sections.index(current_section) + 1, len(sections) - 1)]:
                current_section = all_lines[i]
                i += 1
                break
            current_answer += "\n" + all_lines[i]
            i += 1
        upload_question_to_database(current_question, current_answer, current_section, company, version, question_number, sql_engine)
        question_number += 1

@app.route("/company", methods=["POST"])
def get_company():
    company_id = int(request.form.get("company"))
    questions = list(engine.execute(sqlalchemy.text(f"SELECT number, question, answer, version, company_name, section FROM questions WHERE company_id = {company_id}")))
    return render_template("company.html", questions=questions)

@app.route("/", methods=["GET", "POST"])
def homepage():
    if request.method == "GET":
        current_questions = list(engine.execute(sqlalchemy.text("SELECT number, question, answer, version, company_name, section FROM questions")))
        current_companies = list(engine.execute(sqlalchemy.text("SELECT * FROM companies")))
        return render_template("index.html", questions=current_questions, companies=current_companies)
    else:
        for filename, file in request.files.items():
            file.save(filename)
            text = docx2txt.process(filename)
            all_lines = [line.strip().replace("'", "") for line in text.split("\n") if line != ""]
            with open(filename, 'w') as f:
                pass
            doc = docx.Document(file)
            questions = [paragraph.text.strip().replace("'", "") for paragraph in doc.paragraphs if paragraph.text != "" and paragraph.style.name in ["Heading 2", "Heading 3"]]
            sections = [paragraph.text.strip().replace("'", "") for paragraph in doc.paragraphs if paragraph.text != "" and paragraph.style.name == "Heading 1"]
            i = 0
            version_not_found = True
            current_section = ""
            questions_started = False
            while not questions_started:
                if all_lines[i] == "Technology Assessment of":
                    company_name = all_lines[i+1]
                    i += 2
                elif "Document Version" in all_lines[i] and version_not_found:
                    document_version = all_lines[i].split(" ")[-1]
                    version_not_found = False
                    i += 1
                elif all_lines[i] in sections:
                    current_section = all_lines[i]
                    i += 1
                    questions_started = True
                else:
                    i += 1

            read_questions(company_name, document_version, current_section, i, engine, all_lines, questions, sections)
        current_questions = list(engine.execute(sqlalchemy.text("SELECT number, question, answer, version, company_name, section FROM questions")))
        current_companies = list(engine.execute(sqlalchemy.text("SELECT * FROM companies")))
        return render_template("index.html", questions=current_questions, companies=current_companies)