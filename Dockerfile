FROM python:3.9

WORKDIR /kmsdataextraction

COPY requirements.txt .

RUN pip install -r requirements.txt

COPY ./app ./app

COPY Data_Extraction.ipynb .

COPY Completed_Anonym_KMS_Assessment_Discussion_Framework_R03.3-.docx .

COPY wsgi.py .

CMD ["python", "wsgi.py"]
