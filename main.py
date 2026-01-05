from docxtpl import DocxTemplate
from flask import Flask, render_template, request, send_file
from faker import Faker
import tempfile
import datetime 
app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def create_document():
    if request.method == "POST":
        context = dict(request.form)
        context["date"] = datetime.datetime.today().strftime("%d/%m/%Y")
        print(context)
        doc = DocxTemplate("template.docx")
        doc.render(context)
        
        stream = tempfile.TemporaryFile()
        doc.save(stream)
        stream.seek(0)
        return send_file(stream, download_name="letter.docx")
    else:
        return render_template("form.html")

@app.route("/example")
def example():
    fake = Faker()
    context = {
        "address": fake.address(),
        "insurer_address": fake.address(),
        "policy_number": fake.random_int(min=100000, max=999999),
        "insurance_sum": fake.random_int(min=10000, max=1000000),
        "start_date": fake.date_between(start_date="-1y", end_date="today"),
        "end_date": fake.date_between(start_date="today", end_date="+1y"),
        "regularity": fake.random_element(elements=("monthly", "annually")),
        "excess": f"£{fake.random_int(min=0, max=1000)}",
        "renewal_date": fake.date_between(start_date="+1y", end_date="+2y"),
    }
    doc = DocxTemplate("template.docx")
    doc.render(context)
    
    stream = tempfile.TemporaryFile()
    doc.save(stream)
    stream.seek(0)
    return send_file(stream, download_name="example.docx")
    

@app.route("/template")
def send_template():
    return send_file("template.docx")
if __name__ == "__main__":
    app.run(debug=True)