from docxtpl import DocxTemplate
from flask import Flask, render_template, request, send_file
import io
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



if __name__ == "__main__":
    app.run(debug=True)