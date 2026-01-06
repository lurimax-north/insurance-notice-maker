from docxtpl import DocxTemplate, InlineImage
from flask import Flask, render_template, request, send_file
from faker import Faker
import tempfile
import io
import datetime
from PIL import Image
import base64
from docx.shared import Mm


app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def create_document():
    if request.method == "POST":
        context = dict(request.form)
        print(context)
        start_date_formatted = datetime.datetime.strptime(
            context["start_date"], "%Y-%m-%d").strftime("%d/%m/%Y")
        end_date_formatted = datetime.datetime.strptime(
            context["end_date"], "%Y-%m-%d").strftime("%d/%m/%Y")
        renewal_date_formatted = datetime.datetime.strptime(
            context["renewal_date"], "%Y-%m-%d").strftime("%d/%m/%Y")
        context["renewal_date"] = renewal_date_formatted
        context["start_date"] = start_date_formatted
        context["end_date"] = end_date_formatted

        context["date"] = datetime.datetime.today().strftime("%d/%m/%Y")
        image_data = context.pop("signature")
        image_content = image_data.split(";")[1]
        image_base64 = image_content.split(",")[1]

        image_stream = io.BytesIO()
        Image.open(io.BytesIO(base64.b64decode(image_base64))).save(
            image_stream, format="PNG"
        )
        image_stream.seek(0)

        doc = DocxTemplate("template.docx")
        signature = InlineImage(
            doc, image_descriptor=image_stream, width=Mm(80), height=Mm(30)
        )
        context["signature"] = signature
        doc.render(context)

        stream = tempfile.TemporaryFile()
        doc.save(stream)
        stream.seek(0)
        return send_file(stream, download_name="Notice of Cover.docx")
    else:
        return render_template("index.html")


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
        "excess": f"£{fake.random_int(min=0, max=1000)} escape of water, £{fake.random_int(min=0, max=1000)} subsidence, £{fake.random_int(min=0, max=1000)} for all other claims",
        "renewal_date": fake.date_between(start_date="+1y", end_date="+2y"),
        "signature": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMcAAAA+BAMAAABtthIkAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAADBQTFRFAAAAgAAAAIAAgIAAAACAgACAAICAgICAwMDA/wAAAP8A//8AAAD//wD/AP//////ex+xxAAABFdJREFUWMPtl7uW3CAMhumg86vSQQcddHpVd1JHJMD4Oh4nu+7CSfbseow+/RKSGFXeX7P6D3kdQiA/4XcgtLOTx19RfkP/zZVnEAz1Zdf+smO3Asrowt6HE8M9g8xTNamqsdmM7TonsKHMdxCyw/Q9JDpPyfncLA+TIpA4WOZur1UPlZjMYQl9U1ieZtmdlmB+irRanbqFVCs9vbgGy+oExB+5m6380urCLWQWj7s/bg2WgmS+CYlqc/ZuIWa1hINRNVlTA/ZZiNJbZ28gOEJCZPaPGXIlJMkzObpWwebxHWSEhGB7WLkSSac1I2zZNdOoJm4DfHRR7Ty4gSy5JU6L3z4miAZFGzp+geIESnznt5X2MPOJtw4eQrAL4Q0z5CApoqaPVEAL1bAuIU44leQTv60BkJORdDIPIV4aCeXEpWhoDpLw1PQlXbOPwVsg7jG2SdbiE1cHP4kPw8WNCQ33RFMChsB24lR0F9LS7ksyHLqJgdIxbXD1RJA1qMsziAgBqZXAmzkg7B//LEFqAOqh4BwDFMV6qy7DJ4AUp1CXfUo+QwhyDaybIU3SgKNne5KVWgNGIoQTlTkwQbxR2Ytayirsg3UDSRmqkOCZMAci3XqybJkkWvzPQuYIRkl2RiPua/KSsOkZhJtsDaxD8ZSd5qj3wlQtWqlkzUKg0hFmA/KZk5wf6/QDhHiLtF2cUlHAEpjU/OMxUaNFEyK/VPWRh9Z3ZBzkPji/Qoilo3QGSzXkdWLVrczgauZfYw7R1bBtLbbfY3kAIRlT0gNxciJIjlaRtGfHDM0Oy3CUM0SbSb4u/QSSa+gl7cYjH8pp6ZWUfXZBChLXkXQR6weQ2qisuGNdPVKf7V0u/A5plwxpqNXdYzSofF0PIG0U1qvGLqd9uVPEz5D2RnauV/4J0ieICDkvTgd3wW+XOjmXxHccXtZfQahZSLvJNkRKB4inMjhD8tK8SM7oEUJN6bUQuTBa2g9e9tbx/z12s5ukjx0g3QN7KYSnGNc5yAkmCXhdmSvHQV5Xcjlu0+aOkH7/IXUZdvEJbE7WsV23up/We59o44a23R8PkKV7XmcEAxUl5i2cPjk82EbC7CFLpaLaHXSOj4QBaPIpJC6U44WLztTNtesAWd61W3WUZQYLh6egkwl/7BrJn3XHcROWWttAxmVxFUJVQSWIJAIRsf/CkN0ZUYdOp0hHXiHDwRHQXF2kURaEImQ39vj8luuVlPPSo8IOsvSQ9WjVmmUNYYEE3H+JyxeB2mB4vDU5A7Iw0I6cDQXjAW0OEvPLl7V83RsQM4J59m59kta6uxNxWAtk8cpe1WGXQnIrftDqP0FGsNRlI5fg0185fwUJ3Vf7aVjQPxMGBEewwr+b+gbppueb28GPIb0ZtQnzFqT7b18SUiE9I/hOQjqk2Sb77YLwE0irkRcZAqnXY4oG3oTUQ/xa0lcIvcoQyKydepUhEHLudvj8CuT99R/yd5A/BOwVsIky5msAAAAASUVORK5CYII=",
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
