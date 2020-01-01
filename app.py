from flask import Flask, render_template, request
from Filter_Copy_Generator import operate
from flask_cors import CORS
import re

app = Flask(__name__)
cors = CORS(app)


@app.route('/', methods=["GET", "POST"])
def hello_world():
    # first need to save the json to the current working directory
    # the method should be POST
    # so this should operate with the POST method
    # then, in this case, it should automatically generate the download page (if possible)
    print("Hola")
    if request.method == "GET":
        return render_template("index.html")
    else:
        print(request.json)
        print(operate(request.json))
        # seems to be fine
        # need to turn to other lines (add <br>)
        tmp = re.sub("\n", "<br>", operate(request.json))
        return re.sub(" ", "&nbsp;", tmp)


if __name__ == '__main__':
    app.run()
