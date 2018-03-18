#!usr/bin/python2.7
# -*- coding: utf-8 -*-

from flask import Flask

app = Flask(__name__)

@app.route('/')
def index():
    return 'hello world'

if __name__ == "__main__":
    app.run(host='0.0.0.0',debug=True,port=80)

