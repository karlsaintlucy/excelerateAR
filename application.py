"""Put the docstring here."""
from flask import Flask, render_template

app = Flask(__name__)


@app.route('/')
def index():
    """Return hello flask."""
    name = "Karl Johnson"
    return render_template('excelerate.html', name=name)


if __name__ == "__main__":
    app.run()
