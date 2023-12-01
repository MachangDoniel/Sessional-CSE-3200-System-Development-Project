from flask import Flask, render_template

app = Flask(__name__)

# Sample data for the table
data = [
    {"Name": "John", "Age": 30, "Country": "USA"},
    {"Name": "Alice", "Age": 25, "Country": "Canada"},
    {"Name": "Bob", "Age": 28, "Country": "UK"},
]

@app.route('/')
def display_table():
    return render_template('table.html', data=data)

if __name__ == '__main__':
    app.run(debug=True)
