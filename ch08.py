from flask import Flask, render_template, request, redirect, url_for
import pandas as pd

app = Flask(__name__)

# Excel file path
EXCEL_FILE = 'stock.xlsx'

def read_excel():
    try:
        return pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        return pd.DataFrame()

def write_excel(df):
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    df = read_excel()
    headers = df.columns.values
    data = df.values.tolist()
    return render_template('list.html', headers=headers, data=data)

@app.route('/add', methods=['GET', 'POST'])
def add_item():
    if request.method == 'POST':
        df = read_excel()
        new_item = pd.DataFrame([request.form.to_dict()])
        df = pd.concat([df, new_item], ignore_index=True)
        write_excel(df)
        return redirect(url_for('index'))
    return render_template('add.html')

@app.route('/edit/<int:index>', methods=['GET', 'POST'])
def edit_item(index):
    df = read_excel()
    item = df.iloc[index].to_dict()

    if request.method == 'POST':
        for key, value in request.form.items():
            df.loc[index, key] = value
        write_excel(df)
        return redirect(url_for('index'))

    return render_template('edit.html', item=item, index=index)

@app.route('/delete/<int:index>')
def delete_item(index):
    df = read_excel()
    df = df.drop(df.index[index])
    write_excel(df)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
