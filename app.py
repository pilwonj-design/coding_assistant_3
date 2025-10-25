
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

# Excel file path
EXCEL_FILE = os.path.join(os.path.dirname(__file__), 'stock.xlsx')

def read_excel():
    if not os.path.exists(EXCEL_FILE):
        # If xlsx doesn't exist, create it from csv
        csv_file = os.path.join(os.path.dirname(__file__), 'stock.csv')
        if os.path.exists(csv_file):
            df = pd.read_csv(csv_file)
            df.to_excel(EXCEL_FILE, index=False)
            return df
        return pd.DataFrame()
    return pd.read_excel(EXCEL_FILE)

def write_excel(df):
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    df = read_excel()
    return render_template('list.html', tables=[df.to_html(classes='data', header="true", index=False)])

@app.route('/add', methods=['GET', 'POST'])
def add():
    if request.method == 'POST':
        df = read_excel()
        new_data = {
            'Ticker': [request.form['Ticker']],
            'Company': [request.form['Company']],
            'Price': [float(request.form['Price'])],
            'Market Cap (B)': [float(request.form['Market Cap (B)'])],
            'P/E Ratio': [float(request.form['P/E Ratio'])]
        }
        new_df = pd.DataFrame(new_data)
        df = pd.concat([df, new_df], ignore_index=True)
        write_excel(df)
        return redirect(url_for('index'))
    return render_template('add.html')

@app.route('/edit/<int:index>', methods=['GET', 'POST'])
def edit(index):
    df = read_excel()
    item = df.iloc[index].to_dict()

    if request.method == 'POST':
        # Update the dataframe
        df.loc[index, 'Ticker'] = request.form['Ticker']
        df.loc[index, 'Company'] = request.form['Company']
        df.loc[index, 'Price'] = float(request.form['Price'])
        df.loc[index, 'Market Cap (B)'] = float(request.form['Market Cap (B)'])
        df.loc[index, 'P/E Ratio'] = float(request.form['P/E Ratio'])
        
        write_excel(df)
        return redirect(url_for('index'))

    return render_template('edit.html', item=item, index=index)

@app.route('/delete/<int:index>')
def delete(index):
    df = read_excel()
    df = df.drop(df.index[index]).reset_index(drop=True)
    write_excel(df)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
