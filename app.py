from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Ganti 'your_secret_key' dengan kunci rahasia yang kuat

# Routing untuk halaman awal (formulir)
@app.route('/')
def home():
    return render_template('form.html')

# Routing untuk menyimpan data ke dalam file Excel
@app.route('/simpan_ke_excel', methods=['POST'])
def simpan_ke_excel():
    try:
        nama = request.form['nama']
        email = request.form['email']

        # Membaca data yang sudah ada dari file Excel (jika ada)
        try:
            existing_data = pd.read_excel('data.xlsx')
        except FileNotFoundError:
            existing_data = pd.DataFrame()

        # Mengecek apakah data sudah ada dalam file Excel
        if not existing_data.empty and (existing_data['Nama'] == nama).any() and (existing_data['Email'] == email).any():
            flash('Data sudah ada dalam file Excel.')
        else:
            # Data yang ingin Anda masukkan
            new_data = {'Nama': [nama], 'Email': [email]}

            # Menggabungkan data lama dan data baru
            df = pd.concat([existing_data, pd.DataFrame(new_data)], ignore_index=True)

            # Menyimpan data ke dalam file Excel
            df.to_excel('data.xlsx', index=False)
            flash('Data berhasil disimpan ke dalam file Excel.')

        return redirect(url_for('home'))
    except Exception as e:
        flash(str(e))
        return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(debug=True)

# pip install flask pandas ( library yg perlu di install )
# pip install openpyxl ( library yg perlu di install )
# python app.py ( Untuk Run file nya )