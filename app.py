import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user

app = Flask(__name__)
app.secret_key = 'secretkey'  # Для сессий

# Инициализация Flask-Login
login_manager = LoginManager()
login_manager.init_app(app)

# Путь к файлу Excel
file_path = 'wedding_guests.xlsx'

# Заготовка данных для пользователя администратора
class User(UserMixin):
    def __init__(self, id):
        self.id = id

# Простая база пользователей
users = {'admin': {'password': 'adminpass'}}

# Загрузка пользователя для Flask-Login
@login_manager.user_loader
def load_user(user_id):
    return User(user_id)

def save_to_excel(data):
    # Создаем DataFrame из новых данных
    new_data = pd.DataFrame([data])

    # Проверяем, существует ли файл Excel
    if os.path.exists(file_path):
        # Чтение существующего файла Excel
        df = pd.read_excel(file_path)
        # Конкатенируем новые данные с существующими
        df = pd.concat([df, new_data], ignore_index=True)
    else:
        # Если файл не существует, создаем новый DataFrame
        df = new_data

    # Сохранение в Excel
    df.to_excel(file_path, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    full_name = request.form['full_name']
    attending_with = request.form['attending_with']
    
    if attending_with == 'yes':
        second_full_name = request.form['second_full_name']
    else:
        second_full_name = None

    # Сохранение данных в Excel
    save_to_excel({"Full Name": full_name, "Second Full Name": second_full_name})

    return redirect(url_for('index'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        if username in users and users[username]['password'] == password:
            user = User(username)
            login_user(user)
            return redirect(url_for('admin_panel'))
        
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


@app.route('/admin')
@login_required
def admin_panel():
    # Чтение данных из Excel
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)

        # Убираем лишние пробелы и символы
        df.columns = df.columns.str.strip()  # Убираем пробелы из названий колонок
        df['Full Name'] = df['Full Name'].str.strip()  # Убираем пробелы из значений в столбце
        df['Second Full Name'] = df['Second Full Name'].fillna('—').str.strip()  # Заполняем NaN и удаляем пробелы

        # Переименовываем столбцы
        df = df.rename(columns={'Full Name': 'Первый человек', 'Second Full Name': 'Второй человек'})

        # Убираем все строки с ненужными символами, такие как \n и другие
        df = df.replace({r'\n': ' ', r'\s+': ' '}, regex=True)

        # Удаляем пустые строки
        df = df.dropna(subset=['Первый человек'])

        # Подсчет общего числа людей
        total_people = len(df) + df['Второй человек'].apply(lambda x: 1 if x != '—' else 0).sum()

        # Преобразуем таблицу в HTML без лишних символов
        table_html = df.to_html(classes='dataframe data', index=False)

        # Убираем лишние символы, такие как '[]' и лишние кавычки
        table_html = table_html.replace("[", "").replace("]", "").replace("'", "")

        # Убираем лишние символы, такие как \n, если они все еще есть
        table_html = table_html.replace('\n', '')

        # Отправляем результат в шаблон
        return render_template('admin.html', tables=[table_html], title='Admin Panel', total_people=total_people)
    else:
        return "Excel файл не найден."


if __name__ == '__main__':
    app.run(debug=True)
