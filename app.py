from flask import Flask, render_template, request, redirect, url_for, flash, send_file, json
from datetime import datetime, timedelta
from models import db, Project, Stage, Dependency, Resource, Assignment, Photo, PurchaseOrder, PurchaseItem, Act, ActItem, Contract, ContractWorkItem, User
from plan_logic import calculate_dates, find_critical_path
from io import BytesIO
import openpyxl
from openpyxl.styles import Font
import os
import requests
import csv
from io import StringIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
import json
import re

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///construction.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'your-secret-key-here'
db.init_app(app)

# API ключ для OpenWeatherMap
WEATHER_API_KEY = "45e15fb41b97d5a968cae6fdc72e88c5"

# Настройка DeepSeek API
DEEPSEEK_API_KEY = "sk-806b7e601d5348668394d298b28f3bfd"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1"

# Настройка OpenAI (совместимость с DeepSeek)
import openai
openai.api_key = DEEPSEEK_API_KEY
openai.api_base = DEEPSEEK_API_URL

# Хранилище сгенерированных этапов для сессии
generated_stages_cache = {}

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

with app.app_context():
    db.create_all()

@app.template_filter('strftime')
def _jinja2_filter_datetime(date, fmt=None):
    if date is None:
        return ''
    return date.strftime('%Y-%m-%d')

# Функция для получения просроченных этапов
def get_overdue_stages():
    now = datetime.now().date()
    overdue_stages = []
    
    all_stages = Stage.query.all()
    for stage in all_stages:
        if stage.percent_complete < 100:
            dates = calculate_dates(stage.project, db)
            if dates.get(stage.id):
                planned_end = dates[stage.id][1]
                if planned_end < now:
                    overdue_stages.append({
                        'stage': stage,
                        'project': stage.project,
                        'planned_end': planned_end,
                        'days_overdue': (now - planned_end).days
                    })
    
    return overdue_stages

# Функция для получения погоды по координатам
def get_weather_by_coords(lat, lon):
    try:
        url = f"http://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&units=metric&lang=ru&appid={WEATHER_API_KEY}"
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return {
                'temp': round(data['main']['temp']),
                'feels_like': round(data['main']['feels_like']),
                'humidity': data['main']['humidity'],
                'description': data['weather'][0]['description'],
                'icon': data['weather'][0]['icon'],
                'wind_speed': data['wind']['speed'],
                'city': data.get('name', 'Объект'),
                'lat': lat,
                'lon': lon
            }
        return None
    except Exception as e:
        print(f"Ошибка получения погоды: {e}")
        return None

# Функция для расчёта S-кривой
def calculate_s_curve(project):
    stages = Stage.query.filter_by(project_id=project.id).all()
    if not stages:
        return {'dates': [], 'planned': [], 'actual': []}
    
    dates = calculate_dates(project, db)
    if not dates:
        return {'dates': [], 'planned': [], 'actual': []}
    
    min_date = min(start for start, _ in dates.values())
    max_date = max(end for _, end in dates.values())
    
    progress_by_day = {}
    current = min_date
    while current <= max_date:
        progress_by_day[current] = {'planned': 0, 'actual': 0}
        current += timedelta(days=1)
    
    for stage in stages:
        if stage.id not in dates:
            continue
        
        start_date, end_date = dates[stage.id]
        duration = (end_date - start_date).days + 1
        planned_daily = stage.percent_complete / duration if duration > 0 else 0
        actual_daily = stage.percent_complete / duration if duration > 0 else 0
        
        current = start_date
        while current <= end_date:
            if current in progress_by_day:
                progress_by_day[current]['planned'] += planned_daily
                if stage.actual_end_date and current <= stage.actual_end_date:
                    progress_by_day[current]['actual'] += actual_daily
            current += timedelta(days=1)
    
    dates_list = sorted(progress_by_day.keys())
    planned_curve = []
    actual_curve = []
    planned_sum = 0
    actual_sum = 0
    
    for date in dates_list:
        planned_sum += progress_by_day[date]['planned']
        actual_sum += progress_by_day[date]['actual']
        planned_curve.append(min(planned_sum, 100))
        actual_curve.append(min(actual_sum, 100))
    
    return {
        'dates': [d.strftime('%d.%m.%Y') for d in dates_list],
        'planned': planned_curve,
        'actual': actual_curve
    }

# Функция для определения материалов по этапу
def get_materials_for_stage(stage_name, budget):
    stage_lower = stage_name.lower()
    materials = []
    
    if 'фундамент' in stage_lower:
        materials.append({
            'name': 'Бетон М300',
            'quantity': budget / 5000,
            'unit': 'м³',
            'unit_price': 5000,
            'total_price': budget * 0.6,
            'supplier': 'БетонСтрой'
        })
        materials.append({
            'name': 'Арматура',
            'quantity': budget / 20000,
            'unit': 'т',
            'unit_price': 20000,
            'total_price': budget * 0.4,
            'supplier': 'МеталлТорг'
        })
    elif 'стены' in stage_lower or 'кирпич' in stage_lower:
        materials.append({
            'name': 'Газоблок',
            'quantity': budget / 6000,
            'unit': 'м³',
            'unit_price': 6000,
            'total_price': budget * 0.7,
            'supplier': 'СтройБлок'
        })
        materials.append({
            'name': 'Клей для блоков',
            'quantity': budget / 20000,
            'unit': 'мешок',
            'unit_price': 400,
            'total_price': budget * 0.3,
            'supplier': 'СтройБлок'
        })
    elif 'кровля' in stage_lower:
        materials.append({
            'name': 'Металлочерепица',
            'quantity': budget / 800,
            'unit': 'м²',
            'unit_price': 800,
            'total_price': budget * 0.8,
            'supplier': 'КровляПро'
        })
        materials.append({
            'name': 'Гидроизоляция',
            'quantity': budget / 500,
            'unit': 'рулон',
            'unit_price': 2000,
            'total_price': budget * 0.2,
            'supplier': 'КровляПро'
        })
    elif 'фасад' in stage_lower:
        materials.append({
            'name': 'Утеплитель',
            'quantity': budget / 400,
            'unit': 'м²',
            'unit_price': 400,
            'total_price': budget * 0.6,
            'supplier': 'ТеплоСтрой'
        })
        materials.append({
            'name': 'Штукатурка фасадная',
            'quantity': budget / 800,
            'unit': 'мешок',
            'unit_price': 500,
            'total_price': budget * 0.4,
            'supplier': 'ТеплоСтрой'
        })
    elif 'штукатурка' in stage_lower:
        materials.append({
            'name': 'Штукатурная смесь',
            'quantity': budget / 400,
            'unit': 'мешок',
            'unit_price': 400,
            'total_price': budget,
            'supplier': 'ОтделкаСтрой'
        })
    elif 'стяжка' in stage_lower:
        materials.append({
            'name': 'Цемент',
            'quantity': budget / 500,
            'unit': 'мешок',
            'unit_price': 500,
            'total_price': budget * 0.6,
            'supplier': 'ЦементТрейд'
        })
        materials.append({
            'name': 'Песок',
            'quantity': budget / 2000,
            'unit': 'м³',
            'unit_price': 2000,
            'total_price': budget * 0.4,
            'supplier': 'ЦементТрейд'
        })
    elif 'отделка' in stage_lower or 'покраска' in stage_lower:
        materials.append({
            'name': 'Краска',
            'quantity': budget / 800,
            'unit': 'ведро',
            'unit_price': 800,
            'total_price': budget * 0.5,
            'supplier': 'КраскиМир'
        })
        materials.append({
            'name': 'Обои',
            'quantity': budget / 500,
            'unit': 'рулон',
            'unit_price': 500,
            'total_price': budget * 0.5,
            'supplier': 'КраскиМир'
        })
    elif 'плитка' in stage_lower or 'пол' in stage_lower:
        materials.append({
            'name': 'Плитка',
            'quantity': budget / 1000,
            'unit': 'м²',
            'unit_price': 1000,
            'total_price': budget,
            'supplier': 'ПлиткаСтрой'
        })
    elif 'окна' in stage_lower or 'двери' in stage_lower:
        materials.append({
            'name': 'Окна ПВХ',
            'quantity': 1,
            'unit': 'компл',
            'unit_price': budget,
            'total_price': budget,
            'supplier': 'ОкнаПро'
        })
    else:
        materials.append({
            'name': 'Строительные материалы',
            'quantity': 1,
            'unit': 'компл',
            'unit_price': budget,
            'total_price': budget,
            'supplier': 'СтройМаркет'
        })
    
    return materials

# Генерация этапов на основе бюджета (только загородный дом)
def generate_stages_by_budget(budget, quality):
    quality_mult = {
        'econom': 0.7,
        'standard': 1.0,
        'premium': 1.5
    }
    mult = quality_mult.get(quality, 1.0)
    
    stage_percents = [
        ('Геодезические работы', 1, None, 'Геодезист', 1),
        ('Подготовка участка', 2, 'Геодезические работы', 'Бригада 1', 2),
        ('Земляные работы (котлован)', 4, 'Подготовка участка', 'Бригада 1', 4),
        ('Устройство подушки и опалубки', 3, 'Земляные работы (котлован)', 'Бригада 1', 3),
        ('Армирование и заливка фундамента', 6, 'Устройство подушки и опалубки', 'Бригада 1', 10),
        ('Гидроизоляция фундамента', 2, 'Армирование и заливка фундамента', 'Бригада 1', 2),
        ('Возведение стен', 10, 'Гидроизоляция фундамента', 'Бригада 2', 15),
        ('Монтаж перекрытий', 4, 'Возведение стен', 'Бригада 2', 5),
        ('Кровля', 7, 'Монтаж перекрытий', 'Бригада 3', 10),
        ('Установка окон и дверей', 4, 'Возведение стен', 'Бригада 4', 5),
        ('Фасадные работы', 8, 'Установка окон и дверей', 'Бригада 5', 8),
        ('Электромонтаж', 6, 'Возведение стен', 'Бригада 6', 5),
        ('Сантехника и отопление', 6, 'Возведение стен', 'Бригада 7', 6),
        ('Штукатурка', 7, 'Электромонтаж, Сантехника и отопление', 'Бригада 8', 6),
        ('Стяжка полов', 4, 'Штукатурка', 'Бригада 8', 3),
        ('Чистовая отделка', 12, 'Стяжка полов', 'Бригада 9', 10),
        ('Благоустройство', 6, 'Фасадные работы', 'Бригада 10', 4),
        ('Итоговая уборка', 2, 'Чистовая отделка, Благоустройство', 'Бригада 11', 1)
    ]
    
    stages = []
    remaining_budget = budget
    
    for name, duration, depends_on, resources, percent in stage_percents:
        stage_budget = budget * (percent / 100) * mult
        if stage_budget > remaining_budget:
            stage_budget = remaining_budget * 0.8
        
        stages.append({
            'name': name,
            'duration': duration,
            'depends_on': depends_on,
            'resources': resources,
            'budget': round(stage_budget, 0),
            'percent': percent
        })
    
    total = sum(s['budget'] for s in stages)
    if total > budget:
        ratio = budget / total
        for stage in stages:
            stage['budget'] = round(stage['budget'] * ratio, 0)
    
    return stages

def generate_stages_from_ai(message, project):
    """Генерация этапов на основе запроса через DeepSeek"""
    prompt = f"""На основе запроса: "{message}"
Составь JSON-массив этапов для проекта "{project.name}".
Каждый этап должен содержать:
- name: название этапа
- duration: длительность в днях (число)
- depends_on: название этапа, от которого зависит (или null)
- resources: название бригады

Верни только JSON, без пояснений.
Пример: [{{"name": "Фундамент", "duration": 7, "depends_on": null, "resources": "Бригада 1"}}]"""
    
    try:
        response = openai.ChatCompletion.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=2000
        )
        
        text = response.choices[0].message.content
        json_match = re.search(r'\[.*\]', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        return []
    except Exception as e:
        print(f"Ошибка генерации этапов: {e}")
        return []

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        email = request.form['email']
        password = request.form['password']
        confirm = request.form['confirm_password']
        
        if password != confirm:
            flash('Пароли не совпадают', 'error')
            return redirect(url_for('register'))
        
        if len(password) < 6:
            flash('Пароль должен быть не менее 6 символов', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(username=username).first():
            flash('Имя пользователя уже занято', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(email=email).first():
            flash('Email уже зарегистрирован', 'error')
            return redirect(url_for('register'))
        
        user = User(username=username, email=email)
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        
        flash('Регистрация успешна! Войдите в систему.', 'success')
        return redirect(url_for('login'))
    
    return render_template('register.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        
        user = User.query.filter_by(email=email).first()
        
        if user and user.check_password(password):
            login_user(user)
            flash(f'Добро пожаловать, {user.username}!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Неверный email или пароль', 'error')
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Вы вышли из системы', 'success')
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    projects = Project.query.filter_by(user_id=current_user.id).all()
    return render_template('index.html', projects=projects)

@app.route('/add_project', methods=['POST'])
@login_required
def add_project():
    name = request.form['name']
    start_date = datetime.strptime(request.form['start_date'], '%Y-%m-%d').date()
    address = request.form.get('address', '')
    
    lat_str = request.form.get('latitude', '')
    lon_str = request.form.get('longitude', '')
    
    if lat_str and lon_str:
        try:
            latitude = float(lat_str)
            longitude = float(lon_str)
        except ValueError:
            latitude = 55.751244
            longitude = 37.618423
    else:
        latitude = 55.751244
        longitude = 37.618423
    
    project = Project(
        name=name, 
        start_date=start_date,
        address=address,
        latitude=latitude,
        longitude=longitude,
        user_id=current_user.id
    )
    db.session.add(project)
    db.session.commit()
    flash('Проект создан', 'success')
    return redirect(url_for('index'))

@app.route('/delete_project/<int:project_id>')
@login_required
def delete_project(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    db.session.delete(project)
    db.session.commit()
    flash('Проект удалён', 'success')
    return redirect(url_for('index'))

@app.route('/update_project_location/<int:project_id>', methods=['POST'])
@login_required
def update_project_location(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        return json.dumps({'success': False, 'error': 'Access denied'})
    data = request.get_json()
    
    project.latitude = data['latitude']
    project.longitude = data['longitude']
    db.session.commit()
    
    return json.dumps({'success': True})

@app.route('/update_project_address/<int:project_id>', methods=['POST'])
@login_required
def update_project_address(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('project_view', project_id=project_id))
    project.address = request.form.get('address', '')
    
    try:
        project.latitude = float(request.form.get('latitude', project.latitude))
        project.longitude = float(request.form.get('longitude', project.longitude))
    except ValueError:
        pass
    
    db.session.commit()
    flash('Местоположение обновлено', 'success')
    return redirect(url_for('project_view', project_id=project_id))

@app.route('/api/weather')
def api_weather():
    lat = request.args.get('lat', type=float)
    lon = request.args.get('lon', type=float)
    
    if lat is None or lon is None:
        return json.jsonify({'success': False, 'error': 'No coordinates'})
    
    weather = get_weather_by_coords(lat, lon)
    if weather:
        return json.jsonify({'success': True, **weather})
    return json.jsonify({'success': False, 'error': 'Weather not found'})

@app.route('/dashboard')
@login_required
def dashboard():
    projects = Project.query.filter_by(user_id=current_user.id).all()
    stats = []
    
    overdue_stages = get_overdue_stages()
    overdue_stages = [os for os in overdue_stages if os['project'].user_id == current_user.id]
    
    selected_project_id = request.args.get('project_id', type=int)
    selected_project = None
    weather = None
    
    if selected_project_id:
        selected_project = Project.query.get(selected_project_id)
        if selected_project and selected_project.user_id == current_user.id:
            weather = get_weather_by_coords(selected_project.latitude, selected_project.longitude)
    
    for project in projects:
        stages = Stage.query.filter_by(project_id=project.id).all()
        total_stages = len(stages)
        completed_stages = len([s for s in stages if s.percent_complete >= 100])
        dates = calculate_dates(project, db)
        
        delayed = 0
        for stage in stages:
            if stage.percent_complete < 100 and dates.get(stage.id):
                if datetime.now().date() > dates[stage.id][1]:
                    delayed += 1
        
        total_planned = sum(s.planned_material_cost + s.planned_labor_cost for s in stages)
        total_actual = sum(s.actual_material_cost + s.actual_labor_cost for s in stages)
        
        stats.append({
            'project': project,
            'total_stages': total_stages,
            'completed_stages': completed_stages,
            'delayed_stages': delayed,
            'progress': int(completed_stages / total_stages * 100) if total_stages > 0 else 0,
            'budget': {
                'planned': total_planned,
                'actual': total_actual,
                'diff': total_actual - total_planned
            }
        })
    
    return render_template('dashboard.html', 
                         stats=stats, 
                         overdue_stages=overdue_stages, 
                         weather=weather,
                         selected_project=selected_project,
                         projects=projects)

@app.route('/executive_dashboard')
@login_required
def executive_dashboard():
    projects = Project.query.filter_by(user_id=current_user.id).all()
    
    total_budget = 0
    total_spent = 0
    total_stages = 0
    completed_stages = 0
    overdue_stages = 0
    
    project_stats = []
    
    for project in projects:
        stages = Stage.query.filter_by(project_id=project.id).all()
        dates = calculate_dates(project, db)
        
        project_budget = sum(s.planned_material_cost + s.planned_labor_cost for s in stages)
        project_spent = sum(s.actual_material_cost + s.actual_labor_cost for s in stages)
        project_total_stages = len(stages)
        project_completed = len([s for s in stages if s.percent_complete >= 100])
        
        overdue = 0
        for stage in stages:
            if stage.percent_complete < 100 and dates.get(stage.id):
                if datetime.now().date() > dates[stage.id][1]:
                    overdue += 1
        
        total_budget += project_budget
        total_spent += project_spent
        total_stages += project_total_stages
        completed_stages += project_completed
        overdue_stages += overdue
        
        completion_date = None
        if stages and dates:
            max_date = max(end for _, end in dates.values())
            completion_date = max_date
        
        project_stats.append({
            'project': project,
            'budget': project_budget,
            'spent': project_spent,
            'total_stages': project_total_stages,
            'completed_stages': project_completed,
            'overdue': overdue,
            'completion_date': completion_date,
            'progress': int(project_completed / project_total_stages * 100) if project_total_stages > 0 else 0,
            'budget_usage': int(project_spent / project_budget * 100) if project_budget > 0 else 0
        })
    
    overall_progress = int(completed_stages / total_stages * 100) if total_stages > 0 else 0
    budget_usage = int(total_spent / total_budget * 100) if total_budget > 0 else 0
    
    return render_template('executive_dashboard.html',
                         projects=projects,
                         project_stats=project_stats,
                         total_projects=len(projects),
                         total_budget=total_budget,
                         total_spent=total_spent,
                         total_stages=total_stages,
                         completed_stages=completed_stages,
                         overdue_stages=overdue_stages,
                         overall_progress=overall_progress,
                         budget_usage=budget_usage)

@app.route('/api/ai/chat', methods=['POST'])
@login_required
def ai_chat():
    data = request.get_json()
    message = data.get('message', '')
    project_id = data.get('project_id')
    
    project = Project.query.get(project_id)
    
    context = f"Проект: {project.name}\nДата начала: {project.start_date}\n"
    stages = Stage.query.filter_by(project_id=project_id).all()
    if stages:
        context += f"Всего этапов: {len(stages)}\n"
        completed = len([s for s in stages if s.percent_complete >= 100])
        context += f"Выполнено этапов: {completed}\n"
    
    prompt = f"""Ты профессиональный строительный консультант. 
Контекст проекта:
{context}

Вопрос пользователя: {message}

Ответь полезно и конкретно. Если просят смету, предложи структуру этапов.
Если вопрос не по строительству, вежливо скажи, что ты строительный ассистент."""
    
    try:
        response = openai.ChatCompletion.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "Ты профессиональный строительный консультант. Отвечай кратко и по делу."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=1000
        )
        
        answer = response.choices[0].message.content
        
        stages = []
        if 'смет' in message.lower() or 'этап' in message.lower():
            stages = generate_stages_from_ai(message, project)
            generated_stages_cache[project_id] = stages
        
        return json.jsonify({
            'success': True,
            'response': answer,
            'stages': stages
        })
        
    except Exception as e:
        print(f"Ошибка DeepSeek: {e}")
        return json.jsonify({'success': False, 'error': str(e)})

@app.route('/api/ai/add_stages/<int:project_id>', methods=['POST'])
@login_required
def add_ai_stages(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        return json.jsonify({'success': False, 'error': 'Access denied'})
    
    stages_data = generated_stages_cache.get(project_id, [])
    if not stages_data:
        return json.jsonify({'success': False, 'error': 'No generated stages'})
    
    added = 0
    stage_names = {s.name: s.id for s in Stage.query.filter_by(project_id=project_id).all()}
    
    for stage_data in stages_data:
        name = stage_data.get('name')
        if name in stage_names:
            continue
        
        stage = Stage(
            name=name,
            planned_duration=stage_data.get('duration', 7),
            project_id=project_id,
            percent_complete=0
        )
        db.session.add(stage)
        db.session.flush()
        stage_names[name] = stage.id
        added += 1
        
        if stage_data.get('resources'):
            resource = Resource.query.filter_by(name=stage_data['resources']).first()
            if not resource:
                resource = Resource(name=stage_data['resources'])
                db.session.add(resource)
                db.session.flush()
            assign = Assignment(stage_id=stage.id, resource_id=resource.id)
            db.session.add(assign)
        
        depends_on = stage_data.get('depends_on')
        if depends_on and depends_on in stage_names:
            dep = Dependency(stage_id=stage.id, depends_on_stage_id=stage_names[depends_on])
            db.session.add(dep)
    
    db.session.commit()
    generated_stages_cache[project_id] = []
    
    return json.jsonify({'success': True, 'count': added})

@app.route('/api/generate_estimate_by_budget', methods=['POST'])
@login_required
def generate_estimate_by_budget():
    data = request.get_json()
    budget = data.get('budget')
    quality = data.get('quality', 'standard')
    
    if not budget or budget < 50000:
        return json.jsonify({'success': False, 'error': 'Бюджет должен быть не менее 50 000 ₽'})
    
    stages = generate_stages_by_budget(budget, quality)
    
    return json.jsonify({'success': True, 'stages': stages})

@app.route('/api/add_estimate_stages/<int:project_id>', methods=['POST'])
@login_required
def add_estimate_stages(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        return json.jsonify({'success': False, 'error': 'Access denied'})
    
    data = request.get_json()
    stages_data = data.get('stages', [])
    
    added = 0
    existing_stages = {s.name: s.id for s in Stage.query.filter_by(project_id=project_id).all()}
    
    for stage_data in stages_data:
        name = stage_data.get('name')
        if name in existing_stages:
            continue
        
        duration = stage_data.get('duration', 7)
        
        stage = Stage(
            name=name,
            planned_duration=duration,
            project_id=project_id,
            planned_material_cost=round(stage_data.get('budget', 0) * 0.6, 0),
            planned_labor_cost=round(stage_data.get('budget', 0) * 0.4, 0),
            percent_complete=0
        )
        db.session.add(stage)
        db.session.flush()
        existing_stages[name] = stage.id
        added += 1
        
        resources = stage_data.get('resources', '')
        if resources and resources != '—':
            for r_name in resources.split(', '):
                resource = Resource.query.filter_by(name=r_name).first()
                if not resource:
                    resource = Resource(name=r_name)
                    db.session.add(resource)
                    db.session.flush()
                assign = Assignment(stage_id=stage.id, resource_id=resource.id)
                db.session.add(assign)
        
        depends_on = stage_data.get('depends_on')
        if depends_on and depends_on != '—' and depends_on in existing_stages:
            dep = Dependency(stage_id=stage.id, depends_on_stage_id=existing_stages[depends_on])
            db.session.add(dep)
    
    db.session.commit()
    
    return json.jsonify({'success': True, 'count': added})

@app.route('/s_curve/<int:project_id>')
@login_required
def s_curve(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id).all()
    dates = calculate_dates(project, db)
    critical_stages = find_critical_path(project, db)
    
    project_budget = sum(s.planned_material_cost + s.planned_labor_cost for s in stages)
    project_spent = sum(s.actual_material_cost + s.actual_labor_cost for s in stages)
    
    return render_template('s_curve.html', 
                         project=project, 
                         stages=stages,
                         dates=dates,
                         critical_stages=critical_stages,
                         project_budget=project_budget,
                         project_spent=project_spent)

@app.route('/project/<int:project_id>')
@login_required
def project_view(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id).all()
    dates = calculate_dates(project, db)
    critical_stages = find_critical_path(project, db)
    assignments = {}
    for stage in stages:
        ass = Assignment.query.filter_by(stage_id=stage.id).all()
        assignments[stage.id] = [r.resource.name for r in ass]
    resources = Resource.query.all()
    
    now = datetime.now().date()
    overdue_in_project = []
    for stage in stages:
        if stage.percent_complete < 100 and dates.get(stage.id):
            planned_end = dates[stage.id][1]
            if planned_end < now:
                overdue_in_project.append({
                    'stage': stage,
                    'planned_end': planned_end,
                    'days_overdue': (now - planned_end).days
                })
    
    return render_template('project.html',
                           project=project,
                           stages=stages,
                           dates=dates,
                           assignments=assignments,
                           resources=resources,
                           critical_stages=critical_stages,
                           overdue_in_project=overdue_in_project)

@app.route('/update_completion_ajax/<int:project_id>', methods=['POST'])
@login_required
def update_completion_ajax(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        return json.jsonify({'success': False, 'error': 'Access denied'})
    
    data = request.get_json()
    stage_id = data.get('stage_id')
    percent = int(data.get('percent'))
    
    stage = Stage.query.get_or_404(stage_id)
    if stage.project_id != project_id:
        return json.jsonify({'success': False, 'error': 'Stage not in project'})
    
    stage.percent_complete = percent
    if percent >= 100:
        stage.actual_end_date = datetime.now().date()
    else:
        stage.actual_end_date = None
    
    db.session.commit()
    
    return json.jsonify({'success': True})

@app.route('/get_project_progress/<int:project_id>')
@login_required
def get_project_progress(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        return json.jsonify({'success': False, 'error': 'Access denied'})
    
    stages = Stage.query.filter_by(project_id=project_id).all()
    total_stages = len(stages)
    completed_stages = len([s for s in stages if s.percent_complete >= 100])
    overall_progress = int(completed_stages / total_stages * 100) if total_stages > 0 else 0
    
    return json.jsonify({'success': True, 'overall_progress': overall_progress})

@app.route('/get_project_budget/<int:project_id>')
@login_required
def get_project_budget(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        return json.jsonify({'success': False, 'error': 'Access denied'})
    
    stages = Stage.query.filter_by(project_id=project_id).all()
    total_planned = sum(s.planned_material_cost + s.planned_labor_cost for s in stages)
    total_actual = sum(s.actual_material_cost + s.actual_labor_cost for s in stages)
    
    return json.jsonify({
        'success': True,
        'total_planned': total_planned,
        'total_actual': total_actual
    })

@app.route('/project_calendar/<int:project_id>')
@login_required
def project_calendar(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id).all()
    dates = calculate_dates(project, db)
    assignments = {}
    for stage in stages:
        ass = Assignment.query.filter_by(stage_id=stage.id).all()
        assignments[stage.id] = [r.resource.name for r in ass]
    
    return render_template('project_calendar.html',
                         project=project,
                         stages=stages,
                         dates=dates,
                         assignments=assignments)

@app.route('/purchase_plan/<int:project_id>')
@login_required
def purchase_plan(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    purchase_orders = PurchaseOrder.query.filter_by(project_id=project_id).order_by(PurchaseOrder.order_date.desc()).all()
    return render_template('purchase_plan.html', project=project, purchase_orders=purchase_orders)

@app.route('/generate_purchase_plan/<int:project_id>')
@login_required
def generate_purchase_plan(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id).all()
    dates = calculate_dates(project, db)
    
    materials_by_category = {}
    
    for stage in stages:
        if stage.planned_material_cost > 0 and dates.get(stage.id):
            materials = get_materials_for_stage(stage.name, stage.planned_material_cost)
            
            for material in materials:
                category = material['supplier'] or 'Прочие'
                if category not in materials_by_category:
                    materials_by_category[category] = []
                
                materials_by_category[category].append({
                    'stage': stage,
                    'material_name': material['name'],
                    'quantity': material['quantity'],
                    'unit': material['unit'],
                    'unit_price': material['unit_price'],
                    'total_price': material['total_price'],
                    'delivery_date': dates[stage.id][0] - timedelta(days=7)
                })
    
    idx = 0
    for supplier, items in materials_by_category.items():
        idx += 1
        order_number = f"PO-{project_id}-{datetime.now().strftime('%Y%m%d')}-{idx}"
        
        order = PurchaseOrder(
            project_id=project_id,
            order_number=order_number,
            supplier=supplier if supplier != 'Прочие' else None,
            order_date=datetime.now().date(),
            total_amount=sum(item['total_price'] for item in items),
            status='pending'
        )
        db.session.add(order)
        db.session.flush()
        
        for item in items:
            purchase_item = PurchaseItem(
                order_id=order.id,
                stage_id=item['stage'].id,
                material_name=item['material_name'],
                quantity=item['quantity'],
                unit=item['unit'],
                unit_price=item['unit_price'],
                total_price=item['total_price'],
                status='pending'
            )
            db.session.add(purchase_item)
    
    db.session.commit()
    flash('План закупок сформирован!', 'success')
    return redirect(url_for('purchase_plan', project_id=project_id))

@app.route('/update_order_status/<int:order_id>/<status>')
@login_required
def update_order_status(order_id, status):
    order = PurchaseOrder.query.get_or_404(order_id)
    project = order.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    order.status = status
    
    for item in order.items:
        if status == 'ordered':
            item.status = 'ordered'
        elif status == 'delivered':
            item.status = 'received'
    
    db.session.commit()
    flash(f'Статус заказа №{order.order_number} обновлён на "{status}"', 'success')
    return redirect(url_for('purchase_plan', project_id=order.project_id))

@app.route('/delete_order/<int:order_id>')
@login_required
def delete_order(order_id):
    order = PurchaseOrder.query.get_or_404(order_id)
    project = order.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    project_id = order.project_id
    db.session.delete(order)
    db.session.commit()
    flash('Заказ удалён', 'success')
    return redirect(url_for('purchase_plan', project_id=project_id))

@app.route('/acts/<int:project_id>')
@login_required
def acts(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    acts = Act.query.filter_by(project_id=project_id).order_by(Act.act_date.desc()).all()
    return render_template('acts.html', project=project, acts=acts)

@app.route('/generate_act/<int:project_id>')
@login_required
def generate_act(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id, percent_complete=100).all()
    
    if not stages:
        flash('Нет завершённых этапов для формирования акта', 'error')
        return redirect(url_for('acts', project_id=project_id))
    
    act_number = f"КС-2-{project_id}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    
    act = Act(
        project_id=project_id,
        act_number=act_number,
        act_date=datetime.now().date(),
        period_start=datetime.now().date() - timedelta(days=30),
        period_end=datetime.now().date(),
        total_amount=0,
        status='draft'
    )
    db.session.add(act)
    db.session.flush()
    
    total = 0
    for stage in stages:
        amount = stage.planned_material_cost + stage.planned_labor_cost
        total += amount
        
        act_item = ActItem(
            act_id=act.id,
            stage_id=stage.id,
            work_name=stage.name,
            quantity=1,
            unit='компл',
            unit_price=amount,
            total_price=amount
        )
        db.session.add(act_item)
    
    act.total_amount = total
    db.session.commit()
    
    flash(f'Акт №{act_number} сформирован!', 'success')
    return redirect(url_for('acts', project_id=project_id))

@app.route('/sign_act/<int:act_id>')
@login_required
def sign_act(act_id):
    act = Act.query.get_or_404(act_id)
    project = act.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    act.status = 'signed'
    db.session.commit()
    flash('Акт подписан', 'success')
    return redirect(url_for('acts', project_id=act.project_id))

@app.route('/delete_act/<int:act_id>')
@login_required
def delete_act(act_id):
    act = Act.query.get_or_404(act_id)
    project = act.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    project_id = act.project_id
    db.session.delete(act)
    db.session.commit()
    flash('Акт удалён', 'success')
    return redirect(url_for('acts', project_id=project_id))

@app.route('/contracts/<int:project_id>')
@login_required
def contracts(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    contracts = Contract.query.filter_by(project_id=project_id).order_by(Contract.contract_date.desc()).all()
    return render_template('contracts.html', project=project, contracts=contracts)

@app.route('/add_contract/<int:project_id>', methods=['GET', 'POST'])
@login_required
def add_contract(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        contract = Contract(
            project_id=project_id,
            contract_number=request.form['contract_number'],
            contract_date=datetime.strptime(request.form['contract_date'], '%Y-%m-%d').date(),
            contractor_name=request.form['contractor_name'],
            contractor_inn=request.form.get('contractor_inn', ''),
            contractor_kpp=request.form.get('contractor_kpp', ''),
            contractor_ogrn=request.form.get('contractor_ogrn', ''),
            contractor_address=request.form.get('contractor_address', ''),
            contractor_fact_address=request.form.get('contractor_fact_address', ''),
            contractor_phone=request.form.get('contractor_phone', ''),
            contractor_email=request.form.get('contractor_email', ''),
            contractor_director=request.form.get('contractor_director', ''),
            contractor_basis=request.form.get('contractor_basis', ''),
            contractor_bank_name=request.form.get('contractor_bank_name', ''),
            contractor_bik=request.form.get('contractor_bik', ''),
            contractor_account=request.form.get('contractor_account', ''),
            contractor_correspondent_account=request.form.get('contractor_correspondent_account', ''),
            start_date=datetime.strptime(request.form['start_date'], '%Y-%m-%d').date(),
            end_date=datetime.strptime(request.form['end_date'], '%Y-%m-%d').date() if request.form.get('end_date') else None,
            total_amount=float(request.form['total_amount']),
            object_name=request.form.get('object_name', project.name),
            object_address=request.form.get('object_address', project.address),
            subject=request.form.get('subject', ''),
            documentation=request.form.get('documentation', ''),
            doc_deadline=request.form.get('doc_deadline', ''),
            doc_method=request.form.get('doc_method', ''),
            special_conditions=request.form.get('special_conditions', ''),
            liability=request.form.get('liability', ''),
            warranty=request.form.get('warranty', ''),
            status='draft'
        )
        db.session.add(contract)
        db.session.flush()
        
        work_names = request.form.getlist('work_name[]')
        work_units = request.form.getlist('work_unit[]')
        work_qtys = request.form.getlist('work_qty[]')
        work_prices = request.form.getlist('work_price[]')
        work_totals = request.form.getlist('work_total[]')
        
        for i in range(len(work_names)):
            if work_names[i] and work_names[i].strip():
                work_item = ContractWorkItem(
                    contract_id=contract.id,
                    work_name=work_names[i],
                    unit=work_units[i] if i < len(work_units) else 'шт',
                    quantity=float(work_qtys[i]) if i < len(work_qtys) and work_qtys[i] else 0,
                    unit_price=float(work_prices[i]) if i < len(work_prices) and work_prices[i] else 0,
                    total_price=float(work_totals[i]) if i < len(work_totals) and work_totals[i] else 0
                )
                db.session.add(work_item)
        
        db.session.commit()
        flash('Договор добавлен', 'success')
        return redirect(url_for('contracts', project_id=project_id))
    
    return render_template('add_contract.html', project=project)

@app.route('/export_contract_word/<int:contract_id>')
@login_required
def export_contract_word(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    project = contract.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    
    doc = Document()
    
    section = doc.sections[0]
    section.top_margin = Inches(0.79)
    section.bottom_margin = Inches(0.79)
    section.left_margin = Inches(0.79)
    section.right_margin = Inches(0.79)
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    title = doc.add_heading('ДОГОВОР ПОДРЯДА', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(16)
    title.runs[0].font.bold = True
    title.runs[0].font.name = 'Times New Roman'
    
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'№ {contract.contract_number}').bold = True
    doc.add_paragraph()
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f'г. {project.address.split(",")[0] if project.address else "Москва"}').bold = True
    p2.add_run(f' "{contract.contract_date.day}" {contract.contract_date.strftime("%B").lower()} {contract.contract_date.year} г.').bold = True
    
    doc.add_paragraph()
    
    p3 = doc.add_paragraph()
    p3.add_run(f'{project.name}, именуемый в дальнейшем "Заказчик", в лице _______________, действующего на основании _______________, с одной стороны, и ').italic = True
    p3.add_run(f'{contract.contractor_name}, именуемый в дальнейшем "Подрядчик", в лице {contract.contractor_director or "_______________"}, действующего на основании {contract.contractor_basis or "Устава"}, с другой стороны, ').italic = True
    p3.add_run('заключили настоящий договор о нижеследующем:').italic = True
    
    doc.add_paragraph()
    
    doc.add_heading('1. СТОРОНЫ ДОГОВОРА', level=1)
    doc.add_paragraph('1.1. Заказчик:')
    doc.add_paragraph(f'    Полное наименование: {project.name}', style='Normal')
    doc.add_paragraph(f'    Адрес: {project.address or "не указан"}', style='Normal')
    doc.add_paragraph()
    doc.add_paragraph('1.2. Подрядчик:')
    doc.add_paragraph(f'    Полное наименование: {contract.contractor_name}', style='Normal')
    if contract.contractor_inn:
        doc.add_paragraph(f'    ИНН: {contract.contractor_inn}', style='Normal')
    if contract.contractor_kpp:
        doc.add_paragraph(f'    КПП: {contract.contractor_kpp}', style='Normal')
    if contract.contractor_ogrn:
        doc.add_paragraph(f'    ОГРН: {contract.contractor_ogrn}', style='Normal')
    if contract.contractor_address:
        doc.add_paragraph(f'    Юридический адрес: {contract.contractor_address}', style='Normal')
    if contract.contractor_phone:
        doc.add_paragraph(f'    Телефон: {contract.contractor_phone}', style='Normal')
    if contract.contractor_email:
        doc.add_paragraph(f'    Email: {contract.contractor_email}', style='Normal')
    if contract.contractor_bank_name:
        doc.add_paragraph(f'    Банковские реквизиты: {contract.contractor_bank_name}', style='Normal')
        if contract.contractor_bik:
            doc.add_paragraph(f'    БИК: {contract.contractor_bik}', style='Normal')
        if contract.contractor_account:
            doc.add_paragraph(f'    Расчетный счет: {contract.contractor_account}', style='Normal')
        if contract.contractor_correspondent_account:
            doc.add_paragraph(f'    Корреспондентский счет: {contract.contractor_correspondent_account}', style='Normal')
    
    doc.add_paragraph()
    
    doc.add_heading('2. ПРЕДМЕТ ДОГОВОРА', level=1)
    doc.add_paragraph('2.1. Заказчик поручает, а Подрядчик принимает на себя обязательства выполнить следующие работы:', style='Normal')
    doc.add_paragraph()
    
    if contract.subject:
        doc.add_paragraph(contract.subject)
    else:
        doc.add_paragraph(f'Строительные работы на объекте "{project.name}", расположенном по адресу: {project.address or "не указан"}.')
    
    doc.add_paragraph()
    doc.add_paragraph('2.2. Состав и объем работ определяется настоящим договором и приложением №1 (Перечень работ).')
    
    doc.add_page_break()
    doc.add_heading('ПРИЛОЖЕНИЕ №1', level=1)
    doc.add_heading('к Договору подряда № ' + contract.contract_number, level=2)
    doc.add_heading('ПЕРЕЧЕНЬ РАБОТ', level=3)
    
    if contract.work_items:
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        table.autofit = False
        table.allow_autofit = False
        
        table.columns[0].width = Inches(0.5)
        table.columns[1].width = Inches(3.5)
        table.columns[2].width = Inches(0.8)
        table.columns[3].width = Inches(0.8)
        table.columns[4].width = Inches(1.2)
        table.columns[5].width = Inches(1.2)
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '№ п/п'
        hdr_cells[1].text = 'Наименование работ'
        hdr_cells[2].text = 'Ед. изм.'
        hdr_cells[3].text = 'Кол-во'
        hdr_cells[4].text = 'Цена, руб.'
        hdr_cells[5].text = 'Сумма, руб.'
        
        for idx, item in enumerate(contract.work_items, 1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(idx)
            row_cells[1].text = item.work_name
            row_cells[2].text = item.unit
            row_cells[3].text = str(int(item.quantity) if item.quantity.is_integer() else f"{item.quantity:.2f}")
            row_cells[4].text = f"{item.unit_price:,.2f}".replace(',', ' ')
            row_cells[5].text = f"{item.total_price:,.2f}".replace(',', ' ')
        
        total_row = table.add_row()
        total_cells = total_row.cells
        total_cells[0].text = ''
        total_cells[1].text = ''
        total_cells[2].text = ''
        total_cells[3].text = ''
        total_cells[4].text = 'ИТОГО:'
        total_cells[4].paragraphs[0].runs[0].bold = True
        total_cells[5].text = f"{contract.total_amount:,.2f}".replace(',', ' ')
        total_cells[5].paragraphs[0].runs[0].bold = True
    
    doc.add_paragraph()
    doc.add_paragraph(f'Общая стоимость работ по договору составляет: {contract.total_amount:,.2f} ({"{:.2f}".format(contract.total_amount).replace(".", ",")}) рублей.')
    
    doc.add_page_break()
    doc.add_heading('3. ТЕХНИЧЕСКАЯ ДОКУМЕНТАЦИЯ', level=1)
    if contract.documentation:
        doc.add_paragraph(contract.documentation)
    else:
        doc.add_paragraph('3.1. Заказчик передает Подрядчику следующую документацию:')
        doc.add_paragraph('    - Проектная документация (шифр: _____);', style='Normal')
        doc.add_paragraph('    - Рабочая документация (шифр: _____);', style='Normal')
        doc.add_paragraph('    - Сметная документация;', style='Normal')
        doc.add_paragraph('    - График производства работ.', style='Normal')
    
    if contract.doc_deadline:
        doc.add_paragraph(f'3.2. Срок передачи документации: {contract.doc_deadline}')
    if contract.doc_method:
        doc.add_paragraph(f'3.3. Порядок передачи: {contract.doc_method}')
    
    doc.add_paragraph()
    
    doc.add_heading('4. СРОКИ ВЫПОЛНЕНИЯ РАБОТ', level=1)
    doc.add_paragraph(f'4.1. Начало работ: {contract.start_date.strftime("%d.%m.%Y")}')
    if contract.end_date:
        doc.add_paragraph(f'4.2. Окончание работ: {contract.end_date.strftime("%d.%m.%Y")}')
    doc.add_paragraph('4.3. Подрядчик вправе выполнить работы досрочно.')
    
    doc.add_paragraph()
    
    doc.add_heading('5. СТОИМОСТЬ РАБОТ И ПОРЯДОК РАСЧЕТОВ', level=1)
    doc.add_paragraph(f'5.1. Общая стоимость работ по договору составляет {contract.total_amount:,.2f} ({"{:.2f}".format(contract.total_amount).replace(".", ",")}) рублей.')
    
    if contract.special_conditions:
        doc.add_paragraph(f'5.2. Порядок оплаты: {contract.special_conditions}')
    else:
        doc.add_paragraph('5.2. Аванс: 30% от стоимости работ в течение 5 рабочих дней после подписания договора.')
        doc.add_paragraph('5.3. Окончательный расчет: в течение 10 рабочих дней после подписания акта выполненных работ (КС-2).')
    
    doc.add_paragraph()
    
    if contract.liability:
        doc.add_heading('6. ОТВЕТСТВЕННОСТЬ СТОРОН', level=1)
        doc.add_paragraph(contract.liability)
    else:
        doc.add_heading('6. ОТВЕТСТВЕННОСТЬ СТОРОН', level=1)
        doc.add_paragraph('6.1. За нарушение сроков выполнения работ Подрядчик уплачивает неустойку в размере 0,1% от стоимости работ за каждый день просрочки.')
        doc.add_paragraph('6.2. За нарушение сроков оплаты Заказчик уплачивает неустойку в размере 0,1% от суммы задолженности за каждый день просрочки.')
    
    doc.add_paragraph()
    
    if contract.warranty:
        doc.add_heading('7. ГАРАНТИЙНЫЙ СРОК', level=1)
        doc.add_paragraph(contract.warranty)
    else:
        doc.add_heading('7. ГАРАНТИЙНЫЙ СРОК', level=1)
        doc.add_paragraph('7.1. Гарантийный срок на выполненные работы составляет 36 месяцев с даты подписания акта выполненных работ.')
    
    doc.add_paragraph()
    
    doc.add_heading('8. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ', level=1)
    doc.add_paragraph('8.1. Все споры по настоящему договору разрешаются в претензионном порядке. Срок рассмотрения претензии — 10 рабочих дней.')
    doc.add_paragraph('8.2. При не достижении согласия спор передается на рассмотрение в Арбитражный суд по месту нахождения Заказчика.')
    doc.add_paragraph('8.3. Настоящий договор составлен в двух экземплярах, имеющих одинаковую юридическую силу, по одному для каждой из Сторон.')
    
    doc.add_paragraph()
    
    doc.add_heading('9. ПОДПИСИ СТОРОН', level=1)
    doc.add_paragraph()
    
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.5)
    sig_table.columns[1].width = Inches(3.5)
    
    sig_hdr = sig_table.rows[0].cells
    sig_hdr[0].text = 'Заказчик'
    sig_hdr[1].text = 'Подрядчик'
    sig_hdr[0].paragraphs[0].runs[0].bold = True
    sig_hdr[1].paragraphs[0].runs[0].bold = True
    
    sig_row = sig_table.rows[1].cells
    sig_row[0].text = f'\n\n____________________\n\n{project.name}'
    sig_row[1].text = f'\n\n____________________\n\n{contract.contractor_name}'
    
    doc.add_paragraph()
    p_mp = doc.add_paragraph('М.П.')
    p_mp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name=f'Договор_{contract.contract_number}.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/sign_contract/<int:contract_id>')
@login_required
def sign_contract(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    project = contract.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    contract.status = 'signed'
    db.session.commit()
    flash('Договор подписан', 'success')
    return redirect(url_for('contracts', project_id=contract.project_id))

@app.route('/delete_contract/<int:contract_id>')
@login_required
def delete_contract(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    project = contract.project
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    project_id = contract.project_id
    db.session.delete(contract)
    db.session.commit()
    flash('Договор удалён', 'success')
    return redirect(url_for('contracts', project_id=project_id))

@app.route('/export_csv/<int:project_id>')
@login_required
def export_csv(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id).all()
    dates = calculate_dates(project, db)
    
    si = StringIO()
    writer = csv.writer(si, delimiter=';')
    
    writer.writerow(['Название этапа', 'Длительность (дни)', 'Зависимости', 'Ресурсы', 
                     'План начало', 'План конец', 'Выполнено %', 'План материалы (руб)', 
                     'План работы (руб)', 'Факт материалы (руб)', 'Факт работы (руб)'])
    
    for stage in stages:
        deps = [dep.depends_on_stage.name for dep in stage.dependencies.all()]
        resources = [ass.resource.name for ass in stage.assignments]
        
        writer.writerow([
            stage.name,
            stage.planned_duration,
            ', '.join(deps),
            ', '.join(resources),
            dates[stage.id][0].strftime('%d.%m.%Y') if dates.get(stage.id) else '',
            dates[stage.id][1].strftime('%d.%m.%Y') if dates.get(stage.id) else '',
            stage.percent_complete,
            stage.planned_material_cost,
            stage.planned_labor_cost,
            stage.actual_material_cost,
            stage.actual_labor_cost
        ])
    
    output = BytesIO()
    output.write(si.getvalue().encode('utf-8-sig'))
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name=f'project_{project.name}.csv',
        mimetype='text/csv'
    )

@app.route('/project/<int:project_id>/add_stage', methods=['GET', 'POST'])
@login_required
def add_stage(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        name = request.form['name']
        duration = int(request.form['duration'])
        planned_material_cost = float(request.form.get('planned_material_cost', 0))
        planned_labor_cost = float(request.form.get('planned_labor_cost', 0))
        
        stage = Stage(
            name=name,
            planned_duration=duration,
            project_id=project_id,
            planned_material_cost=planned_material_cost,
            planned_labor_cost=planned_labor_cost
        )
        db.session.add(stage)
        db.session.commit()
        
        depends = request.form.getlist('depends')
        for dep_id in depends:
            dep = Dependency(stage_id=stage.id, depends_on_stage_id=int(dep_id))
            db.session.add(dep)
        
        resource_ids = request.form.getlist('resources')
        for rid in resource_ids:
            assign = Assignment(stage_id=stage.id, resource_id=int(rid))
            db.session.add(assign)
        
        db.session.commit()
        flash('Этап добавлен', 'success')
        return redirect(url_for('project_view', project_id=project_id))
    
    existing_stages = Stage.query.filter_by(project_id=project_id).all()
    all_resources = Resource.query.all()
    return render_template('edit_stage.html',
                           project=project,
                           stage=None,
                           existing_stages=existing_stages,
                           all_resources=all_resources)

@app.route('/project/<int:project_id>/edit_stage/<int:stage_id>', methods=['GET', 'POST'])
@login_required
def edit_stage(project_id, stage_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stage = Stage.query.get_or_404(stage_id)
    if stage.project_id != project_id:
        return redirect(url_for('project_view', project_id=project_id))
    
    if request.method == 'POST':
        stage.name = request.form['name']
        stage.planned_duration = int(request.form['duration'])
        
        if 'percent_complete' in request.form:
            stage.percent_complete = int(request.form['percent_complete'])
            if stage.percent_complete >= 100:
                if not stage.actual_end_date:
                    stage.actual_end_date = datetime.now().date()
            else:
                stage.actual_end_date = None
        
        stage.planned_material_cost = float(request.form.get('planned_material_cost', 0))
        stage.planned_labor_cost = float(request.form.get('planned_labor_cost', 0))
        stage.actual_material_cost = float(request.form.get('actual_material_cost', 0))
        stage.actual_labor_cost = float(request.form.get('actual_labor_cost', 0))
        
        if request.form.get('custom_start_date'):
            try:
                stage.custom_start_date = datetime.strptime(request.form['custom_start_date'], '%Y-%m-%d').date()
            except:
                pass
        if request.form.get('custom_end_date'):
            try:
                stage.custom_end_date = datetime.strptime(request.form['custom_end_date'], '%Y-%m-%d').date()
            except:
                pass
        
        Dependency.query.filter_by(stage_id=stage.id).delete()
        depends = request.form.getlist('depends')
        for dep_id in depends:
            dep = Dependency(stage_id=stage.id, depends_on_stage_id=int(dep_id))
            db.session.add(dep)
        
        Assignment.query.filter_by(stage_id=stage.id).delete()
        resource_ids = request.form.getlist('resources')
        for rid in resource_ids:
            assign = Assignment(stage_id=stage.id, resource_id=int(rid))
            db.session.add(assign)
        
        db.session.commit()
        flash('Этап обновлён', 'success')
        return redirect(url_for('project_view', project_id=project_id))
    
    existing_stages = Stage.query.filter_by(project_id=project_id).filter(Stage.id != stage_id).all()
    all_resources = Resource.query.all()
    current_depends = [dep.depends_on_stage_id for dep in Dependency.query.filter_by(stage_id=stage.id).all()]
    current_resources = [ass.resource_id for ass in Assignment.query.filter_by(stage_id=stage.id).all()]
    return render_template('edit_stage.html',
                           project=project,
                           stage=stage,
                           existing_stages=existing_stages,
                           all_resources=all_resources,
                           current_depends=current_depends,
                           current_resources=current_resources)

@app.route('/project/<int:project_id>/delete_stage/<int:stage_id>')
@login_required
def delete_stage(project_id, stage_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stage = Stage.query.get_or_404(stage_id)
    if stage.project_id != project_id:
        return redirect(url_for('project_view', project_id=project_id))
    db.session.delete(stage)
    db.session.commit()
    flash('Этап удалён', 'success')
    return redirect(url_for('project_view', project_id=project_id))

@app.route('/project/<int:project_id>/update_completion', methods=['POST'])
@login_required
def update_completion(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stage_id = request.form.get('stage_id')
    stage = Stage.query.get_or_404(stage_id)
    if stage.project_id != project_id:
        return redirect(url_for('project_view', project_id=project_id))
    percent = int(request.form['percent'])
    stage.percent_complete = percent
    if percent >= 100:
        stage.actual_end_date = datetime.now().date()
    else:
        stage.actual_end_date = None
    db.session.commit()
    flash('Процент выполнения обновлён', 'success')
    return redirect(url_for('project_view', project_id=project_id))

@app.route('/photos/<int:project_id>')
@login_required
def photos_view(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    photos = Photo.query.filter_by(project_id=project_id).order_by(Photo.created_at.desc()).all()
    return render_template('photos.html', project=project, photos=photos)

@app.route('/resource_calendar')
@login_required
def resource_calendar():
    resources = Resource.query.all()
    projects = Project.query.filter_by(user_id=current_user.id).all()
    now = datetime.now().date()
    
    assignments_data = []
    for resource in resources:
        for assignment in resource.assignments:
            stage = assignment.stage
            if stage and stage.project.user_id == current_user.id:
                dates = calculate_dates(stage.project, db)
                if dates.get(stage.id):
                    start_date = dates[stage.id][0]
                    end_date = dates[stage.id][1]
                    assignments_data.append({
                        'resource_id': resource.id,
                        'resource_name': resource.name,
                        'stage_name': stage.name,
                        'project_name': stage.project.name,
                        'start_date': start_date,
                        'end_date': end_date,
                        'percent_complete': stage.percent_complete,
                        'stage_id': stage.id,
                        'project_id': stage.project.id
                    })
    
    assignments_data.sort(key=lambda x: x['start_date'])
    
    return render_template('resource_calendar.html', 
                         resources=resources, 
                         assignments=assignments_data,
                         projects=projects,
                         now=now)

@app.route('/export_project/<int:project_id>')
@login_required
def export_project(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    stages = Stage.query.filter_by(project_id=project_id).all()
    dates = calculate_dates(project, db)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "План работ"
    
    headers = ['Название', 'Длительность', 'Зависимости', 'Ресурсы', 'План начало', 'План конец', 'Выполнено %', 'Факт завершение', 'План материалы', 'План работы', 'Факт материалы', 'Факт работы']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
    
    for row, stage in enumerate(stages, 2):
        deps = [dep.depends_on_stage.name for dep in stage.dependencies.all()]
        resources = [ass.resource.name for ass in stage.assignments]
        
        ws.cell(row=row, column=1, value=stage.name)
        ws.cell(row=row, column=2, value=stage.planned_duration)
        ws.cell(row=row, column=3, value=', '.join(deps))
        ws.cell(row=row, column=4, value=', '.join(resources))
        
        if dates.get(stage.id):
            ws.cell(row=row, column=5, value=dates[stage.id][0].strftime('%Y-%m-%d'))
            ws.cell(row=row, column=6, value=dates[stage.id][1].strftime('%Y-%m-%d'))
        
        ws.cell(row=row, column=7, value=stage.percent_complete)
        if stage.actual_end_date:
            ws.cell(row=row, column=8, value=stage.actual_end_date.strftime('%Y-%m-%d'))
        
        ws.cell(row=row, column=9, value=stage.planned_material_cost)
        ws.cell(row=row, column=10, value=stage.planned_labor_cost)
        ws.cell(row=row, column=11, value=stage.actual_material_cost)
        ws.cell(row=row, column=12, value=stage.actual_labor_cost)
    
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 30)
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name=f'project_{project.name}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/project/<int:project_id>/import_stages', methods=['GET', 'POST'])
@login_required
def import_stages(project_id):
    project = Project.query.get_or_404(project_id)
    if project.user_id != current_user.id:
        flash('Доступ запрещён', 'error')
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('Файл не выбран', 'error')
            return redirect(url_for('import_stages', project_id=project_id))
        
        file = request.files['file']
        if file.filename == '':
            flash('Файл не выбран', 'error')
            return redirect(url_for('import_stages', project_id=project_id))
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('Поддерживаются только файлы Excel (.xlsx, .xls)', 'error')
            return redirect(url_for('import_stages', project_id=project_id))
        
        try:
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            
            imported_count = 0
            skipped_count = 0
            errors = []
            
            existing_stages = {s.name: s.id for s in Stage.query.filter_by(project_id=project_id).all()}
            
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if not row[0]:
                    continue
                
                try:
                    name = str(row[0]).strip()
                    duration = int(row[1]) if row[1] else 1
                    depends_on = str(row[2]).strip() if row[2] else None
                    resources = str(row[3]).strip() if row[3] else None
                    percent_complete = int(row[6]) if len(row) > 6 and row[6] else 0
                    planned_material = float(row[7]) if len(row) > 7 and row[7] else 0
                    planned_labor = float(row[8]) if len(row) > 8 and row[8] else 0
                    
                    custom_start_date = None
                    custom_end_date = None
                    if len(row) > 4 and row[4]:
                        try:
                            if isinstance(row[4], datetime):
                                custom_start_date = row[4].date()
                            else:
                                custom_start_date = datetime.strptime(str(row[4]), '%d.%m.%Y').date()
                        except:
                            pass
                    if len(row) > 5 and row[5]:
                        try:
                            if isinstance(row[5], datetime):
                                custom_end_date = row[5].date()
                            else:
                                custom_end_date = datetime.strptime(str(row[5]), '%d.%m.%Y').date()
                        except:
                            pass
                    
                    if name in existing_stages:
                        skipped_count += 1
                        errors.append(f"Строка {row_idx}: Этап '{name}' уже существует")
                        continue
                    
                    stage = Stage(
                        name=name,
                        planned_duration=duration,
                        project_id=project_id,
                        percent_complete=percent_complete,
                        planned_material_cost=planned_material,
                        planned_labor_cost=planned_labor,
                        custom_start_date=custom_start_date,
                        custom_end_date=custom_end_date
                    )
                    db.session.add(stage)
                    db.session.flush()
                    
                    if depends_on and depends_on in existing_stages:
                        dep = Dependency(stage_id=stage.id, depends_on_stage_id=existing_stages[depends_on])
                        db.session.add(dep)
                    elif depends_on and depends_on not in existing_stages:
                        errors.append(f"Строка {row_idx}: Зависимость от '{depends_on}' не найдена")
                    
                    if resources:
                        resource_names = [r.strip() for r in resources.split(',')]
                        for r_name in resource_names:
                            resource = Resource.query.filter_by(name=r_name).first()
                            if not resource:
                                resource = Resource(name=r_name)
                                db.session.add(resource)
                                db.session.flush()
                            assign = Assignment(stage_id=stage.id, resource_id=resource.id)
                            db.session.add(assign)
                    
                    imported_count += 1
                    existing_stages[name] = stage.id
                    
                except Exception as e:
                    errors.append(f"Строка {row_idx}: {str(e)}")
                    skipped_count += 1
            
            db.session.commit()
            
            flash(f'Импорт завершён: добавлено {imported_count} этапов, пропущено {skipped_count}', 'success')
            if errors:
                for err in errors[:5]:
                    flash(err, 'error')
                    
        except Exception as e:
            flash(f'Ошибка при обработке файла: {str(e)}', 'error')
            return redirect(url_for('import_stages', project_id=project_id))
        
        return redirect(url_for('project_view', project_id=project_id))
    
    return render_template('import_stages.html', project=project)

@app.route('/resources')
@login_required
def resources_list():
    resources = Resource.query.all()
    return render_template('resources.html', resources=resources)

@app.route('/resources/add', methods=['POST'])
@login_required
def add_resource():
    name = request.form['name']
    if name:
        resource = Resource(name=name)
        db.session.add(resource)
        db.session.commit()
        flash('Ресурс добавлен', 'success')
    return redirect(url_for('resources_list'))

@app.route('/resources/delete/<int:resource_id>')
@login_required
def delete_resource(resource_id):
    resource = Resource.query.get_or_404(resource_id)
    db.session.delete(resource)
    db.session.commit()
    flash('Ресурс удалён', 'success')
    return redirect(url_for('resources_list'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)