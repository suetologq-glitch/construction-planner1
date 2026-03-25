import os
import logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler
from telegram.error import TimedOut, NetworkError

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Конфигурация - ЗАМЕНИТЕ НА ВАШ ТОКЕН
BOT_TOKEN = "8453224290:AAH6roU50OwWC7mURtHTIykLPL74272O5Lw"
UPLOAD_FOLDER = "static/uploads"

# Создаем папку для загрузок если её нет
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Состояния пользователей
user_projects = {}
user_waiting_photo = {}

# Импортируем модели позже, чтобы избежать циклических импортов
def get_db_models():
    from models import db, Project, Stage, Resource, Assignment, Photo
    from app import app
    return app, db, Project, Stage, Resource, Assignment, Photo

# Главное меню
async def main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE, user_id=None):
    keyboard = [
        [InlineKeyboardButton("📋 Мои проекты", callback_data="show_projects")]
    ]
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    if update.callback_query:
        await update.callback_query.message.edit_text(
            "🏗️ Главное меню\n\nВыберите действие:",
            reply_markup=reply_markup
        )
        await update.callback_query.answer()
    else:
        await update.message.reply_text(
            "🏗️ Главное меню\n\nВыберите действие:",
            reply_markup=reply_markup
        )

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await main_menu(update, context)

async def show_projects(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    try:
        app, db, Project, _, _, _, _ = get_db_models()
        with app.app_context():
            projects = Project.query.all()
            if projects:
                keyboard = []
                for p in projects:
                    keyboard.append([InlineKeyboardButton(f"📁 {p.name}", callback_data=f"select_project_{p.id}")])
                keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_menu")])
                
                reply_markup = InlineKeyboardMarkup(keyboard)
                await query.message.edit_text(
                    "📋 Выберите проект:",
                    reply_markup=reply_markup
                )
            else:
                keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_menu")]]
                reply_markup = InlineKeyboardMarkup(keyboard)
                await query.message.edit_text(
                    "❌ Нет активных проектов.\n\nСоздайте проект в веб-приложении.",
                    reply_markup=reply_markup
                )
    except Exception as e:
        logger.error(f"Ошибка при получении проектов: {e}")
        await query.message.edit_text("❌ Ошибка при получении списка проектов")

async def select_project(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    project_id = int(query.data.split("_")[2])
    user_id = query.from_user.id
    
    try:
        app, db, Project, _, _, _, _ = get_db_models()
        with app.app_context():
            project = Project.query.get(project_id)
            if project:
                user_projects[user_id] = project_id
                
                keyboard = [
                    [InlineKeyboardButton("📊 Этапы проекта", callback_data="show_stages")],
                    [InlineKeyboardButton("✅ Отметить выполнение", callback_data="complete_stage")],
                    [InlineKeyboardButton("📈 Указать прогресс", callback_data="set_progress")],
                    [InlineKeyboardButton("📸 Отправить фото", callback_data="send_photo")],
                    [InlineKeyboardButton("◀️ Главное меню", callback_data="back_to_menu")]
                ]
                reply_markup = InlineKeyboardMarkup(keyboard)
                
                await query.message.edit_text(
                    f"✅ Выбран проект: {project.name}\n\nЧто хотите сделать?",
                    reply_markup=reply_markup
                )
            else:
                await query.message.edit_text("❌ Проект не найден.")
    except Exception as e:
        logger.error(f"Ошибка при выборе проекта: {e}")
        await query.message.edit_text("❌ Ошибка при выборе проекта")

async def show_stages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if user_id not in user_projects:
        await query.message.edit_text("❌ Сначала выберите проект")
        return
    
    try:
        app, db, Project, Stage, _, _, _ = get_db_models()
        with app.app_context():
            project = Project.query.get(user_projects[user_id])
            stages = Stage.query.filter_by(project_id=project.id).all()
            
            if stages:
                msg = f"📊 <b>{project.name}</b>\n\n"
                for i, stage in enumerate(stages, 1):
                    if stage.percent_complete >= 100:
                        status = "✅"
                    elif stage.percent_complete > 0:
                        status = "🔄"
                    else:
                        status = "⏳"
                    msg += f"{i}. {status} <b>{stage.name}</b> - {stage.percent_complete}%\n"
                
                keyboard = [
                    [InlineKeyboardButton("◀️ Назад к проекту", callback_data="back_to_project")],
                    [InlineKeyboardButton("🏠 Главное меню", callback_data="back_to_menu")]
                ]
                reply_markup = InlineKeyboardMarkup(keyboard)
                
                await query.message.edit_text(
                    msg,
                    parse_mode="HTML",
                    reply_markup=reply_markup
                )
            else:
                keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_project")]]
                reply_markup = InlineKeyboardMarkup(keyboard)
                await query.message.edit_text(
                    "📭 Нет добавленных этапов",
                    reply_markup=reply_markup
                )
    except Exception as e:
        logger.error(f"Ошибка при получении этапов: {e}")
        await query.message.edit_text("❌ Ошибка при получении списка этапов")

async def complete_stage_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if user_id not in user_projects:
        await query.message.edit_text("❌ Сначала выберите проект")
        return
    
    try:
        app, db, Project, Stage, _, _, _ = get_db_models()
        with app.app_context():
            project = Project.query.get(user_projects[user_id])
            stages = Stage.query.filter_by(project_id=project.id).all()
            
            if stages:
                keyboard = []
                for i, stage in enumerate(stages, 1):
                    if stage.percent_complete < 100:
                        keyboard.append([InlineKeyboardButton(f"{i}. {stage.name}", callback_data=f"complete_{stage.id}")])
                
                if not keyboard:
                    keyboard.append([InlineKeyboardButton("✅ Все этапы уже завершены", callback_data="back_to_project")])
                keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_project")])
                reply_markup = InlineKeyboardMarkup(keyboard)
                
                await query.message.edit_text(
                    f"✅ Выберите этап для отметки выполнения:\n\nПроект: {project.name}",
                    reply_markup=reply_markup
                )
            else:
                keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_project")]]
                reply_markup = InlineKeyboardMarkup(keyboard)
                await query.message.edit_text(
                    "📭 Нет этапов для отметки",
                    reply_markup=reply_markup
                )
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await query.message.edit_text("❌ Ошибка")

async def complete_stage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    stage_id = int(query.data.split("_")[1])
    
    try:
        app, db, _, Stage, _, _, _ = get_db_models()
        with app.app_context():
            stage = Stage.query.get(stage_id)
            if stage:
                stage.percent_complete = 100
                stage.actual_end_date = datetime.now().date()
                db.session.commit()
                
                await query.message.edit_text(
                    f"✅ Этап '{stage.name}' отмечен как выполненный!\n\nПрогресс: 100%"
                )
                
                # Показываем обновлённый список этапов
                await show_stages(update, context)
            else:
                await query.message.edit_text("❌ Этап не найден")
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await query.message.edit_text("❌ Ошибка при отметке выполнения")

async def set_progress_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if user_id not in user_projects:
        await query.message.edit_text("❌ Сначала выберите проект")
        return
    
    try:
        app, db, Project, Stage, _, _, _ = get_db_models()
        with app.app_context():
            project = Project.query.get(user_projects[user_id])
            stages = Stage.query.filter_by(project_id=project.id).all()
            
            if stages:
                keyboard = []
                for i, stage in enumerate(stages, 1):
                    keyboard.append([InlineKeyboardButton(f"{i}. {stage.name} ({stage.percent_complete}%)", callback_data=f"progress_stage_{stage.id}")])
                
                keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_project")])
                reply_markup = InlineKeyboardMarkup(keyboard)
                
                await query.message.edit_text(
                    f"📈 Выберите этап для указания прогресса:\n\nПроект: {project.name}",
                    reply_markup=reply_markup
                )
            else:
                keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_project")]]
                reply_markup = InlineKeyboardMarkup(keyboard)
                await query.message.edit_text(
                    "📭 Нет этапов",
                    reply_markup=reply_markup
                )
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await query.message.edit_text("❌ Ошибка")

async def ask_progress(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    stage_id = int(query.data.split("_")[2])
    context.user_data['progress_stage_id'] = stage_id
    
    keyboard = []
    for percent in [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]:
        keyboard.append([InlineKeyboardButton(f"{percent}%", callback_data=f"set_percent_{percent}_{stage_id}")])
    keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="set_progress")])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    try:
        app, db, _, Stage, _, _, _ = get_db_models()
        with app.app_context():
            stage = Stage.query.get(stage_id)
            await query.message.edit_text(
                f"📈 Выберите процент выполнения для этапа:\n\n<b>{stage.name}</b>\nТекущий прогресс: {stage.percent_complete}%",
                parse_mode="HTML",
                reply_markup=reply_markup
            )
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await query.message.edit_text("❌ Ошибка")

async def set_progress(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    parts = query.data.split("_")
    percent = int(parts[2])
    stage_id = int(parts[3])
    
    try:
        app, db, _, Stage, _, _, _ = get_db_models()
        with app.app_context():
            stage = Stage.query.get(stage_id)
            if stage:
                stage.percent_complete = percent
                if percent >= 100:
                    stage.actual_end_date = datetime.now().date()
                else:
                    stage.actual_end_date = None
                db.session.commit()
                
                await query.message.edit_text(
                    f"✅ Для этапа '{stage.name}' установлен прогресс: {percent}%"
                )
                
                # Показываем обновлённый список этапов
                await show_stages(update, context)
            else:
                await query.message.edit_text("❌ Этап не найден")
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await query.message.edit_text("❌ Ошибка")

async def send_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if user_id not in user_projects:
        await query.message.edit_text("❌ Сначала выберите проект")
        return
    
    user_waiting_photo[user_id] = True
    
    keyboard = [[InlineKeyboardButton("◀️ Отмена", callback_data="back_to_project")]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await query.message.edit_text(
        "📸 Отправьте фото с подписью в формате:\n\n"
        "<b>Название этапа: описание работы</b>\n\n"
        "Пример: Фундамент: Заливка завершена 20 марта\n\n"
        "Или отправьте фото без подписи — оно будет сохранено в общую папку.",
        parse_mode="HTML",
        reply_markup=reply_markup
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in user_projects:
        await update.message.reply_text("❌ Сначала выберите проект через меню")
        return
    
    try:
        caption = update.message.caption or ""
        
        if ":" in caption:
            stage_name, description = caption.split(":", 1)
            stage_name = stage_name.strip()
            description = description.strip()
        else:
            stage_name = "Общее"
            description = caption
        
        photo_file = await update.message.photo[-1].get_file()
        filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{user_id}.jpg"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        await photo_file.download_to_drive(filepath)
        
        app, db, _, _, _, _, Photo = get_db_models()
        with app.app_context():
            photo = Photo(
                project_id=user_projects[user_id],
                stage_name=stage_name,
                caption=description,
                filename=filename,
                filepath=f"uploads/{filename}",
                telegram_user=str(update.effective_user.first_name)
            )
            db.session.add(photo)
            db.session.commit()
        
        await update.message.reply_text(
            f"✅ Фото сохранено!\n"
            f"🏗️ Этап: {stage_name}\n"
            f"📝 Подпись: {description}"
        )
        
        # Показываем меню проекта
        user_id = update.effective_user.id
        if user_id in user_projects:
            app, db, Project, _, _, _, _ = get_db_models()
            with app.app_context():
                project = Project.query.get(user_projects[user_id])
                if project:
                    keyboard = [
                        [InlineKeyboardButton("📊 Этапы проекта", callback_data="show_stages")],
                        [InlineKeyboardButton("✅ Отметить выполнение", callback_data="complete_stage")],
                        [InlineKeyboardButton("📈 Указать прогресс", callback_data="set_progress")],
                        [InlineKeyboardButton("📸 Отправить фото", callback_data="send_photo")],
                        [InlineKeyboardButton("◀️ Главное меню", callback_data="back_to_menu")]
                    ]
                    reply_markup = InlineKeyboardMarkup(keyboard)
                    await update.message.reply_text(
                        f"✅ Проект: {project.name}\n\nЧто хотите сделать?",
                        reply_markup=reply_markup
                    )
        
    except Exception as e:
        logger.error(f"Ошибка при сохранении фото: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)[:100]}")

async def back_to_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await main_menu(update, context, query.from_user.id)

async def back_to_project(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if user_id in user_projects:
        try:
            app, db, Project, _, _, _, _ = get_db_models()
            with app.app_context():
                project = Project.query.get(user_projects[user_id])
                if project:
                    keyboard = [
                        [InlineKeyboardButton("📊 Этапы проекта", callback_data="show_stages")],
                        [InlineKeyboardButton("✅ Отметить выполнение", callback_data="complete_stage")],
                        [InlineKeyboardButton("📈 Указать прогресс", callback_data="set_progress")],
                        [InlineKeyboardButton("📸 Отправить фото", callback_data="send_photo")],
                        [InlineKeyboardButton("◀️ Главное меню", callback_data="back_to_menu")]
                    ]
                    reply_markup = InlineKeyboardMarkup(keyboard)
                    
                    await query.message.edit_text(
                        f"✅ Проект: {project.name}\n\nЧто хотите сделать?",
                        reply_markup=reply_markup
                    )
                else:
                    await main_menu(update, context, user_id)
        except Exception as e:
            await main_menu(update, context, user_id)
    else:
        await main_menu(update, context, user_id)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await main_menu(update, context)

def run_bot():
    """Запуск бота"""
    try:
        application = (
            Application.builder()
            .token(BOT_TOKEN)
            .connect_timeout(30.0)
            .read_timeout(30.0)
            .write_timeout(30.0)
            .build()
        )
        
        # Команды
        application.add_handler(CommandHandler("start", start))
        application.add_handler(CommandHandler("help", help_command))
        
        # Callback обработчики (кнопки)
        application.add_handler(CallbackQueryHandler(show_projects, pattern="^show_projects$"))
        application.add_handler(CallbackQueryHandler(select_project, pattern="^select_project_"))
        application.add_handler(CallbackQueryHandler(show_stages, pattern="^show_stages$"))
        application.add_handler(CallbackQueryHandler(complete_stage_menu, pattern="^complete_stage$"))
        application.add_handler(CallbackQueryHandler(complete_stage, pattern="^complete_\\d+$"))
        application.add_handler(CallbackQueryHandler(set_progress_menu, pattern="^set_progress$"))
        application.add_handler(CallbackQueryHandler(ask_progress, pattern="^progress_stage_"))
        application.add_handler(CallbackQueryHandler(set_progress, pattern="^set_percent_"))
        application.add_handler(CallbackQueryHandler(send_photo, pattern="^send_photo$"))
        application.add_handler(CallbackQueryHandler(back_to_menu, pattern="^back_to_menu$"))
        application.add_handler(CallbackQueryHandler(back_to_project, pattern="^back_to_project$"))
        
        # Обработчик фото
        application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
        
        logger.info("Бот запускается...")
        application.run_polling(allowed_updates=Update.ALL_TYPES)
        
    except TimedOut:
        logger.error("Таймаут подключения. Проверьте интернет-соединение.")
        print("\n" + "="*50)
        print("❌ ОШИБКА: Не удалось подключиться к Telegram")
        print("Проверьте интернет-соединение и токен бота")
        print("="*50)
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        print(f"\n❌ Ошибка: {e}")

if __name__ == "__main__":
    run_bot()