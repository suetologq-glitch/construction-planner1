from app import app, db
from models import Stage

with app.app_context():
    stages = Stage.query.all()
    for s in stages:
        s.custom_start_date = None
        s.custom_end_date = None
    db.session.commit()
    print(f"✅ Сброшены пользовательские даты у {len(stages)} этапов.")