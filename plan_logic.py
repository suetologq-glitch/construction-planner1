from datetime import datetime, timedelta, date
from models import Stage, Dependency

HOLIDAYS = {
    date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 7),
    date(2026, 2, 23), date(2026, 3, 8), date(2026, 5, 1),
    date(2026, 5, 9), date(2026, 6, 12), date(2026, 11, 4),
}

def is_workday(d):
    if d.weekday() >= 5: return False
    if d in HOLIDAYS: return False
    return True

def next_workday(d):
    cur = d
    while not is_workday(cur):
        cur += timedelta(days=1)
    return cur

def add_workdays(start_date, days):
    if days <= 0: return start_date
    current = start_date
    count = 0
    while count < days:
        current += timedelta(days=1)
        if is_workday(current): count += 1
    return current

def calculate_dates(project, db):
    stages = Stage.query.filter_by(project_id=project.id).order_by(Stage.id).all()
    if not stages: return {}
    deps = {s.id: [] for s in stages}
    for dep in Dependency.query.all():
        stage = Stage.query.get(dep.stage_id)
        if stage and stage.project_id == project.id:
            deps[dep.stage_id].append(dep.depends_on_stage_id)
    from collections import deque
    in_degree = {s.id: 0 for s in stages}
    for sid, preds in deps.items():
        for p in preds: in_degree[sid] += 1
    queue = deque([sid for sid, deg in in_degree.items() if deg == 0])
    order = []
    while queue:
        sid = queue.popleft()
        order.append(sid)
        for other, preds in deps.items():
            if sid in preds:
                in_degree[other] -= 1
                if in_degree[other] == 0: queue.append(other)
    start = {}
    end = {}
    for sid in order:
        stage = next(s for s in stages if s.id == sid)
        if stage.percent_complete == 100 and stage.actual_end_date:
            end[sid] = stage.actual_end_date
            start[sid] = stage.actual_end_date - timedelta(days=stage.planned_duration - 1)
            continue
        if stage.custom_start_date and stage.custom_end_date:
            start[sid], end[sid] = stage.custom_start_date, stage.custom_end_date
            continue
        max_end = project.start_date - timedelta(days=1)
        for pred_id in deps[sid]:
            if pred_id in end:
                if end[pred_id] > max_end: max_end = end[pred_id]
        start[sid] = next_workday(max_end + timedelta(days=1))
        end[sid] = add_workdays(start[sid], stage.planned_duration - 1)
    return {sid: (start[sid], end[sid]) for sid in order}

def find_critical_path(project, db):
    return []

def predict_completion_date(project, db):
    stages = Stage.query.filter_by(project_id=project.id).all()
    if not stages: return None, None
    dates = {}
    for s in stages:
        if s.custom_start_date and s.custom_end_date:
            dates[s.id] = (s.custom_start_date, s.custom_end_date)
        else:
            dates.update(calculate_dates(project, db))
    if not dates: return None, None
    today = date.today()
    total_weight = sum(s.planned_duration for s in stages)
    if total_weight == 0: return None, None
    planned_progress = 0
    actual_progress = 0
    for stage in stages:
        if stage.id not in dates: continue
        start_d, end_d = dates[stage.id]
        weight = stage.planned_duration / total_weight
        if today < start_d:
            stage_planned = 0
        elif today > end_d:
            stage_planned = 100
        else:
            elapsed = (today - start_d).days + 1
            total_days = (end_d - start_d).days + 1
            stage_planned = min(100, (elapsed / total_days) * 100)
        planned_progress += stage_planned * weight
        actual_progress += stage.percent_complete * weight
    if planned_progress == 0: return None, None
    speed_factor = actual_progress / planned_progress if planned_progress > 0 else 1.0
    project_end_planned = max(end for _, end in dates.values())
    remaining_days = (project_end_planned - today).days
    if remaining_days <= 0: return project_end_planned, 0
    predicted_remaining = remaining_days / speed_factor if speed_factor > 0 else remaining_days
    predicted_end = today + timedelta(days=int(predicted_remaining))
    deviation = (project_end_planned - predicted_end).days
    return predicted_end, deviation