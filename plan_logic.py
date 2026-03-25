from datetime import datetime, timedelta
from models import Stage, Dependency

def calculate_dates(project, db):
    """
    Рассчитывает для всех этапов проекта даты начала и окончания.
    Если у этапа есть custom_start_date и custom_end_date, использует их.
    Иначе рассчитывает на основе зависимостей.
    """
    stages = Stage.query.filter_by(project_id=project.id).all()
    if not stages:
        return {}

    # Проверяем, есть ли этапы с пользовательскими датами
    custom_dates_exist = any(s.custom_start_date and s.custom_end_date for s in stages)
    
    if custom_dates_exist:
        # Используем даты из Excel
        result = {}
        for stage in stages:
            if stage.custom_start_date and stage.custom_end_date:
                result[stage.id] = (stage.custom_start_date, stage.custom_end_date)
            else:
                # Если даты не заданы, рассчитываем
                result[stage.id] = (project.start_date, project.start_date + timedelta(days=stage.planned_duration - 1))
        return result
    
    # Иначе стандартный расчёт на основе зависимостей
    deps = {}
    for stage in stages:
        deps[stage.id] = []
    dependencies = Dependency.query.all()
    for dep in dependencies:
        stage = Stage.query.get(dep.stage_id)
        if stage and stage.project_id == project.id:
            deps[dep.stage_id].append(dep.depends_on_stage_id)

    from collections import deque
    in_degree = {stage.id: 0 for stage in stages}
    for stage_id, preds in deps.items():
        for p in preds:
            in_degree[stage_id] += 1

    queue = deque([sid for sid, deg in in_degree.items() if deg == 0])
    order = []
    while queue:
        sid = queue.popleft()
        order.append(sid)
        for other, preds in deps.items():
            if sid in preds:
                in_degree[other] -= 1
                if in_degree[other] == 0:
                    queue.append(other)

    start = {sid: project.start_date for sid in order}
    end = {}
    for sid in order:
        stage = next((s for s in stages if s.id == sid), None)
        if not stage:
            continue
        if deps[sid]:
            max_end = max(end[p] for p in deps[sid] if p in end)
            start[sid] = max(max_end, project.start_date)
        else:
            start[sid] = max(start[sid], project.start_date)
        end[sid] = start[sid] + timedelta(days=stage.planned_duration - 1)

    return {sid: (start[sid], end[sid]) for sid in order}


def find_critical_path(project, db):
    dates = calculate_dates(project, db)
    if not dates:
        return []
    
    max_end = max(end for _, end in dates.values())
    critical = set()
    
    def find_predecessors(stage_id):
        if stage_id in critical:
            return
        critical.add(stage_id)
        deps = Dependency.query.filter_by(stage_id=stage_id).all()
        for dep in deps:
            pred_end = dates[dep.depends_on_stage_id][1]
            stage = Stage.query.get(stage_id)
            if pred_end == dates[stage_id][1] - timedelta(days=stage.planned_duration - 1):
                find_predecessors(dep.depends_on_stage_id)
    
    for stage_id, (_, end) in dates.items():
        if end == max_end:
            find_predecessors(stage_id)
    
    return list(critical)