from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(50), default='user')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    projects = db.relationship('Project', backref='owner', lazy=True, cascade='all, delete-orphan')
    templates = db.relationship('ProjectTemplate', backref='owner', lazy=True, cascade='all, delete-orphan')
    notifications = db.relationship('Notification', backref='user', lazy=True, cascade='all, delete-orphan')
    team_memberships = db.relationship('ProjectTeam', backref='user', lazy=True, cascade='all, delete-orphan')
    
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
    
    def is_admin(self):
        return self.role == 'admin'
    
    def is_manager(self):
        return self.role in ('admin', 'manager')
    
    def is_foreman(self):
        return self.role in ('admin', 'manager', 'foreman')
    
    def is_supplier(self):
        return self.role in ('admin', 'manager', 'supplier')


class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    latitude = db.Column(db.Float, default=55.751244)
    longitude = db.Column(db.Float, default=37.618423)
    address = db.Column(db.String(300), nullable=True)
    
    stages = db.relationship('Stage', backref='project', lazy=True, cascade='all, delete-orphan')
    photos = db.relationship('Photo', backref='project', lazy=True, cascade='all, delete-orphan')
    purchase_orders = db.relationship('PurchaseOrder', backref='project', lazy=True, cascade='all, delete-orphan')
    acts = db.relationship('Act', backref='project', lazy=True, cascade='all, delete-orphan')
    contracts = db.relationship('Contract', backref='project', lazy=True, cascade='all, delete-orphan')
    team = db.relationship('ProjectTeam', backref='project', lazy=True, cascade='all, delete-orphan')


class ProjectTeam(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    role = db.Column(db.String(50), default='viewer')
    __table_args__ = (db.UniqueConstraint('project_id', 'user_id', name='unique_project_user'),)


class Stage(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    planned_duration = db.Column(db.Integer, nullable=False, default=1)
    actual_end_date = db.Column(db.Date, nullable=True)
    percent_complete = db.Column(db.Integer, default=0)
    
    custom_start_date = db.Column(db.Date, nullable=True)
    custom_end_date = db.Column(db.Date, nullable=True)
    
    planned_material_cost = db.Column(db.Float, default=0.0)
    planned_labor_cost = db.Column(db.Float, default=0.0)
    actual_material_cost = db.Column(db.Float, default=0.0)
    actual_labor_cost = db.Column(db.Float, default=0.0)
    
    dependencies = db.relationship('Dependency',
                                   foreign_keys='Dependency.stage_id',
                                   backref='stage',
                                   lazy='dynamic',
                                   cascade='all, delete-orphan')
    
    dependents = db.relationship('Dependency',
                                 foreign_keys='Dependency.depends_on_stage_id',
                                 backref='depends_on_stage',
                                 lazy='dynamic',
                                 cascade='all, delete-orphan')
    
    assignments = db.relationship('Assignment', backref='stage', lazy=True, cascade='all, delete-orphan')
    purchase_items = db.relationship('PurchaseItem', backref='stage', lazy=True, cascade='all, delete-orphan')


class Dependency(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    stage_id = db.Column(db.Integer, db.ForeignKey('stage.id'), nullable=False)
    depends_on_stage_id = db.Column(db.Integer, db.ForeignKey('stage.id'), nullable=False)
    __table_args__ = (db.UniqueConstraint('stage_id', 'depends_on_stage_id', name='unique_dep'),)


class Resource(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    assignments = db.relationship('Assignment', backref='resource', lazy=True, cascade='all, delete-orphan')


class Assignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    stage_id = db.Column(db.Integer, db.ForeignKey('stage.id'), nullable=False)
    resource_id = db.Column(db.Integer, db.ForeignKey('resource.id'), nullable=False)
    __table_args__ = (db.UniqueConstraint('stage_id', 'resource_id', name='unique_assign'),)


class Photo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    stage_name = db.Column(db.String(100), nullable=False)
    caption = db.Column(db.String(500), nullable=True)
    filename = db.Column(db.String(200), nullable=False)
    filepath = db.Column(db.String(500), nullable=False)
    telegram_user = db.Column(db.String(100), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class PurchaseOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    order_number = db.Column(db.String(50), nullable=False, unique=True)
    supplier = db.Column(db.String(200), nullable=True)
    order_date = db.Column(db.Date, nullable=False)
    delivery_date = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(50), default='pending')
    total_amount = db.Column(db.Float, default=0.0)
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship('PurchaseItem', backref='order', lazy=True, cascade='all, delete-orphan')


class PurchaseItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('purchase_order.id'), nullable=False)
    stage_id = db.Column(db.Integer, db.ForeignKey('stage.id'), nullable=False)
    material_name = db.Column(db.String(200), nullable=False)
    quantity = db.Column(db.Float, default=0.0)
    unit = db.Column(db.String(20), default='шт')
    unit_price = db.Column(db.Float, default=0.0)
    total_price = db.Column(db.Float, default=0.0)
    status = db.Column(db.String(50), default='pending')


class Act(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    act_number = db.Column(db.String(50), nullable=False, unique=True)
    act_date = db.Column(db.Date, nullable=False)
    period_start = db.Column(db.Date, nullable=False)
    period_end = db.Column(db.Date, nullable=False)
    total_amount = db.Column(db.Float, default=0.0)
    status = db.Column(db.String(50), default='draft')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship('ActItem', backref='act', lazy=True, cascade='all, delete-orphan')


class ActItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    act_id = db.Column(db.Integer, db.ForeignKey('act.id'), nullable=False)
    stage_id = db.Column(db.Integer, db.ForeignKey('stage.id'), nullable=False)
    work_name = db.Column(db.String(200), nullable=False)
    quantity = db.Column(db.Float, default=0.0)
    unit = db.Column(db.String(20), default='шт')
    unit_price = db.Column(db.Float, default=0.0)
    total_price = db.Column(db.Float, default=0.0)


class Contract(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    contract_number = db.Column(db.String(50), nullable=False, unique=True)
    contract_date = db.Column(db.Date, nullable=False)
    contractor_name = db.Column(db.String(200), nullable=False)
    contractor_inn = db.Column(db.String(20), nullable=True)
    contractor_kpp = db.Column(db.String(20), nullable=True)
    contractor_ogrn = db.Column(db.String(20), nullable=True)
    contractor_address = db.Column(db.String(300), nullable=True)
    contractor_fact_address = db.Column(db.String(300), nullable=True)
    contractor_phone = db.Column(db.String(20), nullable=True)
    contractor_email = db.Column(db.String(100), nullable=True)
    contractor_director = db.Column(db.String(200), nullable=True)
    contractor_basis = db.Column(db.String(200), nullable=True)
    contractor_bank_name = db.Column(db.String(200), nullable=True)
    contractor_bik = db.Column(db.String(20), nullable=True)
    contractor_account = db.Column(db.String(30), nullable=True)
    contractor_correspondent_account = db.Column(db.String(30), nullable=True)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=True)
    total_amount = db.Column(db.Float, default=0.0)
    object_name = db.Column(db.String(200), nullable=True)
    object_address = db.Column(db.String(300), nullable=True)
    subject = db.Column(db.Text, nullable=True)
    documentation = db.Column(db.Text, nullable=True)
    doc_deadline = db.Column(db.String(200), nullable=True)
    doc_method = db.Column(db.String(100), nullable=True)
    special_conditions = db.Column(db.Text, nullable=True)
    liability = db.Column(db.Text, nullable=True)
    warranty = db.Column(db.String(200), nullable=True)
    status = db.Column(db.String(50), default='draft')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    work_items = db.relationship('ContractWorkItem', backref='contract', lazy=True, cascade='all, delete-orphan')


class ContractWorkItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    contract_id = db.Column(db.Integer, db.ForeignKey('contract.id'), nullable=False)
    work_name = db.Column(db.String(200), nullable=False)
    unit = db.Column(db.String(20), default='шт')
    quantity = db.Column(db.Float, default=0.0)
    unit_price = db.Column(db.Float, default=0.0)
    total_price = db.Column(db.Float, default=0.0)


class ProjectTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text, nullable=True)
    stages_data = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    message = db.Column(db.Text, nullable=False)
    link = db.Column(db.String(500), nullable=True)
    is_read = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)