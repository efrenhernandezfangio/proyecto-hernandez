from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import json

db = SQLAlchemy()

class Role(db.Model):
    __tablename__ = 'roles'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    description = db.Column(db.String(200))
    permissions = db.Column(db.Text)  # JSON string de permisos
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relaciones
    users = db.relationship('User', back_populates='role')
    
    def __repr__(self):
        return f'<Role {self.name}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'description': self.description,
            'permissions': json.loads(self.permissions) if self.permissions else {},
            'created_at': self.created_at.isoformat() if self.created_at else None
        }

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    first_name = db.Column(db.String(50))
    last_name = db.Column(db.String(50))
    role_id = db.Column(db.Integer, db.ForeignKey('roles.id'), nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    is_verified = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relaciones
    role = db.relationship('Role', back_populates='users')
    projects = db.relationship('Project', back_populates='created_by')
    notifications = db.relationship('Notification', back_populates='user')
    
    def __repr__(self):
        return f'<User {self.username}>'
    
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
    
    def has_permission(self, permission):
        if not self.role or not self.role.permissions:
            return False
        permissions = json.loads(self.role.permissions)
        return permission in permissions
    
    def to_dict(self):
        return {
            'id': self.id,
            'username': self.username,
            'email': self.email,
            'first_name': self.first_name,
            'last_name': self.last_name,
            'role': self.role.to_dict() if self.role else None,
            'is_active': self.is_active,
            'is_verified': self.is_verified,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'last_login': self.last_login.isoformat() if self.last_login else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }

class Project(db.Model):
    __tablename__ = 'projects'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    user_id = db.Column(db.String(50), nullable=False)  # ID del usuario del proyecto
    project_type = db.Column(db.String(20), nullable=False)  # PTP, PTMP
    status = db.Column(db.String(20), default='draft')  # draft, in_progress, completed, archived
    description = db.Column(db.Text)
    created_by_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    assigned_to_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    
    # Relaciones
    created_by = db.relationship('User', foreign_keys=[created_by_id], back_populates='projects')
    assigned_to = db.relationship('User', foreign_keys=[assigned_to_id])
    images = db.relationship('ProjectImage', back_populates='project', cascade='all, delete-orphan')
    files = db.relationship('ProjectFile', back_populates='project', cascade='all, delete-orphan')
    
    def __repr__(self):
        return f'<Project {self.name}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'user_id': self.user_id,
            'project_type': self.project_type,
            'status': self.status,
            'description': self.description,
            'created_by': self.created_by.to_dict() if self.created_by else None,
            'assigned_to': self.assigned_to.to_dict() if self.assigned_to else None,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None,
            'image_count': len(self.images),
            'file_count': len(self.files)
        }

class ProjectImage(db.Model):
    __tablename__ = 'project_images'
    
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    image_type = db.Column(db.String(50), nullable=False)  # planos_a, planos_b, fotos_a, fotos_b
    image_number = db.Column(db.Integer, nullable=False)
    file_path = db.Column(db.String(255), nullable=False)
    file_name = db.Column(db.String(255))
    file_size = db.Column(db.Integer)  # en bytes
    uploaded_by_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relaciones
    project = db.relationship('Project', back_populates='images')
    uploaded_by = db.relationship('User')
    
    def __repr__(self):
        return f'<ProjectImage {self.image_type}_{self.image_number}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'image_type': self.image_type,
            'image_number': self.image_number,
            'file_path': self.file_path,
            'file_name': self.file_name,
            'file_size': self.file_size,
            'uploaded_by': self.uploaded_by.to_dict() if self.uploaded_by else None,
            'uploaded_at': self.uploaded_at.isoformat() if self.uploaded_at else None
        }

class ProjectFile(db.Model):
    __tablename__ = 'project_files'
    
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    file_type = db.Column(db.String(50), nullable=False)  # excel, pdf, doc, etc.
    file_path = db.Column(db.String(255), nullable=False)
    file_name = db.Column(db.String(255), nullable=False)
    file_size = db.Column(db.Integer)  # en bytes
    uploaded_by_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relaciones
    project = db.relationship('Project', back_populates='files')
    uploaded_by = db.relationship('User')
    
    def __repr__(self):
        return f'<ProjectFile {self.file_name}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'file_type': self.file_type,
            'file_path': self.file_path,
            'file_name': self.file_name,
            'file_size': self.file_size,
            'uploaded_by': self.uploaded_by.to_dict() if self.uploaded_by else None,
            'uploaded_at': self.uploaded_at.isoformat() if self.uploaded_at else None
        }

class Notification(db.Model):
    __tablename__ = 'notifications'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    title = db.Column(db.String(100), nullable=False)
    message = db.Column(db.Text, nullable=False)
    type = db.Column(db.String(20), default='info')  # info, warning, success, error
    is_read = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relaciones
    user = db.relationship('User', back_populates='notifications')
    
    def __repr__(self):
        return f'<Notification {self.title}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'user_id': self.user_id,
            'title': self.title,
            'message': self.message,
            'type': self.type,
            'is_read': self.is_read,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }

class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    action = db.Column(db.String(100), nullable=False)  # create, update, delete, login, etc.
    resource_type = db.Column(db.String(50))  # user, project, file, etc.
    resource_id = db.Column(db.Integer)
    details = db.Column(db.Text)  # JSON string con detalles adicionales
    ip_address = db.Column(db.String(45))
    user_agent = db.Column(db.String(500))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relaciones
    user = db.relationship('User')
    
    def __repr__(self):
        return f'<AuditLog {self.action} by {self.user_id}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'user_id': self.user_id,
            'action': self.action,
            'resource_type': self.resource_type,
            'resource_id': self.resource_id,
            'details': json.loads(self.details) if self.details else {},
            'ip_address': self.ip_address,
            'user_agent': self.user_agent,
            'created_at': self.created_at.isoformat() if self.created_at else None
        } 