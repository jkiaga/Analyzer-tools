import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'your_secret_key'
    DEBUG = False
    TESTING = False

class ProductionConfig(Config):
    DATABASE_URI = 'mysql://user@localhost/foo'

class DevelopmentConfig(Config):
    DEBUG = True
    DATABASE_URI = 'sqlite:///:memory:'

class TestingConfig(Config):
    TESTING = True
    DATABASE_URI = 'sqlite:///:memory:'
