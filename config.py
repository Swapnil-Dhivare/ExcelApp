import os

# Database configuration for Railway
DB_USER = os.environ.get('DB_USER', 'root')
DB_PASSWORD = os.environ.get('DB_PASSWORD', 'aYumXAVkRkDDPsahqrlpmWHTQFYfFOrC')
DB_HOST = os.environ.get('DB_HOST', 'tramway.proxy.rlwy.net')
DB_PORT = os.environ.get('DB_PORT', '58062')
DB_NAME = os.environ.get('DB_NAME', 'railway')

# SQLAlchemy database URI - updated for Railway
SQLALCHEMY_DATABASE_URI = f'mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
SQLALCHEMY_TRACK_MODIFICATIONS = False

# Application secret key
SECRET_KEY = os.environ.get('SECRET_KEY', 'your_secret_key_for_development')

# Debug mode
DEBUG = os.environ.get('FLASK_ENV') != 'production'
