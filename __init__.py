from flask import Flask
from flask_wtf.csrf import CSRFProtect
import os

app = Flask(__name__)
# In production, this should be set as an environment variable
# For development, we generate a random secret key - this should be kept secret and not shared in version control
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY') or os.urandom(24).hex()

csrf = CSRFProtect(app)