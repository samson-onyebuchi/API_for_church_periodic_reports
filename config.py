import os
from dotenv import load_dotenv
# basedir = os.path.abspath(os.path.dirname(__file__))

# # creating a configuration class
# class Config(object):
#     SECRET_KEY = os.environ.get('SECRET_KEY') 

# BASE_URI = os.environ.get('BASE_URI')  
# SAFEPAY_URI = os.environ.get('SAFEPAY_URI') 
# API_KEY = os.environ.get('API_KEY') 
# #MONGO_URI = os.environ.get('MONGO_URI') 
# USSD_username = os.environ.get('USSD_username') 
# USSD_api_key = os.environ.get('USSD_api_key') 

load_dotenv()
class Config(object):
    MONGO_URI = os.environ.get('MONGO_URI')
    BASE_URI = "http://127.0.0.1:5000"