import sys
sys.path.append('/home/garciagi/GS1')
sys.path.append('/home/garciagi/.local/lib/python3.6/site-packages')

from main import app as application

if __name__ == '__main__':
    application.run(timeout=600)