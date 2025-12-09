#!/bin/sh
# start.sh - development helper
export FLASK_APP=app.py
export FLASK_ENV=development
export PORT=${PORT:-5000}
flask run --host=0.0.0.0 --port=$PORT
