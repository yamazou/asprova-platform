"""
Shared paths for asprova-platform (absolute, cwd-independent).

Used by apps that may be started with cwd set to apps/viewer or apps/bridge.
"""
import os

BASE_DIR = os.path.dirname(os.path.dirname(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(BASE_DIR, "data", "schedule.db")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "data", "uploads")
