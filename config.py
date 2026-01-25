import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

TECH_MAP = {
    "2G_Down":"2G",
    "3G_Down":"3G",
    "4G_Down":"4G",
    "5G_Down":"5G"
}
