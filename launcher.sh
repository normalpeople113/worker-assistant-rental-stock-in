#!/bin/bash

# Aktifin virtual environment (kalau pakai)
source /www/wwwroot/Worker_Rental_Stock_In/worker_env/bin/activate

# Masuk ke folder project
cd /www/wwwroot/Worker_Rental_Stock_In

# Jalanin script dan log output
python3 worker.py >> logs/worker.log 2>&1
