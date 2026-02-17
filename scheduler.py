import subprocess
import sys
import time
from datetime import datetime

INTERVAL_MINUTES = 10

while True:
    print(f"\n=== Execution : {datetime.now().strftime('%d/%m/%Y %H:%M')} ===")
    subprocess.run([sys.executable, "download_reports.py"])
    subprocess.run([sys.executable, "send_reports.py"])
    print(f"Prochaine execution dans {INTERVAL_MINUTES} minutes...")
    time.sleep(INTERVAL_MINUTES * 60)
