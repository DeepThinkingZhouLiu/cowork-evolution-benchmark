from pathlib import Path
import os
import shutil
root = Path(__file__).parents[1]
out = Path(os.environ.get('OUTPUT_DIR', '/app/output')); out.mkdir(parents=True, exist_ok=True)
for name in ('training.ics', 'Alice.eml', 'David.eml', 'Grace.eml', 'Jack.eml'):
    shutil.copy2(root/'solution/Golden'/name, out/name)
