from pathlib import Path
import os
import shutil
root = Path(__file__).parents[1]
out = root / "solution" / "Golden" / 'inventory_eom.xlsx'
output_dir = Path(os.environ.get("OUTPUT_DIR", "/app/output"))
target = output_dir / 'inventory_eom.xlsx'
target.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(out, target)
