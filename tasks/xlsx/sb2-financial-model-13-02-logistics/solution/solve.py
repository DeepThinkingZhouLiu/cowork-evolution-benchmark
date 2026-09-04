from pathlib import Path
import os
import shutil
root = Path(__file__).parents[1]
out = root / "solution" / "Golden" / 'logistics_model.xlsx'
output_dir = Path(os.environ.get("OUTPUT_DIR", "/app/output"))
target = output_dir / 'logistics_model.xlsx'
target.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(out, target)
