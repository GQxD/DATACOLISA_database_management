import sys
from pathlib import Path
# Ensure project root is on sys.path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from infrastructure.internal_target_workbook import build_numero_identification_formula

print(build_numero_identification_formula(4))
