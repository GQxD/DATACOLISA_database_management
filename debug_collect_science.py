from pathlib import Path
import sys

root = Path(__file__).resolve().parent
sys.path.insert(0, str(root / 'code'))

from generer_collec_science import generer_collec_science

path = Path(r'c:/Users/anaubin/Desktop/ECAILLE BASE ORDINATEUR/colisa en cours/COLISA en cours.xlsx')
output = root / 'collect_science_debug_output.xlsx'
result = generer_collec_science(
    colisa_path=path,
    output_path=output,
    collection_id=1,
    sample_status_id=1,
    referent_id=1,
    sample_multiple_value=5,
    containers={},
    forcer_anomalies=False,
    colisa_sheet='Feuil1',
    allowed_num_individus=None,
    prefer_fixed_num_individu_column=True,
    md_num_individu_column_index=None,
)

print('result', result)
print('output exists', output.exists())
