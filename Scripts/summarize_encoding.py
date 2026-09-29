"""Summarize completed encoding runs and test the General-format assumption."""
import argparse,csv,json
from collections import defaultdict
from pathlib import Path
from run_encoding_study import load_released_module,dump_csv

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--encoding-dir',type=Path,required=True);p.add_argument('--source-archive',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();a.output.mkdir(exist_ok=True,parents=True)
 method=load_released_module(a.source_archive,a.encoding_dir/'downloaded_source')
 rows=list(csv.DictReader((a.encoding_dir/'encoding_by_sheet.csv').open(encoding='utf-8-sig')))
 summaries=[]
 for name in sorted({r['workbook'] for r in rows}):
  subset=[r for r in rows if r['workbook']==name];s=lambda k:sum(int(r[k]) for r in subset)
  unchanged=0
  from openpyxl.utils.cell import range_boundaries
  for file in (a.encoding_dir/name).glob('*_lossless_index.json'):
   # Exports are UTF-8, including German sheet labels and currency symbols.
   # Windows otherwise defaults to a legacy encoding such as cp1252.
   for dtype,text,fmt,regions in json.loads(file.read_text(encoding='utf-8-sig')):
    if dtype!='f':continue
    if method.get_type('None' if fmt=='General' else fmt,text)==text:
     for r in regions:
      c1,r1,c2,r2=range_boundaries(r);unchanged+=(c2-c1+1)*(r2-r1+1)
  summaries.append(dict(workbook=name,sheets=len(subset),populated_cells=s('populated_cells'),formula_cells=s('formula_cells'),
    aggregation_completed=sum(r['aggregation_status']=='completed' for r in subset),aggregation_recursion_failures=sum(r['aggregation_status']=='recursion_failure' for r in subset),
    default_type_mapping_formula_content_changed=s('get_type_changed_formula_cells'),
    general_as_none_formula_content_preserved=unchanged,
    baseline_bytes=s('sparse_baseline_bytes'),lossless_index_bytes=s('lossless_index_bytes'),
    lossless_index_compression_ratio=s('sparse_baseline_bytes')/s('lossless_index_bytes')))
 dump_csv(a.output/'encoding_by_workbook.csv',summaries)
 # Minimal geometry test of the unchanged authors' aggregation method.
 data=[['1','2'],['3','word']];nfs=[['None','None'],['None','None']]
 areas=method.aggregate_similar_areas(data,nfs)
 geometry={'synthetic_software_test_not_empirical_LLM_result':True,'input':data,'formats':nfs,'areas':areas,'finding':'The integer component bounding box includes the different-type bottom-right cell. Region bounding boxes are not exact component membership.'}
 (a.output/'aggregation_geometry_test.json').write_text(json.dumps(geometry,indent=2),encoding='utf-8')
 print(json.dumps(summaries,indent=2))
if __name__=='__main__':main()
