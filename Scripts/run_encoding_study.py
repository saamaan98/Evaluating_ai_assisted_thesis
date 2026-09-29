"""Evaluate released SheetEncoder aggregation and a separate lossless index.

This runs the authors' get_type and aggregate_similar_areas method bodies,
without changing them. It does not run structural anchors or any LLM.
The supplementary package is downloaded and hash-checked; no private data is sent.
"""
import argparse, ast, csv, fnmatch, hashlib, importlib.util, json, sys, time, urllib.request, zipfile
from collections import defaultdict
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter

URL='https://aclanthology.org/attachments/2024.emnlp-main.1154.software.zip'
# Filled from the verified publisher attachment when the study was prepared.
EXPECTED='4cdf1efd0839c0720daa7566b9d7123f7eb66e434e916458986aa02a7dd8bc86'

def load_released_module(archive, destination):
 destination.mkdir(parents=True,exist_ok=True)
 if hashlib.sha256(archive.read_bytes()).hexdigest()!=EXPECTED:raise ValueError('Unexpected publisher source archive hash')
 with zipfile.ZipFile(archive) as z:
  for name in ('config.py','CaseJudge.py','TableDataAggregation.py'):
   (destination/name).write_bytes(z.read('code/compress_aggregation/'+name))
 for name in ('config','CaseJudge'):
  spec=importlib.util.spec_from_file_location(name,destination/f'{name}.py')
  module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);sys.modules[name]=module
 # Skip imports for unrelated file/token-count helpers. The actual method bodies
 # and supplied default configuration are unchanged. No tokenizer is simulated.
 tree=ast.parse((destination/'TableDataAggregation.py').read_text(encoding='utf-8-sig'))
 tree.body=[node for node in tree.body if not isinstance(node,ast.ImportFrom) or node.module!='utils']
 module=type(sys)('released_aggregation')
 exec(compile(tree,'released_TableDataAggregation.py','exec'),module.__dict__)
 return module.TableDataAggregation('','')

def dump_csv(path,rows):
 if not rows:
  raise ValueError('No worksheet results to save. Check the input folder and filename pattern. No CSV was written.')
 with path.open('w',newline='',encoding='utf-8-sig') as f:
  writer=csv.DictWriter(f,fieldnames=list(rows[0]));writer.writeheader();writer.writerows(rows)

def spans(points):
 rows=defaultdict(list)
 for r,c in points:rows[r].append(c)
 result=[]
 for r,columns in sorted(rows.items()):
  columns.sort();start=last=columns[0]
  for c in columns[1:]+[None]:
   if c is not None and c==last+1:last=c;continue
   a=f'{get_column_letter(start+1)}{r+1}';b=f'{get_column_letter(last+1)}{r+1}'
   result.append(a if start==last else a+':'+b)
   start=last=c
 return result

def main():
 parser=argparse.ArgumentParser(description=__doc__)
 parser.add_argument('--inputs',type=Path,required=True)
 parser.add_argument('--output',type=Path,required=True)
 parser.add_argument('--source-archive',type=Path)
 parser.add_argument('--pattern',default='W*_seeded.xlsx',
                     help='Input filename pattern (default: W*_seeded.xlsx). Searches this folder only.')
 args=parser.parse_args()
 folder=args.inputs.expanduser().resolve()
 if not folder.is_dir():
  parser.error(f'Input folder does not exist or is not a folder: {folder}\n'
               'Create it and place the seeded Excel files inside, or correct --inputs.')
 paths=sorted((p for p in folder.iterdir() if p.is_file()
               and not p.name.startswith('~$')
               and p.suffix.lower() in ('.xlsx','.xlsm')
               and fnmatch.fnmatchcase(p.name.lower(),args.pattern.lower())),
              key=lambda p:p.name.lower())
 if not paths:
  examples=', '.join(sorted(p.name for p in folder.iterdir() if p.is_file())[:10]) or '(no files)'
  parser.error(f'No Excel files matched {args.pattern!r} in: {folder}\n'
               f'Files found there: {examples}\n'
               'Use names such as W1_Real_seeded.xlsx. Files inside subfolders or a ZIP are not searched.\n'
               'If your names differ, use --pattern with a pattern matching only the intended seeded workbooks.')
 print(f'Input folder: {folder}',flush=True)
 print(f'Found {len(paths)} workbook(s):',flush=True)
 for path in paths: print(f'  {path.name}',flush=True)
 out=args.output.expanduser().resolve();out.mkdir(parents=True,exist_ok=True)
 archive=args.source_archive or out/'publisher_software.zip'
 if not archive.exists():
  archive.write_bytes(urllib.request.urlopen(URL,timeout=60).read())
 method=load_released_module(archive,out/'downloaded_source')
 records=[]
 for path in paths:
  wb=load_workbook(path,read_only=True,data_only=False,keep_links=False)
  wbhash=hashlib.sha256(path.read_bytes()).hexdigest()
  dest=out/path.stem;dest.mkdir(exist_ok=True)
  for number,ws in enumerate(wb,1):
   start=time.perf_counter();data=[];formats=[];indices=defaultdict(list);populated=formulas=erased_formulas=changed=0;plain_bytes=0
   # Uses each worksheet's saved extent, including blank cells inside the grid.
   for r,row in enumerate(ws.iter_rows()):
    values=[];nfs=[]
    for c,cell in enumerate(row):
     value=cell.value
     if value is None:text=''
     elif isinstance(value,(str,int,float,bool)):text=str(value)
     elif hasattr(value,'isoformat'):text=value.isoformat()
     elif value.__class__.__name__ in ('DataTableFormula','ArrayFormula'):
      text=json.dumps({'excel_object':value.__class__.__name__,'attributes':vars(value)},sort_keys=True)
     elif hasattr(value,'total_seconds'):text=json.dumps({'timedelta_seconds':value.total_seconds()})
     else:raise TypeError(f'Unsupported Excel value type at {ws.title}: {type(value)}')
     fmt=getattr(cell,'number_format','General');values.append(text);nfs.append(fmt)
     if value is not None:
      populated+=1;f=cell.data_type=='f';formulas+=int(f)
      tag=method.get_type(fmt,text)
      changed+=int(tag!=text);erased_formulas+=int(f and tag!=text)
      key=json.dumps([cell.data_type,text,fmt],ensure_ascii=False,separators=(',',':'))
      indices[key].append((r,c))
      # Matched-information sparse baseline, not the paper's dense Markdown baseline.
      plain_bytes+=len(json.dumps([cell.coordinate,cell.data_type,text,fmt],ensure_ascii=False,separators=(',',':')).encode())+1
    data.append(values);formats.append(nfs)
   indexed=[json.loads(key)+[spans(points)] for key,points in indices.items()]
   encoded=json.dumps(indexed,ensure_ascii=False,separators=(',',':'))
   (dest/f'sheet_{number:02d}_lossless_index.json').write_text(encoded,encoding='utf-8')
   # Independent decoding checks each emitted span against the source matrices.
   from openpyxl.utils.cell import range_boundaries
   recovered=0
   for dtype,text,fmt,regions in indexed:
    for region in regions:
     c1,r1,c2,r2=range_boundaries(region)
     for r in range(r1-1,r2):
      for c in range(c1-1,c2):
       if data[r][c]!=text or formats[r][c]!=fmt:raise ValueError('Index round-trip mismatch')
       recovered+=1
   if recovered!=populated:raise ValueError('Index coverage mismatch')
   state='completed';error='';area_count='';area_bytes=''
   try:
    areas=method.aggregate_similar_areas(data,formats)
    area_count=len(areas);blob=json.dumps(areas,ensure_ascii=False,separators=(',',':'))
    area_bytes=len(blob.encode());(dest/f'sheet_{number:02d}_released_areas.json').write_text(blob,encoding='utf-8')
   except RecursionError:
    state='recursion_failure';error='Released recursive connected-component aggregation exceeded its recursion limit; no substitute used.'
   records.append(dict(workbook=path.stem,source_sha256=wbhash,sheet=ws.title,state=ws.sheet_state,
     rows=len(data),columns=len(data[0]) if data else 0,populated_cells=populated,formula_cells=formulas,
     get_type_changed_content_cells=changed,get_type_changed_formula_cells=erased_formulas,
     aggregation_status=state,aggregation_error=error,aggregation_regions=area_count,aggregation_region_bytes=area_bytes,
     sparse_baseline_bytes=plain_bytes,lossless_index_bytes=len(encoded.encode()),
     lossless_index_compression_ratio=plain_bytes/len(encoded.encode()) if encoded else '',
     index_round_trip_cells=recovered,elapsed_seconds=time.perf_counter()-start))
   dump_csv(out/'encoding_by_sheet.csv',records)
   print(path.stem,number,ws.title,populated,state,flush=True)
   del data,formats,indices,indexed
  wb.close()
 if not records:
  raise SystemExit('The matched files contained no readable worksheets. No summary CSV was written.')
 dump_csv(out/'encoding_by_sheet.csv',records)
 print('Finished. Byte sizes are NOT model token counts. No detection or FLAME inference was run.')
if __name__=='__main__':main()
