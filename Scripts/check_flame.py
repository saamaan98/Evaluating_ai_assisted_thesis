import argparse,importlib.util,json,urllib.request
from pathlib import Path

if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--output',type=Path,default=Path('flame_availability.json'));a=p.parse_args()
 url='https://api.github.com/repos/microsoft/prose-benchmarks/contents/FLAME'
 data=json.loads(urllib.request.urlopen(url,timeout=30).read())
 names=[r['name'] for r in data]
 result={'checked_url':url,'files':names,'inference_executed':False,
         'torch_installed':importlib.util.find_spec('torch') is not None,
         'transformers_installed':importlib.util.find_spec('transformers') is not None,
         'conclusion':'Directory inventory only. A verified FLAME checkpoint, matching tokenizer, inference code and approved formula repair pairs are still required. This script is not FLAME.'}
 a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(result,indent=2));print(json.dumps(result,indent=2))
