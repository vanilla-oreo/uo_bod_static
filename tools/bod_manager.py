#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""UO BOD 靜態網站資料工具。純標準庫解析 Excel；修改/產生 Excel 使用 artifact_tool。

指令：
 python tools/bod_manager.py init
 python tools/bod_manager.py export
 python tools/bod_manager.py rewards blacksmith
 python tools/bod_manager.py orders blacksmith
 python tools/bod_manager.py set-reward blacksmith --id R027 --points 650
 python tools/bod_manager.py add-reward inscription --name '新獎勵' --points 300 --category '其他'
 python tools/bod_manager.py delete-reward inscription --id R018
 python tools/bod_manager.py set-order blacksmith --id O0007 --points 650
 python tools/bod_manager.py add-order blacksmith --kind '小訂單' --material Gold --qty 10 --quality 卓越 --points 610
 python tools/bod_manager.py delete-order blacksmith --id O0007
 python tools/bod_manager.py rename-reward blacksmith --id R027 --name '新名稱'
 python tools/bod_manager.py rename-order blacksmith --id O0007 --kind '鎖甲（4種）' --material Gold --qty 20 --quality 普通

更新職業後會自動：產出 Excel（保留五種工作表）、重新輸出 data/bod_data.json；有差異備份。
"""
import argparse, json, re, zipfile, xml.etree.ElementTree as ET, shutil
from datetime import datetime
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
EXCEL=ROOT/'excel'; DATA=ROOT/'data'; BACKUP=ROOT/'backup'
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
JOBS={
 'blacksmith':('鐵匠','Blacksmith','UO_Blacksmith_鐵匠BOD遊戲實測修正版.xlsx','遊戲實測修正版'),
 'tinkering':('鉗工','Tinkering','UO_Tinkering_鉗工BOD官方點數與直接兌換表.xlsx','官方版／遊戲截圖點數核對'),
 'tailoring':('裁縫','Tailoring','UO_Tailoring_裁縫BOD官方點數與直接兌換表.xlsx','官方版／遊戲截圖點數核對'),
 'inscription':('抄錄','Inscription','UO_Inscription_抄錄BOD遊戲實測修正版.xlsx','遊戲實測修正版'),
 'alchemy':('鍊金','Alchemy','UO_Alchemy_鍊金BOD官方點數與直接兌換表.xlsx','官方版／遊戲截圖點數核對'),
 'bowcraft':('製弓／製箭','Bowcraft & Fletching','UO_Bowcraft_製弓BOD官方點數與直接兌換表.xlsx','官方版／未經遊戲截圖核對'),
 'carpentry':('木工','Carpentry','UO_Carpentry_木工BOD官方點數與直接兌換表.xlsx','官方版／遊戲截圖點數核對'),
 'cooking':('烹飪','Cooking','UO_Cooking_烹飪BOD官方點數與直接兌換表.xlsx','官方版／遊戲截圖點數核對')}

def colindex(s):
 n=0
 for c in s:n=n*26+ord(c)-64
 return n-1

def read_xlsx(path):
 """讀取儲存格文字與快取值（不用執行公式）。"""
 with zipfile.ZipFile(path) as z:
  ss=[]
  if 'xl/sharedStrings.xml' in z.namelist():
   for si in ET.fromstring(z.read('xl/sharedStrings.xml')):
    ss.append(''.join(t.text or '' for t in si.iter(NS+'t')))
  wb=ET.fromstring(z.read('xl/workbook.xml'))
  names=[s.get('name') for s in wb.find(NS+'sheets')]
  out={}
  for i,name in enumerate(names,1):
   path_i=f'xl/worksheets/sheet{i}.xml'
   if path_i not in z.namelist():continue
   tree=ET.fromstring(z.read(path_i))
   cells=[]
   for row in tree.iter(NS+'row'):
    rd={}
    for c in row.findall(NS+'c'):
     v=c.find(NS+'v');inline=c.find(NS+'is')
     raw=v.text if v is not None else ''.join((t.text or '') for t in inline.iter(NS+'t')) if inline is not None else None
     if c.get('t')=='s' and raw is not None:raw=ss[int(raw)]
     if c.get('t')=='b' and raw is not None:raw=raw=='1'
     if raw is not None:rd[colindex(re.match(r'[A-Z]+',c.get('r')).group())]=raw
    if rd:
     cols=[None]*(max(rd)+1)
     for j,val in rd.items():cols[j]=val
     cells.append((int(row.get('r')),cols))
   out[name]=cells
  return out

def at(row,i,default=''):
 return row[i] if len(row)>i and row[i] is not None else default

def numeric(s):
 try:
  n=float(str(s).replace(',',''));return int(n) if n.is_integer() else n
 except (TypeError,ValueError):return None

def extract_original(path,slug):
 s=read_xlsx(path); keys=list(s);reward_sheet=next(k for k in keys if '獎勵點數' in k);comp=s['直接兌換對照表']
 rewards=[]
 for r,a in s[reward_sheet]:
  p=numeric(at(a,1));name=at(a,0)
  if r==1 or not name or p is None:continue
  cat=at(a,2);source=next((x for x in a[3:] if isinstance(x,str) and x.startswith('http')),'')
  rewards.append({'id':f'R{len(rewards)+1:03d}','name':str(name),'points':p,'category':str(cat),'source':source,'note':''})
 orders=[]
 for r,a in comp:
  if r<7:continue
  p=numeric(at(a,4));qty=numeric(at(a,2));kind=at(a,0);material=at(a,1);quality=at(a,3)
  if p is None or not kind or qty is None:continue
  orders.append({'id':f'O{len(orders)+1:04d}','kind':str(kind),'material':str(material),'qty':qty,'quality':str(quality),'points':p,'source':'','note':''})
 if not rewards or not orders:raise ValueError(f'{path}: 無法解析獎勵或訂單')
 j=JOBS[slug]
 return {'slug':slug,'name':j[0],'english':j[1],'status':j[3], 'rewards':rewards,'orders':orders}

def extract_managed(path,slug):
 s=read_xlsx(path); rewards=[];orders=[]
 for r,a in s['獎勵點數']:
  if r<2:continue
  p=numeric(at(a,2));rid=at(a,0)
  if not rid or p is None:continue
  rewards.append({'id':str(rid),'name':str(at(a,1)),'points':p,'category':str(at(a,3)),'source':str(at(a,4)),'note':str(at(a,5))})
 for sheet in ('小訂單點數','大型訂單點數'):
  for r,a in s[sheet]:
   if r<2:continue
   p=numeric(at(a,5));oid=at(a,0)
   if not oid or p is None:continue
   orders.append({'id':str(oid),'kind':str(at(a,1)),'material':str(at(a,2)),'qty':numeric(at(a,3)),'quality':str(at(a,4)),'points':p,'source':str(at(a,6)),'note':str(at(a,7))})
 j=JOBS[slug];return {'slug':slug,'name':j[0],'english':j[1],'status':j[3]+'／Python 維護版','rewards':rewards,'orders':orders}

def load(slug):
 path=EXCEL/(slug+'.xlsx')
 if not path.exists():raise FileNotFoundError(f'{path} 不存在；請先執行 init')
 s=read_xlsx(path)
 return extract_managed(path,slug) if '獎勵點數' in s else extract_original(path,slug)

def validate(job):
 ids=set()
 for r in job['rewards']:
  if not r['name'].strip() or not isinstance(r['points'],(int,float)) or r['points']<0:raise ValueError('獎勵名稱／點數無效')
  if r['id'] in ids:raise ValueError('重複獎勵 ID')
  ids.add(r['id'])
 ids=set()
 for o in job['orders']:
  if not o['kind'].strip() or not o['material'].strip() or not o['quality'].strip() or not isinstance(o['points'],(int,float)) or o['points']<0 or not isinstance(o['qty'],(int,float)) or o['qty']<=0:raise ValueError('訂單資料無效')
  if o['id'] in ids:raise ValueError('重複訂單 ID')
  ids.add(o['id'])

def export_all():
 result=[]
 for slug in JOBS:
  job=load(slug);validate(job);result.append(job)
 DATA.mkdir(parents=True,exist_ok=True)
 obj={'schema':2,'source':'以本套工具管理的 Excel 工作表為準；達標由網站動態比較點數','jobs':result}
 p=DATA/'bod_data.json';p.write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding='utf-8')
 print('匯出',p, '|',len(result),'職業',sum(len(j['orders']) for j in result),'訂單',sum(len(j['rewards']) for j in result),'獎勵')
 return result

def backup(path):
 if path.exists():
  BACKUP.mkdir(exist_ok=True)
  dest=BACKUP/(path.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+path.suffix)
  shutil.copy2(path,dest)
  print('Excel 備份：',dest.name)

def save_managed(job):
 """使用 Python 標準函式庫重建五工作表 Excel，Windows 不需額外安裝套件。"""
 from xlsx_writer import write, col
 validate(job)
 intro=[
  ['UO BOD 訂單資料管理',job['name']+' / '+job['english']],
  ['資料狀態',job['status']],
  ['更新方式','使用 tools/bod_manager.py 或 Windows 啟動器進行維護'],
  ['資料主表','獎勵點數、小訂單點數、大型訂單點數'],
  ['對照表規則','訂單點數 >= 獎勵門檻，即為可換'],
  ['篩選排序','網站按剩餘點數由小至大排序'],
  ['提醒','每張訂單只能直接兌換一項獎勵；不含材料成本'],
  ['資料同步','每次成功修改 Excel 後會自動重建 JSON']]
 reward_head=['ID','獎勵名稱','兌換點數','分類','資料來源','備註']
 reward_rows=[reward_head]+[[r['id'],r['name'],r['points'],r.get('category',''),r.get('source',''),r.get('note','')] for r in job['rewards']]
 order_head=['ID','訂單類型','材料／類型','每種件數','品質','直接點數','資料來源','備註']
 def order_rows(items):
  return [order_head]+[[o['id'],o['kind'],o['material'],o['qty'],o['quality'],o['points'],o.get('source',''),o.get('note','')] for o in items]
 small=[o for o in job['orders'] if '小訂單' in o['kind']]
 large=[o for o in job['orders'] if '小訂單' not in o['kind']]
 small_rows=order_rows(small);large_rows=order_rows(large)
 reward_headers=['訂單類型','材料／類型','每種件數','品質','直接點數']+[r['name'] for r in job['rewards']]
 cp=[reward_headers]
 for i,o in enumerate(job['orders'],2):
  formula=[f'=IF($E{i}>=\'獎勵點數\'!$C${r+2},"可換","—")' for r in range(len(job['rewards']))]
  cp.append([o['kind'],o['material'],o['qty'],o['quality'],o['points']]+formula)
 dest=EXCEL/(job['slug']+'.xlsx');backup(dest)
 write(dest,[('使用說明',intro,0,{1:29,2:90}),('獎勵點數',reward_rows,0,{2:48,5:56,6:36}),('小訂單點數',small_rows,0,{2:37,3:27,7:44,8:35}),('大型訂單點數',large_rows,0,{2:37,3:27,7:44,8:35}),('直接兌換對照表',cp,5,{1:34,2:27})])
 print('已更新 Excel：',dest)

def next_id(rows,prefix,width):
 vals=[int(r['id'][1:]) for r in rows if re.fullmatch(prefix+r'\d+',r['id'])]
 return f'{prefix}{(max(vals,default=0)+1):0{width}d}'

def command(args):
 if args.cmd=='init':
  source=Path(args.source).expanduser().resolve() if args.source else ROOT/'source_excels'
  EXCEL.mkdir(parents=True,exist_ok=True)
  for slug,j in JOBS.items():
   orig=source/j[2]
   if not orig.exists():raise FileNotFoundError(f'找不到 {orig}；請將原始 Excel 放在 source_excels 資料夾')
   target=EXCEL/(slug+'.xlsx')
   if target.exists() and not args.force:print('保留已存在的 Excel:',target.name);continue
   if target.exists():backup(target)
   shutil.copy2(orig,target)
  export_all();return
 if args.cmd=='export':export_all();return
 job=load(args.job)
 if args.cmd=='rewards':
  for r in job['rewards']:print(r['id'],r['points'],r['name'])
  return
 if args.cmd=='orders':
  for o in job['orders']:print(o['id'],o['kind'],o['material'],o['qty'],o['quality'],o['points'])
  return
 if args.cmd in ('set-reward','rename-reward','delete-reward'):
  rows=job['rewards'];target=next((r for r in rows if r['id']==args.id),None)
  if target is None:raise ValueError('找不到獎勵 ID: '+args.id)
  if args.cmd=='delete-reward':rows.remove(target)
  elif args.cmd=='rename-reward':target['name']=args.name
  else:
   if args.name is not None:target['name']=args.name
   if args.points is not None:target['points']=args.points
   if args.category is not None:target['category']=args.category
 elif args.cmd=='add-reward':
  job['rewards'].append({'id':next_id(job['rewards'],'R',3),'name':args.name,'points':args.points,'category':args.category or '其他','source':'遊戲實測／手動更新','note':args.note or ''})
 elif args.cmd in ('set-order','rename-order','delete-order'):
  rows=job['orders'];target=next((o for o in rows if o['id']==args.id),None)
  if target is None:raise ValueError('找不到訂單 ID: '+args.id)
  if args.cmd=='delete-order':rows.remove(target)
  else:
   for key in ('kind','material','qty','quality','points'):
    v=getattr(args,key,None)
    if v is not None:target[key]=v
 elif args.cmd=='add-order':
  job['orders'].append({'id':next_id(job['orders'],'O',4),'kind':args.kind,'material':args.material,'qty':args.qty,'quality':args.quality,'points':args.points,'source':'遊戲實測／手動更新','note':args.note or ''})
 else:raise ValueError('未知操作')
 save_managed(job)
 export_all()

if __name__=='__main__':
 import sys
 if len(sys.argv)==1 or '--gui' in sys.argv:
  try:
   from bod_gui import launch
   launch()
  except Exception as exc:
   print('無法啟動視窗介面：',exc)
   print('可改用 python tools/bod_manager.py export')
   if sys.stdin.isatty(): input('按 Enter 結束...')
  sys.exit(0)
 p=argparse.ArgumentParser(description='UO 八大工匠 BOD Excel + JSON 維護工具')
 sub=p.add_subparsers(dest='cmd',required=True)
 i=sub.add_parser('init');i.add_argument('--source');i.add_argument('--force',action='store_true')
 sub.add_parser('export')
 for cmd in ['rewards','orders','set-reward','rename-reward','delete-reward','add-reward','set-order','rename-order','delete-order','add-order']:
  x=sub.add_parser(cmd);x.add_argument('job',choices=JOBS)
  if cmd in ('set-reward','rename-reward','delete-reward','set-order','rename-order','delete-order'):x.add_argument('--id',required=True)
  if cmd in ('add-reward','rename-reward'):x.add_argument('--name',required=True)
  elif cmd=='set-reward':x.add_argument('--name')
  if cmd in ('add-reward','add-order'):x.add_argument('--points',required=True,type=int)
  elif cmd in ('set-reward','set-order'):x.add_argument('--points',type=int)
  if cmd in ('set-reward','add-reward'):x.add_argument('--category')
  if cmd in ('add-reward','add-order'):x.add_argument('--note')
  if cmd in ('add-order',):
   x.add_argument('--kind',required=True);x.add_argument('--material',required=True);x.add_argument('--qty',required=True,type=int);x.add_argument('--quality',required=True)
  elif cmd in ('set-order','rename-order'):
   x.add_argument('--kind');x.add_argument('--material');x.add_argument('--qty',type=int);x.add_argument('--quality')
 args=p.parse_args()
 try:command(args)
 except Exception as e:p.error(str(e))
