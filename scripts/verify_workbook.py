from pathlib import Path
from decimal import Decimal,ROUND_HALF_UP
import zipfile,xml.etree.ElementTree as E,json,csv

root=Path(__file__).resolve().parent
project=root.parent
source=project/'templates/单品利润与退货测算样品.xlsx'
result=project/'examples/merchant-v0.3/填写结果.xlsx'
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
q=lambda x:'{'+ns['m']+'}'+x
old_input={f'{c}{r}' for c in 'CDE' for r in [6,7,*range(10,16),*range(20,26)]}
prep_input=set(json.loads((root/'归集输入地址.json').read_text()))

def read_book(path):
 with zipfile.ZipFile(path) as z:
  assert z.testzip() is None
  names=[s.get('name') for s in E.fromstring(z.read('xl/workbook.xml')).findall('m:sheets/m:sheet',ns)]
  strings=[''.join(t.text or '' for t in s.findall('.//m:t',ns)) for s in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
  xfs=E.fromstring(z.read('xl/styles.xml')).find(q('cellXfs'));sheets=[]
  for i in range(1,5):
   tree=E.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'));cells={};unlocked=set()
   for c in tree.findall('.//m:sheetData/m:row/m:c',ns):
    a=c.get('r');f=c.find(q('f'));v=c.find(q('v'));value=v.text if v is not None else None
    if c.get('t')=='s' and value is not None:value=strings[int(value)]
    elif c.get('t')=='inlineStr':value=''.join(t.text or '' for t in c.findall('.//m:t',ns))
    cells[a]={'value':value,'formula':f.text if f is not None else None,'type':c.get('t')}
    p=xfs[int(c.get('s',0))].find(q('protection'))
    if p is not None and p.get('locked')=='0':unlocked.add(a)
   sheets.append({'cells':cells,'unlocked':unlocked,'protected':tree.find(q('sheetProtection')) is not None})
  return names,sheets

checks=[]
def check(label,condition):
 checks.append({'check':label,'pass':bool(condition)})
def money(x):return x.quantize(Decimal('.01'),rounding=ROUND_HALF_UP)
def d(cells,a):
 c=cells[a];assert c['type'] not in ['s','str','inlineStr','e'] and c['value'] is not None,a
 return Decimal(c['value'])

before_names,before=read_book(source);names,after=read_book(result)
check('四页名称与顺序保留',names==before_names)
for i,(b,a) in enumerate(zip(before,after),1):
 allowed=old_input if i==2 else prep_input if i==4 else set()
 check(f'第{i}页所有公式保留',{k:v['formula'] for k,v in b['cells'].items() if v['formula']}=={k:v['formula'] for k,v in a['cells'].items() if v['formula']})
 check(f'第{i}页非输入内容保留',all(v['value']==a['cells'].get(k,{}).get('value') for k,v in b['cells'].items() if not v['formula'] and k not in allowed))
 check(f'第{i}页保护保留',a['protected'])
 check(f'第{i}页黄色区解锁',a['unlocked']==(old_input if i in [2,3] else prep_input if i==4 else set()))
 check(f'第{i}页没有保存公式错误',all(c['type']!='e' for c in a['cells'].values()))
p=after[3]['cells'];m=after[1]['cells'];cases=[]
check('订单分类总数核对',d(p,'C28')==sum(d(p,f'C{r}') for r in range(29,35)))
count=d(p,'C29')+d(p,'C30')
check('模型内分母核对',d(p,'C35')==count and abs(d(p,'C7')-d(p,'C30')/count)<Decimal('.0000001'))
check('回收成本核对',abs(d(p,'D59')-(d(p,'D57')-d(p,'D58')))<Decimal('.0000001') and abs(d(p,'D42')-d(p,'D59'))<Decimal('.0000001'))
check('广告总额核对',abs(d(p,'C45')+d(p,'D45')-d(p,'C51'))<Decimal('.0000001'))
batch=d(p,'C41')-sum(d(p,f'C{r}') for r in range(42,49))+d(p,'D41')-sum(d(p,f'D{r}') for r in range(42,49))
check('归集整批贡献独立计算',abs(d(p,'C18')-batch)<Decimal('.0000001'))
check('归集单均不提前舍入',abs(d(p,'C19')-batch/count)<Decimal('.0000001'))
check('方案一比例为归集值',abs(d(p,'C7')-d(m,'C7'))<Decimal('.0000001'))
check('方案一均值只复制数值',all(abs(d(m,f'C{r+1}')-d(p,f'C{r}'))<Decimal('.0000001') and m[f'C{r+1}']['formula'] is None for r in range(9,15)) and all(abs(d(m,f'C{r+11}')-d(p,f'D{r}'))<Decimal('.0000001') and m[f'C{r+11}']['formula'] is None for r in range(9,15)))
with (result.parent/'独立虚构订单.csv').open(encoding='utf-8-sig',newline='') as f:ledger=list(csv.DictReader(f))
check('CSV记录数对应归集总数',Decimal(len(ledger))==d(p,'C28'))
ledger_contribution=Decimal(0)
for category,col,countcell in [('正常完成','C','C29'),('退货','D','C30')]:
 rows=[r for r in ledger if r['category']==category]
 check(f'CSV{category}数量一致',Decimal(len(rows))==d(p,countcell))
 total=lambda key:sum((Decimal(r[key]) for r in rows),Decimal(0))
 controls={41:total('retained_yuan'),42:total('net_goods_cost_yuan'),43:total('outbound_yuan')+total('return_freight_yuan'),44:total('platform_yuan'),45:total('ad_yuan'),46:total('packaging_yuan'),47:total('paid_piecework_yuan'),48:total('other_variable_yuan')}
 check(f'CSV{category}全部费用与归集吻合',all(abs(amount-d(p,f'{col}{row}'))<Decimal('.0000001') for row,amount in controls.items()))
 independently=controls[41]-sum(controls[r] for r in range(42,49))
 check(f'CSV{category}贡献列独立核对',independently==total('contribution_yuan'))
 ledger_contribution+=independently
check('CSV逐单账单独立总额吻合',abs(ledger_contribution-batch)<Decimal('.0000001'))
for col in 'CD':
 normal=money(d(m,f'{col}10')-sum(d(m,f'{col}{r}') for r in range(11,16)))
 returned=money(d(m,f'{col}20')-sum(d(m,f'{col}{r}') for r in range(21,26)))
 rate=d(m,f'{col}7');average=money(normal*(1-rate)+returned*rate)
 check(f'{col}方案正常贡献',normal==d(m,f'{col}16'))
 check(f'{col}方案退货贡献',returned==d(m,f'{col}26'))
 check(f'{col}方案单均贡献',average==d(m,f'{col}30'))
 unrounded_normal=d(m,f'{col}10')-sum(d(m,f'{col}{r}') for r in range(11,16))
 unrounded_return=d(m,f'{col}20')-sum(d(m,f'{col}{r}') for r in range(21,26))
 direct_batch=money(unrounded_normal*d(p,'C29')+unrounded_return*d(p,'C30'))
 cases.append({'name':m[f'{col}6']['value'],'normal':str(normal),'returned':str(returned),'average':str(average),'rate':str(rate),'unroundedBatchFromInputs':str(direct_batch)})
data={'checks':checks,'allPass':all(c['pass'] for c in checks),'count':len(checks),'batchContribution':str(batch),'modelOrderCount':str(count),'cases':cases}
(project/'reports').mkdir(exist_ok=True)
(project/'reports/复核结果.json').write_text(json.dumps(data,ensure_ascii=False,indent=2));print(json.dumps({**{k:v for k,v in data.items() if k!='checks'},'pending':[c['check'] for c in checks if not c['pass']]},ensure_ascii=False))

if not data['allPass']:
    raise SystemExit(1)
