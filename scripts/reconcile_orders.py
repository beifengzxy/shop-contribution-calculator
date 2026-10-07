from pathlib import Path
from decimal import Decimal as D, ROUND_HALF_UP
import csv,json,zipfile,xml.etree.ElementTree as E

PROJECT=Path(__file__).resolve().parent.parent
OWN=PROJECT/'examples/merchant-v0.3'
SRC=PROJECT/'templates/单品利润与退货测算样品.xlsx'
TARGET=OWN/'填写结果.xlsx'
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';Q=lambda s:'{'+NS+'}'+s
def package(path):
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None
        return {n:z.read(n) for n in z.namelist()}
def cells(files,i):return {x.get('r'):x for x in E.fromstring(files[f'xl/worksheets/sheet{i}.xml']).findall('.//'+Q('c'))}
def contents(c,files):
    if c is None:return None
    t=c.get('t');v=c.find(Q('v'))
    if t=='inlineStr':return ''.join(x.text or '' for x in c.findall('.//'+Q('t')))
    if v is None or v.text is None:return None
    if t=='s':return ''.join(x.text or '' for x in E.fromstring(files['xl/sharedStrings.xml']).findall(Q('si'))[int(v.text)].iter(Q('t')))
    if t in ['str','e']:return v.text
    if t=='b':return v.text=='1'
    return D(v.text)
source=package(SRC);target=package(TARGET)
data={i:cells(target,i) for i in range(1,5)}
def value(i,addr):return contents(data[i].get(addr),target)
rows=list(csv.DictReader((OWN/'独立虚构订单.csv').open(encoding='utf-8-sig')))
moneykeys=[k for k in rows[0] if k.endswith('_yuan')]
for r in rows:
    for k in moneykeys:r[k]=D(r[k])
expensekeys=['net_goods_cost_yuan','outbound_yuan','return_freight_yuan','platform_yuan','ad_yuan','packaging_yuan','paid_piecework_yuan','other_variable_yuan']
linechecks=[]
for r in rows:
    retained=r['gross_yuan']-r['refund_yuan']
    goods=r['goods_issued_yuan']-r['recovered_cost_yuan']
    contribution=retained-sum(r[k] for k in expensekeys)
    assert retained==r['retained_yuan'] and goods==r['net_goods_cost_yuan'] and contribution==r['contribution_yuan'],r['id']
    linechecks.append({'id':r['id'],'category':r['category'],'independent_contribution':str(contribution)})
groups={c:[r for r in rows if r['category']==c] for c in ['正常完成','退货','未发货取消','仅退款未退货','换货']}
def total(rs,key):return sum((r[key] for r in rs),D(0))
model=groups['正常完成']+groups['退货'];outside=[r for r in rows if r not in model]
checks=[]
def eq(label,actual,expected,tol=D('0.000000001')):
    assert isinstance(actual,D) and abs(actual-expected)<=tol,(label,actual,expected)
    checks.append({'check':label,'actual':str(actual),'expected':str(expected)})
for col,rs in [('C',groups['正常完成']),('D',groups['退货'])]:
    mapping={41:'retained_yuan',42:'net_goods_cost_yuan',44:'platform_yuan',45:'ad_yuan',46:'packaging_yuan',47:'paid_piecework_yuan',48:'other_variable_yuan'}
    for row,key in mapping.items():eq(f'归集{col}{row}',value(4,f'{col}{row}'),total(rs,key))
    eq(f'归集{col}43',value(4,f'{col}43'),total(rs,'outbound_yuan')+total(rs,'return_freight_yuan'))
eq('归集D57退回原成本',value(4,'D57'),total(groups['退货'],'goods_issued_yuan'))
eq('归集D58可回收成本',value(4,'D58'),total(groups['退货'],'recovered_cost_yuan'))
eq('归集C51模型广告',value(4,'C51'),total(model,'ad_yuan'))
eq('归集C18模型整批贡献',value(4,'C18'),total(model,'contribution_yuan'))
eq('归集C19模型未舍入单均',value(4,'C19'),total(model,'contribution_yuan')/len(model))
eq('退货率',value(4,'C7'),D(3)/21)
eq('21笔模型订单',value(4,'C35'),D(21))
roundcent=lambda x:x.quantize(D('.01'),rounding=ROUND_HALF_UP)
normal=sum((r['contribution_yuan'] for r in groups['正常完成']),D(0))/18
returned=sum((r['contribution_yuan'] for r in groups['退货']),D(0))/3
scene1=roundcent((roundcent(normal)*18+roundcent(returned)*3)/21)
eq('方案一正常贡献',value(2,'C16'),roundcent(normal));eq('方案一退货贡献',value(2,'C26'),roundcent(returned));eq('方案一单均',value(2,'C30'),scene1)
normal2=normal-D('4.75');returned2=returned+D('.05')
scene2=roundcent((roundcent(normal2)*18+roundcent(returned2)*3)/21)
eq('方案二正常贡献',value(2,'D16'),roundcent(normal2));eq('方案二退货贡献',value(2,'D26'),roundcent(returned2));eq('方案二单均',value(2,'D30'),scene2)
scenario2_batch=total(model,'contribution_yuan')-D('4.75')*18+D('.05')*3
preserve=[]
for i in range(1,5):
    old=cells(source,i);new=data[i]
    oldformula={k:c.find(Q('f')).text for k,c in old.items() if c.find(Q('f')) is not None}
    newformula={k:c.find(Q('f')).text for k,c in new.items() if c.find(Q('f')) is not None}
    assert oldformula==newformula,(i,'公式改变')
    if i in [1,3]:
        differences=[]
        for addr in set(old)|set(new):
            if addr in oldformula:continue
            a=contents(old.get(addr),source);b=contents(new.get(addr),target)
            if a in ['',None] and b in ['',None]:continue
            if a!=b:differences.append(addr)
        assert not differences,(i,'说明或示例内容改变',differences)
    xs=E.fromstring(source['xl/styles.xml']).find(Q('cellXfs'));xt=E.fromstring(target['xl/styles.xml']).find(Q('cellXfs'))
    def unlocked(xfs,cs):
        return {a for a,c in cs.items() if (p:=xfs[int(c.get('s',0))].find(Q('protection'))) is not None and p.get('locked','1')=='0'}
    before=unlocked(xs,old);after=unlocked(xt,new);assert before==after,(i,'解锁范围改变')
    os=E.fromstring(source[f'xl/worksheets/sheet{i}.xml']);ns=E.fromstring(target[f'xl/worksheets/sheet{i}.xml'])
    assert os.find(Q('sheetProtection')).attrib==ns.find(Q('sheetProtection')).attrib,(i,'保护属性改变')
    source_fills=E.fromstring(source['xl/styles.xml']).find(Q('fills'));target_fills=E.fromstring(target['xl/styles.xml']).find(Q('fills'))
    for addr in before:
        sf=source_fills[int(xs[int(old[addr].get('s',0))].get('fillId',0))]
        tf=target_fills[int(xt[int(new[addr].get('s',0))].get('fillId',0))]
        def fillcolor(f):
            pattern=f.find(Q('patternFill'))
            return (pattern.get('patternType'),pattern.find(Q('fgColor')).get('rgb'))
        assert fillcolor(sf)==fillcolor(tf),(i,addr,'黄色输入色改变')
    for addr in set(old)|set(new):
        if addr in oldformula or addr in before:continue
        a=contents(old.get(addr),source);b=contents(new.get(addr),target)
        if a in ['',None] and b in ['',None]:continue
        assert a==b,(i,addr,'非输入说明内容改变')
    preserve.append({'sheet':i,'all_original_formulas_unchanged':True,'unlocked_cells_unchanged':len(after),'yellow_inputs_unchanged':True,'protection_unchanged':True,'non_input_static_contents_unchanged':True})
for addr in ['E6','E7',*[f'E{x}' for x in range(10,16)],*[f'E{x}' for x in range(20,26)]]:assert value(2,addr) in [None,''],('方案三应空白',addr)
for addr in ['C7',*[f'C{x}' for x in range(10,16)],*[f'C{x}' for x in range(20,26)]]:assert data[2][addr].find(Q('f')) is None,('值粘贴包含公式',addr)
rounding={
  'scenario1_direct_batch':str(total(model,'contribution_yuan')),
  'scenario1_unrounded_mean':str(total(model,'contribution_yuan')/21),
  'scenario1_rounded_type_contributions_times_counts':str(roundcent(normal)*18+roundcent(returned)*3),
  'scenario1_display_mean_times21':str(scene1*21),
  'scenario1_display_times21_minus_direct':str(scene1*21-total(model,'contribution_yuan')),
  'scenario2_hypothetical_direct_batch':str(scenario2_batch),
  'scenario2_unrounded_mean':str(scenario2_batch/21),
  'scenario2_rounded_type_contributions_times_counts':str(roundcent(normal2)*18+roundcent(returned2)*3),
  'scenario2_display_mean_times21':str(scene2*21),
  'scenario2_display_times21_minus_direct':str(scene2*21-scenario2_batch),
  'hypothetical_same_unit_economics_volume_increase_to_keep_batch_contribution':str(total(model,'contribution_yuan')/scenario2_batch-1)
}
summary={'all_data_fictional':True,'method':'独立读取CSV，Decimal逐单收入减现金变动费用及商品净成本；不使用工作簿公式计算归集账单核对',
'counts':{k:len(v) for k,v in groups.items()},'line_checks':linechecks,'checks':checks,'preservation':preserve,'rounding':rounding,
'outside_contribution':str(total(outside,'contribution_yuan')),'ad_invoice_total':str(total(rows,'ad_yuan')),'outside_ad':str(total(outside,'ad_yuan'))}
(PROJECT/'reports').mkdir(exist_ok=True)
(PROJECT/'reports/独立逐单核对.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2))
print(json.dumps({'checked_orders':len(rows),'checks':len(checks),'preservation':preserve,'rounding':rounding},ensure_ascii=False,indent=2))
