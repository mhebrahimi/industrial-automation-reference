import openpyxl, json, sys, re, hashlib
X=sys.argv[1]; OUT=sys.argv[2]; REP=sys.argv[3] if len(sys.argv)>3 else None
w=openpyxl.load_workbook(X,data_only=True)
def key(h): return str(h).split('\n')[-1].strip()
f=w['Full']; hdr=[key(c.value) for c in f[2]]
recs=[]
for r in f.iter_rows(min_row=3,values_only=True):
    if not r[1] or not str(r[1]).startswith('IA-'): continue
    d={k:(str(v) if isinstance(v,(int,float)) and k in('hs_ir','cn_eu','isic_rev4','isic_rev5') else v) for k,v in zip(hdr,r)}
    for k in ('key_features','standards_reference','target_industries'):
        if d.get(k): d[k+'_list']=[t.strip() for t in re.split(r'\s*\|\s*',str(d[k])) if t.strip()]
    d['secondary_codes']=sorted(set(re.findall(r'(?<!\d)\d{8}(?!\d)',str(d.get('classification_note') or '')))-{d.get('hs_ir')})
    recs.append(d)
U=w['Unique_HS_Codes']; codes=[]; rel='کد اصلی'
for r in U.iter_rows(min_row=3,values_only=True):
    if r[0] and isinstance(r[0],str) and not str(r[2] or '').isdigit():
        if 'فرعی' in r[0]: rel=None
        continue
    if not (r[2] and str(r[2]).isdigit() and len(str(r[2]))==8): continue
    if rel=='کد اصلی':
        codes.append(dict(code=str(r[2]),hs6=str(r[3]),relation='کد اصلی',cn_eu=r[4],desc_1405=r[6],unit=r[7],duty_pct=r[8],records=[x.strip() for x in str(r[9] or '').split('|') if x.strip()]))
    else:
        codes.append(dict(code=str(r[2]),hs6=str(r[3]),relation=r[4],desc_1405=r[6],unit=r[7],duty_pct=r[8],records=[x.strip() for x in str(r[9] or '').split('|') if x.strip()]))
def table(ws,start):
    rows=list(ws.iter_rows(min_row=start,values_only=True)); h=rows[0]
    return [dict(zip(h,x)) for x in rows[1:] if any(v is not None for v in x)]
L=w['Lists']; lists={}
for col in zip(*L.iter_rows(min_row=2,values_only=True)):
    if col[0]: lists[col[0]]=[x for x in col[1:] if x]
guide={r[0]:r[1] for r in w['Guide'].iter_rows(values_only=True) if r[0] and r[1]}
kb=dict(meta=dict(version=guide.get('نسخه'),date=guide.get('تاریخ انتشار'),source_file=X.split('/')[-1],sha256=hashlib.sha256(open(X,'rb').read()).hexdigest(),publisher='انجمن صنفی کارفرمایی اتوماسیون صنعتی ایران'),
  guide=guide,records=recs,hs_codes=codes,taxonomy=table(w['Taxonomy'],2),lists=lists,
  dictionary=[x for x in table(w['dictionary'],2) if x.get('فیلد (key)') and x.get('نام فارسی')],
  merged_records=table(w['Merged_Records'],2),hs_isic=table(w['HS_ISIC_Combinations'],2))
if REP: kb['company_code_evaluation']=[{k:o[k] for k in ('n','code','df','ref','status','an','duty')} for o in json.load(open(REP))]
json.dump(kb,open(OUT,'w'),ensure_ascii=False,indent=1,default=str)
print(len(recs),len([c for c in codes if c['relation']=='کد اصلی']),len(codes),len(kb['taxonomy']),len(kb.get('company_code_evaluation',[])))
