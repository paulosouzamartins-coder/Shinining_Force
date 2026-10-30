import openpyxl, json, sys
INV={'Atacar':'atk','Mover':'mov','Curar':'heal','Magia':'magic','Saquear':'loot'}
def sim(v): return str(v or '').strip().lower().startswith('s')
def parse_range(v):
    s=str(v if v is not None else '1').strip()
    if '-' in s:
        a,b=s.split('-'); return [int(a),int(b)]
    n=int(float(s)); return [1,n]          # "2" = ataca de 1 a 2 casas
FX={'perfurar':'perf','área':'area','area':'area','empurrar':'emp','ferir':'ferir','atordoar':'atord','congelar':'gelo','escudo':'esc','retaliar':'ret','cura em área':'carea','cura em area':'carea','salto':'salto','drenar':'dren'}
def parse_fx(v):
    o=[]
    for e in str(v or '').split(';'):
        e=e.strip().lower()
        if not e: continue
        n=1; p=e.rsplit(' ',1)
        if len(p)==2 and p[1].isdigit(): e,n=p[0].strip(),int(p[1])
        assert e in FX,('efeito desconhecido',e); o.append([FX[e],n])
    return o
def lst(v): return [x.strip() for x in str(v or '').split(',') if x.strip()]
def read_data(path):
    wb=openpyxl.load_workbook(path,data_only=True)
    def rows(n): return [r for r in list(wb[n].iter_rows(values_only=True))[1:] if r[0]]
    info={}
    for r in rows('Heróis'):
        info[str(r[0])]={'race':str(r[2] or ''),'cls':str(r[3] or ''),'desc':str(r[4] or ''),'base':str(r[5] or r[3] or ''),
            'hp':int(r[6]),'mov':int(r[7]),'atk':int(r[8]),'df':int(r[9]),'ag':int(r[10]),'fly':sim(r[11])}
    def cd(r,o):  # r[o:] = nome, topo ação, topo valor, base ação, base valor, queimar
        c={'n':str(r[o]),'t':[INV[r[o+1]],int(r[o+2])],'b':[INV[r[o+3]],int(r[o+4])]}
        if sim(r[o+5]): c['k']=1
        fx=[parse_fx(r[o+6] if len(r)>o+6 else ''),parse_fx(r[o+7] if len(r)>o+7 else '')]
        if fx[0] or fx[1]: c['x']=fx
        return c
    decks={}
    for r in rows('Decks Classe'): decks.setdefault(str(r[0]),[]).append(cd(r,2))
    sig={}
    for r in rows('Assinaturas'):
        assert str(r[0]) not in sig,('assinatura duplicada',r[0]); sig[str(r[0])]=cd(r,2)
    for k,h in info.items():
        assert h['base'] in decks,('classe base sem deck',k,h['base']); assert k in sig,('herói sem assinatura',k)
    weapons=[]
    for r in rows('Armas'):
        weapons.append({'n':str(r[0]),'t':str(r[1]),'r':parse_range(r[2]),'a':int(r[3]),'cls':lst(r[6]),'only':lst(r[7])})
    monsters={}
    for r in rows('Cartas Monstros'):
        c={'n':str(r[2]),'t':[INV[r[3]],int(r[4])],'b':[INV[r[5]],int(r[6])]}
        if sim(r[7]): c['s']=1
        monsters.setdefault(str(r[0]),[]).append(c)
    for k,v in monsters.items(): assert any(c.get('s') for c in v),('sem carta de embaralhamento',k)
    rg={}; sp={}
    for r in rows('Alcance Monstros'):
        if r[1] is not None: rg[str(r[0])]=int(r[1])
        sp[str(r[0])]=str(r[2] or '')
    return {'info':info,'decks':decks,'sig':sig,'weapons':weapons,'monsters':monsters,'range':rg,'sp':sp}
def build(xlsx='dados.xlsx',tpl='tpl.html',out='catacumba.html'):
    cards=read_data(xlsx)
    t=open(tpl).read()
    assert '/*__MONSTERS__*/[]' in t and '/*__CARDS__*/{}' in t
    t=t.replace('/*__MONSTERS__*/[]',open('monsters.json').read()).replace('/*__CARDS__*/{}',json.dumps(cards,ensure_ascii=False,separators=(',',':')))
    open(out,'w').write(t); return cards
if __name__=='__main__':
    c=build(*(sys.argv[1:2] or [])); print(len(c['info']),'heróis',len(c['decks']),'decks',len(c['weapons']),'armas',len(c['monsters']),'monstros')
