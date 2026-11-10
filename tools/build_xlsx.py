import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.utils import get_column_letter as L

items = json.load(open(sys.argv[1]))
out = sys.argv[2]
STAGES = {'F':'Fundação (civil)','S':'Suporte / estrutura','E':'Montagem do equipamento','C':'Cabos e conexões'}
ORDER = ['F','S','E','C']
AREAS = {'LT':'Entrada de linha','RE':'Banco de reatores','BA':'Vão do lado da barra','SX':'Serviços auxiliares'}
FASES = {'VM':'VM','BR':'BR','AZ':'AZ','RS':'Reserva','3F':'Comum (3 fases)'}

F = 'Arial'
fN = Font(name=F, size=10); fB = Font(name=F, size=10, bold=True)
fH = Font(name=F, size=10, bold=True, color='FFFFFF'); fT = Font(name=F, size=14, bold=True)
fIn = Font(name=F, size=10, color='0000FF'); fMono = Font(name='Consolas', size=10)
hdrFill = PatternFill('solid', fgColor='2F3B42'); inFill = PatternFill('solid', fgColor='FFF2CC')
thin = Side(style='thin', color='C9D0D3'); bd = Border(bottom=thin)
okF = PatternFill('solid', fgColor='C6EFCE'); prF = PatternFill('solid', fgColor='FFF2A8'); tdF = PatternFill('solid', fgColor='E7EBED')

wb = Workbook()
# ---------------- Instruções ----------------
ws = wb.active; ws.title = 'Instruções'
rows = [
 ('Planilha de avanço – SE Campo Formoso II 500 kV (ampliação, Lote 05)', fT),
 ('Uso: preencha somente a coluna "% Avanço" (amarela) da aba Avanço. As demais colunas de % e status são calculadas.', fN),
 ('Depois de salvar, carregue este arquivo no modelo 3D (botão "Carregar planilha de avanço"). O 3D lê as colunas "ID Atividade" e "% Avanço".', fN),
 ('Datas, responsável e observações são opcionais e não alteram o 3D.', fN),
 ('', fN),
 ('Formato do ID do equipamento: ÁREA-FASE-TIPO-NN  (ex.: LT-VM-TC-01).  ID da atividade = ID do equipamento + ".etapa" (ex.: LT-VM-TC-01.F).', fB),
]
for i,(t,f) in enumerate(rows,1):
    c = ws.cell(i,1,t); c.font = f
r = 8
def table(title, mapping, r):
    ws.cell(r,1,title).font = fB; r += 1
    for k,v in mapping.items():
        ws.cell(r,1,k).font = fMono; ws.cell(r,2,v).font = fN; r += 1
    return r + 1
r = table('Áreas', AREAS, r)
r = table('Fases', FASES, r)
r = table('Etapas (sufixo da atividade)', {'.'+k:v for k,v in STAGES.items()}, r)
tipos = {}
for it in items: tipos[it['code']] = (it['tipo'], ''.join(it['stages']))
ws.cell(r,1,'Tipos (código – descrição – etapas aplicáveis)').font = fB; r += 1
for k,(t,s) in tipos.items():
    ws.cell(r,1,k).font = fMono; ws.cell(r,2,t).font = fN; ws.cell(r,3,'.'+' .'.join(s)).font = fMono; r += 1
r += 1
ws.cell(r,1,'Exemplo de preenchimento (apenas ilustrativo – não faz parte do controle)').font = fB; r += 1
ex_h = ['ID Atividade','Etapa','% Avanço','Status','Início real','Término real','Responsável','Observações']
ex_v = ['LT-VM-TC-01.F','Fundação (civil)',0.6,'Em andamento','05/11/2026','','Equipe civil A','Concretagem do bloco pendente']
for j,(h,v) in enumerate(zip(ex_h,ex_v),1):
    ws.cell(r,j,h).font = fH; ws.cell(r,j).fill = hdrFill
    c = ws.cell(r+1,j,v); c.font = fN
    if j==3: c.number_format='0%'; c.fill=inFill; c.font=fIn
ws.column_dimensions['A'].width = 22; ws.column_dimensions['B'].width = 46; ws.column_dimensions['C'].width = 14
for col in 'DEFGH': ws.column_dimensions[col].width = 16

# ---------------- Avanço ----------------
wa = wb.create_sheet('Avanço')
H = ['ID Atividade','ID Equipamento','Área','Fase','Tipo','Descrição','Etapa','% Avanço','Status',
     'Início previsto','Término previsto','Início real','Término real','Responsável','Observações']
W = [18,16,20,14,30,46,24,11,14,13,13,13,13,18,36]
for j,(h,w) in enumerate(zip(H,W),1):
    c = wa.cell(1,j,h); c.font = fH; c.fill = hdrFill; c.alignment = Alignment(vertical='center', wrap_text=True)
    wa.column_dimensions[L(j)].width = w
acts = []
for it in items:
    for s in ORDER:
        if s in it['stages']: acts.append((it,s))
n = len(acts); last = n+1
for i,(it,s) in enumerate(acts, start=2):
    vals = [f"{it['id']}.{s}", it['id'], AREAS[it['area']], FASES[it['ph']], it['tipo'], it['name'], STAGES[s]]
    for j,v in enumerate(vals,1):
        c = wa.cell(i,j,v); c.font = fMono if j<=2 else fN; c.border = bd
    c = wa.cell(i,8,0); c.number_format = '0%'; c.fill = inFill; c.font = fIn; c.border = bd
    c = wa.cell(i,9,f'=IF(H{i}>=1,"Concluído",IF(H{i}>0,"Em andamento","Não iniciada"))'); c.font = fN; c.border = bd
    for j in range(10,14): wa.cell(i,j).number_format = 'DD/MM/YYYY'; wa.cell(i,j).font = fIn; wa.cell(i,j).border = bd
    for j in (14,15): wa.cell(i,j).font = fIn; wa.cell(i,j).border = bd
dv = DataValidation(type='decimal', operator='between', formula1='0', formula2='1', allow_blank=True,
                    showErrorMessage=True, errorTitle='% Avanço', error='Digite um valor entre 0% e 100%.')
wa.add_data_validation(dv); dv.add(f'H2:H{last}')
dvd = DataValidation(type='date', operator='greaterThan', formula1='40000', allow_blank=True)
wa.add_data_validation(dvd); dvd.add(f'J2:M{last}')
rng = f'I2:I{last}'
wa.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"Concluído"'], fill=okF))
wa.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"Em andamento"'], fill=prF))
wa.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"Não iniciada"'], fill=tdF))
wa.freeze_panes = 'C2'; wa.auto_filter.ref = f'A1:O{last}'

AV = lambda col: f"Avanço!${col}$2:${col}${last}"

# ---------------- Equipamentos ----------------
we = wb.create_sheet('Equipamentos')
H2 = ['ID Equipamento','Área','Fase','Tipo','Descrição','X (m)'] + [STAGES[s] for s in ORDER] + ['Avanço do item','Status']
W2 = [16,20,14,30,46,8,13,13,13,13,12,14]
for j,(h,w) in enumerate(zip(H2,W2),1):
    c = we.cell(1,j,h); c.font = fH; c.fill = hdrFill; c.alignment = Alignment(vertical='center', wrap_text=True)
    we.column_dimensions[L(j)].width = w
we.row_dimensions[1].height = 30
for i,it in enumerate(items, start=2):
    for j,v in enumerate([it['id'],AREAS[it['area']],FASES[it['ph']],it['tipo'],it['name'],it.get('x')],1):
        c = we.cell(i,j,v); c.font = fMono if j==1 else fN; c.border = bd
    we.cell(i,6).number_format = '0.0'
    for k,s in enumerate(ORDER):
        col = 7+k
        c = we.cell(i,col, f'=IFERROR(AVERAGEIFS({AV("H")},{AV("B")},$A{i},{AV("G")},{L(col)}$1),"–")')
        c.number_format = '0%'; c.font = fN; c.border = bd; c.alignment = Alignment(horizontal='center')
    c = we.cell(i,11, f'=IFERROR(AVERAGEIFS({AV("H")},{AV("B")},$A{i}),0)'); c.number_format='0%'; c.font=fB; c.border=bd
    c = we.cell(i,12, f'=IF(K{i}>=1,"Concluído",IF(K{i}>0,"Em andamento","Não iniciada"))'); c.font=fN; c.border=bd
le = len(items)+1
we.conditional_formatting.add(f'K2:K{le}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='8FD6A0'))
for rr in (f'L2:L{le}',):
    we.conditional_formatting.add(rr, CellIsRule(operator='equal', formula=['"Concluído"'], fill=okF))
    we.conditional_formatting.add(rr, CellIsRule(operator='equal', formula=['"Em andamento"'], fill=prF))
    we.conditional_formatting.add(rr, CellIsRule(operator='equal', formula=['"Não iniciada"'], fill=tdF))
we.freeze_panes = 'B2'; we.auto_filter.ref = f'A1:L{le}'

# ---------------- Resumo ----------------
wr = wb.create_sheet('Resumo', 1)
wr['A1'] = 'Resumo do avanço – SE Campo Formoso II'; wr['A1'].font = fT
wr['A2'] = 'Data de referência'; wr['A2'].font = fB
wr['B2'].number_format = 'DD/MM/YYYY'; wr['B2'].fill = inFill; wr['B2'].font = fIn
wr['C2'] = '← preencha a data da medição'; wr['C2'].font = Font(name=F, size=9, italic=True, color='5D6A71')
wr['A3'] = 'Percentuais = média simples das atividades (sem pesos).'; wr['A3'].font = Font(name=F, size=9, italic=True, color='5D6A71')
hdr = ['Etapa','Atividades','Concluídas','Em andamento','Não iniciadas','% médio']
for j,h in enumerate(hdr,1):
    c = wr.cell(5,j,h); c.font = fH; c.fill = hdrFill
r0 = 6
for k,s in enumerate(ORDER):
    r = r0+k; wr.cell(r,1,STAGES[s]).font = fN
    wr.cell(r,2,f'=COUNTIF({AV("G")},$A{r})')
    wr.cell(r,3,f'=COUNTIFS({AV("G")},$A{r},{AV("I")},"Concluído")')
    wr.cell(r,4,f'=COUNTIFS({AV("G")},$A{r},{AV("I")},"Em andamento")')
    wr.cell(r,5,f'=COUNTIFS({AV("G")},$A{r},{AV("I")},"Não iniciada")')
    wr.cell(r,6,f'=IFERROR(AVERAGEIF({AV("G")},$A{r},{AV("H")}),0)').number_format = '0.0%'
    for j in range(2,7): wr.cell(r,j).font = fN; wr.cell(r,j).border = bd
rt = r0+len(ORDER)
wr.cell(rt,1,'Total da obra').font = fB
for j,col in zip(range(2,6),'BCDE'):
    c = wr.cell(rt,j,f'=SUM({col}{r0}:{col}{rt-1})'); c.font = fB
c = wr.cell(rt,6,f'=IFERROR(AVERAGE({AV("H")}),0)'); c.font = fB; c.number_format = '0.0%'
wr.conditional_formatting.add(f'F{r0}:F{rt}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='8FD6A0'))

ra = rt + 3
wr.cell(ra-1,1,'% médio por área e etapa').font = fB
wr.cell(ra,1,'Área').font = fH; wr.cell(ra,1).fill = hdrFill
for k,s in enumerate(ORDER):
    c = wr.cell(ra,2+k,STAGES[s]); c.font = fH; c.fill = hdrFill; c.alignment = Alignment(wrap_text=True)
c = wr.cell(ra,2+len(ORDER),'Total da área'); c.font = fH; c.fill = hdrFill
wr.row_dimensions[ra].height = 30
for i,(k,a) in enumerate(AREAS.items()):
    r = ra+1+i; wr.cell(r,1,a).font = fN
    for kk,s in enumerate(ORDER):
        col = 2+kk
        c = wr.cell(r,col,f'=IFERROR(AVERAGEIFS({AV("H")},{AV("C")},$A{r},{AV("G")},{L(col)}${ra}),"–")'); c.number_format='0%'; c.font=fN; c.alignment=Alignment(horizontal='center')
    c = wr.cell(r,2+len(ORDER),f'=IFERROR(AVERAGEIF({AV("C")},$A{r},{AV("H")}),"–")'); c.number_format='0%'; c.font=fB; c.alignment=Alignment(horizontal='center')
wr.column_dimensions['A'].width = 26
for col in 'BCDEFG': wr.column_dimensions[col].width = 15

wb.save(out)
print('ok', len(items), 'itens', n, 'atividades')

# ---------------- Quantitativo (Prev./Real./Saldo/Avanço por etapa) ----------------
from openpyxl import load_workbook
from openpyxl.styles import Alignment as _Al
wb = load_workbook(out)
wq = wb.create_sheet('Quantitativo', 2)
NOQ = ['CB','TB','CN','CG','CAD','ALB']
cnt = {}
for it in items:
    if it['code'] not in NOQ: cnt[it['tipo']] = cnt.get(it['tipo'],0)+1
rows_t = [t for t,_ in sorted(cnt.items(), key=lambda x:-x[1])]
for k in ['CB','TB','CN','CAD','CG','ALB']:
    t = next((it['tipo'] for it in items if it['code']==k), None)
    if t: rows_t.append(t)
appl = {t:set() for t in rows_t}
for it in items:
    if it['tipo'] in appl: appl[it['tipo']].update(it['stages'])
STQ = {'F':'Fundação','S':'Suporte','E':'Equipamento','C':'Conexões'}
wq['A1'] = 'Quantitativo do escopo'; wq['A1'].font = fT
wq['A2'] = 'Prev. = previsto · Real. = atividades com 100 % · Saldo = Prev. − Real. · Avanço = média do % informado na aba Avanço'
wq['A2'].font = Font(name=F, size=9, italic=True, color='5D6A71')
navy = PatternFill('solid', fgColor='06172B'); sub = PatternFill('solid', fgColor='E7EFFE')
thick = Side(style='medium', color='06172B')
wq.merge_cells('A4:A5'); c = wq['A4']; c.value = 'Tipo'; c.font = fH; c.fill = navy; c.alignment = _Al(vertical='center')
wq['A5'].fill = navy
for k,s in enumerate(ORDER):
    c0 = 2 + k*4
    wq.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0+3)
    c = wq.cell(4, c0, f'{STQ[s]} (.{s})'); c.font = fH; c.fill = navy; c.alignment = _Al(horizontal='center')
    for j in range(4): wq.cell(4, c0+j).fill = navy
    for j,h in enumerate(['Prev.','Real.','Saldo','Avanço']):
        c = wq.cell(5, c0+j, h); c.font = fB; c.fill = sub; c.alignment = _Al(horizontal='center')
r0 = 6
E_ = AV('E'); G_ = AV('G'); H_ = AV('H')
for i,t in enumerate(rows_t):
    r = r0 + i
    wq.cell(r,1,t).font = fN; wq.cell(r,1).border = bd
    for k,s in enumerate(ORDER):
        c0 = 2 + k*4; P,R_,S_,A_ = [L(c0+j) for j in range(4)]
        if s not in appl[t]:
            wq.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0+3)
            c = wq.cell(r, c0, 'não se aplica'); c.font = Font(name=F, size=9, color='9AA5B4'); c.alignment = _Al(horizontal='center')
            continue
        stg = STAGES[s]
        wq.cell(r,c0,  f'=COUNTIFS({E_},$A{r},{G_},"{stg}")')
        wq.cell(r,c0+1,f'=COUNTIFS({E_},$A{r},{G_},"{stg}",{H_},">=1")')
        wq.cell(r,c0+2,f'={P}{r}-{R_}{r}')
        wq.cell(r,c0+3,f'=IFERROR(AVERAGEIFS({H_},{E_},$A{r},{G_},"{stg}"),0)').number_format = '0%'
        for j in range(4):
            c = wq.cell(r,c0+j); c.font = fB if j<3 else fN; c.alignment = _Al(horizontal='center'); c.border = bd
rt = r0 + len(rows_t)
wq.cell(rt,1,'Total').font = fB
for k,s in enumerate(ORDER):
    c0 = 2 + k*4; P,R_,S_,A_ = [L(c0+j) for j in range(4)]
    wq.cell(rt,c0,  f'=SUM({P}{r0}:{P}{rt-1})')
    wq.cell(rt,c0+1,f'=SUM({R_}{r0}:{R_}{rt-1})')
    wq.cell(rt,c0+2,f'={P}{rt}-{R_}{rt}')
    wq.cell(rt,c0+3,f'=IFERROR(AVERAGEIF({G_},"{STAGES[s]}",{H_}),0)').number_format = '0%'
for col in range(1, 2+len(ORDER)*4):
    c = wq.cell(rt,col); c.font = fB; c.border = Border(top=thick); c.alignment = _Al(horizontal='center' if col>1 else 'left')
for k in range(len(ORDER)):
    rng = f'{L(5+k*4)}{r0}:{L(5+k*4)}{rt}'
    wq.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['1'], fill=okF))
    wq.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['0.0001','0.9999'], fill=prF))
wq.column_dimensions['A'].width = 44
for col in range(2, 2+len(ORDER)*4): wq.column_dimensions[L(col)].width = 9.5
wq.freeze_panes = 'B6'
wb.save(out)
print('quantitativo', len(rows_t), 'tipos')
