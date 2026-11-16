import datetime, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment

OUT=sys.argv[1]
F='Arial'
RED='7A0A12'
H_FILL=PatternFill('solid',fgColor=RED); H_FONT=Font(name=F,bold=True,color='FFFFFF')
IN_FILL=PatternFill('solid',fgColor='FFF2CC'); IN_FONT=Font(name=F,color='0000FF')
LINK=Font(name=F,color='008000'); BASE=Font(name=F); BOLD=Font(name=F,bold=True)
TITLE=Font(name=F,bold=True,size=16,color=RED); NOTE=Font(name=F,italic=True,size=9,color='666666')
thin=Side(style='thin',color='D9D9D9'); BOX=Border(top=thin,bottom=thin,left=thin,right=thin)
BRL='"R$" #,##0.00;[Red]-"R$" #,##0.00;"-"'; PCT='0.0%;[Red]-0.0%;"-"'; MES='mmm/yyyy'; DATA='dd/mm/yyyy'

wb=Workbook()
def sheet(name,first=False):
    ws=wb.active if first else wb.create_sheet(name); ws.title=name; ws.sheet_view.showGridLines=False; return ws
def head(ws,row,labels,col=1):
    for i,l in enumerate(labels):
        c=ws.cell(row=row,column=col+i,value=l);c.fill=H_FILL;c.font=H_FONT;c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);c.border=BOX
def inp(c,v=None,fmt=None):
    if v is not None:c.value=v
    c.fill=IN_FILL;c.font=IN_FONT;c.border=BOX
    if fmt:c.number_format=fmt
def fx(c,v,fmt=None,font=BASE):
    c.value=v;c.font=font;c.border=BOX
    if fmt:c.number_format=fmt

CATS=['Moradia (aluguel, condomínio)','Contas (luz, água, internet, celular)','Mercado','Transporte','Saúde e farmácia','Educação','Lazer e saídas','Assinaturas e apps','Roupas e cuidados','Dívidas e parcelas','Outros']
LIM=[900,250,700,250,150,0,200,60,80,0,100]
ENTR=['Seguro-desemprego','Trabalho / freela','Rendimento da reserva','Proventos (FIIs e ações)','Outro']

# ---------- Como usar ----------
ws=sheet('Como usar',True)
ws.column_dimensions['A'].width=110
lines=[('Cripta — gastos, orçamento e plano de investimentos',TITLE),('',BASE),
('Como usar',BOLD),
('1. Aba Gastos: lance cada gasto (data, descrição, categoria, valor). A categoria tem lista pronta.',BASE),
('2. Aba Entradas: lance o que entrou (seguro-desemprego, freelas, rendimentos).',BASE),
('3. Aba Orçamento: defina o limite de cada categoria. O mês mostrado é o atual; troque a data para ver outro mês.',BASE),
('4. Aba Plano: preencha quanto você tem hoje em cada classe (reserva, FIIs, ações…) e se está empregado.',BASE),
('   A planilha calcula a meta da sua reserva, em que fase você está e onde colocar o aporte do mês.',BASE),
('',BASE),('Legenda',BOLD),
('Células amarelas com texto azul: você preenche. O resto é calculado sozinho (não apague as fórmulas).',BASE),
('Texto verde: valor que vem de outra aba.',BASE),
('A primeira linha de Gastos e de Entradas é um exemplo: apague ou troque pelos seus dados.',BASE),
('',BASE),('As fases',BOLD),
('Fase 1 — Erguer a muralha: a reserva ainda não chegou à meta. Quase tudo vai para CDB de liquidez diária e Tesouro Selic.',BASE),
('Fase 2 — Vigília: reserva completa, mas ainda sem emprego. Um pouco mais de risco e renda (FIIs e dividendos).',BASE),
('Fase 3 — Caçada: com emprego e reserva completa. A reserva fica parada na meta e todo o resto vai para os investimentos de risco.',BASE),
('',BASE),('Meta da reserva = gasto mensal × meses de reserva (12 sem emprego, 6 com emprego; você pode mudar na aba Plano).',BASE),
('As porcentagens de cada fase também podem ser mudadas na aba Plano. Isto é um plano pessoal, não recomendação oficial de investimento.',NOTE)]
for i,(t,f) in enumerate(lines,1):
    c=ws.cell(row=i,column=1,value=t);c.font=f;c.alignment=Alignment(wrap_text=True)

# ---------- Categorias ----------
cat=sheet('Listas')
head(cat,1,['Categorias de gasto','','Tipos de entrada'])
for i,c in enumerate(CATS,2):cat.cell(row=i,column=1,value=c).font=BASE
for i,c in enumerate(ENTR,2):cat.cell(row=i,column=3,value=c).font=BASE
cat.column_dimensions['A'].width=40;cat.column_dimensions['C'].width=30
cat.cell(row=14,column=1,value='Pode renomear as categorias aqui; as outras abas acompanham.').font=NOTE

# ---------- Gastos / Entradas ----------
N=1000
def registro(name,cols,lst,example):
    w=sheet(name);head(w,1,cols)
    for col,wd in zip('ABCDE',[13,38,36,15,12]):w.column_dimensions[col].width=wd
    for r in range(2,N+2):
        for col in 'ABCD':inp(w[f'{col}{r}'])
        w[f'A{r}'].number_format=DATA;w[f'D{r}'].number_format=BRL
        fx(w[f'E{r}'],f'=IF(A{r}="","",DATE(YEAR(A{r}),MONTH(A{r}),1))',MES)
    for col,v in zip('ABCD',example):w[f'{col}2'].value=v
    dv=DataValidation(type='list',formula1=lst,allow_blank=True);w.add_data_validation(dv);dv.add(f'C2:C{N+1}')
    w.freeze_panes='A2'
    w['E1'].comment=Comment('Calculado: mês do lançamento, usado no Orçamento.','Cripta')
    return w
today=datetime.date.today()
registro('Gastos',['Data','Descrição','Categoria','Valor','Mês'],'=Listas!$A$2:$A$12',[datetime.date(today.year,today.month,5),'Exemplo: aluguel',CATS[0],900])
registro('Entradas',['Data','Descrição','Tipo','Valor','Mês'],'=Listas!$C$2:$C$6',[datetime.date(today.year,today.month,5),'Exemplo: parcela do seguro',ENTR[0],1800])

# ---------- Orçamento ----------
o=sheet('Orçamento')
for col,wd in zip('ABCDE',[40,18,18,18,12]):o.column_dimensions[col].width=wd
o['A1']='Orçamento do mês';o['A1'].font=TITLE
o['A3']='Mês (troque para ver outro)';o['A3'].font=BOLD
inp(o['B3'],'=DATE(YEAR(TODAY()),MONTH(TODAY()),1)',MES)
head(o,5,['Categoria','Limite (você define)','Gasto no mês','Sobra do limite','% usado'])
for i in range(len(CATS)):
    r=6+i
    fx(o[f'A{r}'],f'=Listas!A{2+i}',font=LINK)
    inp(o[f'B{r}'],LIM[i],BRL)
    fx(o[f'C{r}'],f"=SUMIFS(Gastos!$D$2:$D${N+1},Gastos!$C$2:$C${N+1},A{r},Gastos!$E$2:$E${N+1},$B$3)",BRL)
    fx(o[f'D{r}'],f'=B{r}-C{r}',BRL)
    fx(o[f'E{r}'],f'=IF(B{r}>0,C{r}/B{r},0)',PCT)
T=6+len(CATS)  # 17
fx(o[f'A{T}'],'Total',font=BOLD)
for col,fmt in zip('BCD',[BRL,BRL,BRL]):fx(o[f'{col}{T}'],f'=SUM({col}6:{col}{T-1})',fmt,BOLD)
fx(o[f'E{T}'],f'=IF(B{T}>0,C{T}/B{T},0)',PCT,BOLD)
o[f'A{T+2}']='Entradas do mês';fx(o[f'B{T+2}'],f"=SUMIFS(Entradas!$D$2:$D${N+1},Entradas!$E$2:$E${N+1},$B$3)",BRL)
o[f'A{T+3}']='Gastos do mês';fx(o[f'B{T+3}'],f'=C{T}',BRL)
o[f'A{T+4}']='Sobrou para investir';fx(o[f'B{T+4}'],f'=B{T+2}-B{T+3}',BRL,BOLD)
for r in (T+2,T+3,T+4):o[f'A{r}'].font=BOLD
SOBRA=f"'Orçamento'!$B${T+4}"
H=T+6 # 23
o[f'A{H-1}']='Últimos 12 meses';o[f'A{H-1}'].font=BOLD
head(o,H,['Mês','Entradas','Gastos','Sobra'])
for i in range(12):
    r=H+1+i
    fx(o[f'A{r}'],f'=EDATE($B$3,{i-11})',MES)
    fx(o[f'B{r}'],f"=SUMIFS(Entradas!$D$2:$D${N+1},Entradas!$E$2:$E${N+1},A{r})",BRL)
    fx(o[f'C{r}'],f"=SUMIFS(Gastos!$D$2:$D${N+1},Gastos!$E$2:$E${N+1},A{r})",BRL)
    fx(o[f'D{r}'],f'=B{r}-C{r}',BRL)
L=H+12  # last month row
o[f'A{L+1}']='Média dos últimos 3 meses com gastos';o[f'A{L+1}'].font=BOLD
fx(o[f'C{L+1}'],f'=IFERROR(SUMIFS(C{L-2}:C{L},C{L-2}:C{L},">0")/COUNTIF(C{L-2}:C{L},">0"),0)',BRL,BOLD)
fx(o[f'D{L+1}'],f'=IFERROR(SUMIFS(D{L-2}:D{L},C{L-2}:C{L},">0")/COUNTIF(C{L-2}:C{L},">0"),0)',BRL,BOLD)
GMED=f"'Orçamento'!$C${L+1}";SMED=f"'Orçamento'!$D${L+1}"
ch=BarChart();ch.type='col';ch.title='Entradas × gastos';ch.y_axis.title='R$';ch.height=8;ch.width=18
ch.add_data(Reference(o,min_col=2,max_col=3,min_row=H,max_row=L),titles_from_data=True)
ch.set_categories(Reference(o,min_col=1,min_row=H+1,max_row=L))
ch.roundedCorners=False;ch.legend.overlay=False;ch.title.overlay=False;ch.y_axis.title.overlay=False
ch.x_axis.delete=False;ch.y_axis.delete=False;ch.x_axis.number_format=MES
o.add_chart(ch,'G3')

# ---------- Plano ----------
p=sheet('Plano')
for col,wd in zip('ABCDEF',[46,20,20,20,20,20]):p.column_dimensions[col].width=wd
p['A1']='Plano: reserva, fases e onde investir';p['A1'].font=TITLE
head(p,3,['Suas informações','Valor'])
rows=[('Gasto mensal médio (automático, últimos 3 meses)',f'={GMED}',BRL,'link'),
('Gasto mensal (se quiser digitar à mão)',None,BRL,'in'),
('Gasto usado no plano','=IF(B5>0,B5,B4)',BRL,'f'),
('Estou empregado? (Sim / Não)','Não',None,'in'),
('Meses de reserva sem emprego',12,'0','in'),
('Meses de reserva com emprego',6,'0','in'),
('Aporte deste mês (vazio = o que sobrou no Orçamento)',None,BRL,'in'),
('Aporte usado',f'=IF(B10>0,B10,MAX(0,{SOBRA}))',BRL,'f')]
for i,(lab,v,fmt,k) in enumerate(rows):
    r=4+i;p[f'A{r}']=lab;p[f'A{r}'].font=BASE;p[f'A{r}'].border=BOX
    if k=='in':inp(p[f'B{r}'],v,fmt)
    else:fx(p[f'B{r}'],v,fmt,LINK if k=='link' else BASE)
dv=DataValidation(type='list',formula1='"Sim,Não"',allow_blank=False);p.add_data_validation(dv);dv.add('B7')
p['B5'].comment=Comment('Deixe vazio para usar a média automática da aba Orçamento.','Cripta')

head(p,13,['A muralha (reserva de emergência)','Valor'])
CL=['Reserva (CDB liquidez diária + Tesouro Selic)','FIIs (renda mensal)','Ações de dividendos','Ações de crescimento','Exterior e BDRs','Cripto']
R0=36 # first class row in aporte table
rows2=[('Meta da reserva','=B6*IF(B7="Sim",B9,B8)',BRL),
('Reserva atual',f'=B{R0}',BRL),
('Falta para a meta','=MAX(0,B14-B15)',BRL),
('Progresso','=IF(B14>0,MIN(1,B15/B14),0)','0%'),
('Meses até completar (guardando a sobra média)',f'=IF(B16=0,"completa!",IF({SMED}<=0,"sem sobra para guardar",ROUNDUP(B16/{SMED},0)))','0'),
('Fase atual','=IF(B15<B14,1,IF(B7="Sim",3,2))','0'),
('Nome da fase','=INDEX(B24:D24,1,B19)',None),
('O que fazer agora','=INDEX(B25:D25,1,B19)',None)]
for i,(lab,v,fmt) in enumerate(rows2):
    r=14+i;p[f'A{r}']=lab;p[f'A{r}'].font=BASE;p[f'A{r}'].border=BOX;fx(p[f'B{r}'],v,fmt,BOLD if r in (19,20) else BASE)
p['B21'].alignment=Alignment(wrap_text=True,vertical='top');p.merge_cells('B21:F21');p.row_dimensions[21].height=48
p['B20'].alignment=Alignment(horizontal='left')

head(p,23,['Balança por fase','Fase 1','Fase 2','Fase 3','Exemplos'])
p['A24']='Nome';p['A25']='O que fazer'
for c,v in zip('BCD',['Erguer a muralha','Vigília','Caçada']):fx(p[f'{c}24'],v,font=BOLD)
for c,v in zip('BCD',['Guarde quase tudo na reserva até completar a meta. Risco pequeno, só com o que sobrar.',
                      'Reserva completa. Aumente aos poucos a renda (FIIs e dividendos), sem mexer na reserva.',
                      'Com salário e reserva completa: a reserva fica parada na meta e todo o resto vai para os investimentos.']):
    fx(p[f'{c}25'],v);p[f'{c}25'].alignment=Alignment(wrap_text=True,vertical='top')
p.row_dimensions[25].height=62
PCTS=[(0.85,0.70,0),(0.06,0.12,0.30),(0.04,0.08,0.25),(0.02,0.04,0.20),(0.00,0.02,0.15),(0.03,0.04,0.10)]
EX=['CDB 100% CDI liquidez diária, Tesouro Selic','KNCR11, KNIP11, BTLG11, XPML11, HGRU11','PINE3, ALOS3, FIQE3, PSSA3, VBBR3','BPAC11, EMBJ3, PRIO3, CURY3, WEGE3','TSMC34, IVVB11, ações dos EUA','BITH11 / ETHE11 ou BTC e ETH']
for i,(cl,pc,ex) in enumerate(zip(CL,PCTS,EX)):
    r=26+i;p[f'A{r}']=cl;p[f'A{r}'].font=BASE;p[f'A{r}'].border=BOX
    for c,v in zip('BCD',pc):inp(p[f'{c}{r}'],v,'0%')
    fx(p[f'E{r}'],ex);p[f'E{r}'].font=Font(name=F,size=9)
fx(p['A32'],'Total (precisa dar 100%)',font=BOLD)
for c in 'BCD':fx(p[f'{c}32'],f'=SUM({c}26:{c}31)','0%',BOLD)
p['A33'].value='Na Fase 3 a reserva fica fixa na meta; as porcentagens da Fase 3 dividem só o que passar dela.';p['A33'].font=NOTE

head(p,R0-1,['Onde colocar o aporte','Quanto tenho hoje (você preenche)','% hoje','Alvo após o aporte','Falta para o alvo','Colocar agora'])
TOT=R0+len(CL)
for i,cl in enumerate(CL):
    r=R0+i;fx(p[f'A{r}'],f'=A{26+i}',font=LINK)
    inp(p[f'B{r}'],[8500,600,400,200,0,300][i],BRL)
    fx(p[f'C{r}'],f'=IF($B${TOT}>0,B{r}/$B${TOT},0)',PCT)
    if i==0: alvo=f'=IF($B$19=3,$B$14,($B${TOT}+$B$11)*INDEX(B{26+i}:D{26+i},1,$B$19))'
    else: alvo=f'=IF($B$19=3,MAX(0,$B${TOT}+$B$11-$B$14)*D{26+i},($B${TOT}+$B$11)*INDEX(B{26+i}:D{26+i},1,$B$19))'
    fx(p[f'D{r}'],alvo,BRL)
    fx(p[f'E{r}'],f'=MAX(0,D{r}-B{r})',BRL)
    fx(p[f'F{r}'],f'=IF($E${TOT}>0,$B$11*E{r}/$E${TOT},0)',BRL,BOLD)
fx(p[f'A{TOT}'],'Total',font=BOLD)
for c,fmt in zip('BCDEF',[BRL,PCT,BRL,BRL,BRL]):fx(p[f'{c}{TOT}'],f'=SUM({c}{R0}:{c}{TOT-1})',fmt,BOLD)
p[f'A{TOT+1}'].value='O aporte vai para as classes que estão mais abaixo do alvo. Os valores de exemplo (R$ 10 mil) são só para mostrar: troque pelos seus.';p[f'A{TOT+1}'].font=NOTE
p[f'B{R0}'].comment=Comment('Exemplo. Coloque quanto você tem hoje na reserva (CDB + Tesouro Selic).','Cripta')
p.freeze_panes='A2'
wb.move_sheet('Plano',offset=-4)
wb.save(OUT);print('salvo',OUT)
