import os
HERE=os.path.dirname(os.path.abspath(__file__))
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

F='Arial'
def font(**k): return Font(name=F,size=k.pop('size',10),**k)
BLUE=font(color='0000FF'); BLK=font(); BOLD=font(bold=True); GREEN=font(color='008000')
TITLE=font(bold=True,size=14); HDRF=font(bold=True,color='FFFFFF')
HDR=PatternFill('solid',fgColor='1F3864'); YEL=PatternFill('solid',fgColor='FFFF00')
GREY=PatternFill('solid',fgColor='F2F2F2'); SEC=PatternFill('solid',fgColor='D9E1F2')
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
WRAP=Alignment(wrap_text=True,vertical='top')

def hdr(ws,row,labels,col=1):
    for i,l in enumerate(labels):
        c=ws.cell(row=row,column=col+i,value=l); c.font=HDRF; c.fill=HDR
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BOX
def sec(ws,row,text,ncols):
    for i in range(1,ncols+1):
        ws.cell(row=row,column=i).fill=SEC
    c=ws.cell(row=row,column=1,value=text); c.font=BOLD
def put(ws,ref,val,font_=BLK,fmt=None,fill=None,border=True):
    c=ws[ref]; c.value=val; c.font=font_
    if fmt: c.number_format=fmt
    if fill: c.fill=fill
    if border: c.border=BOX
    return c

rows=list(csv.DictReader(open(os.path.join(HERE,'..','data','cheers_monday_stats_2026-11-24.csv'))))
roster={"Granny's Bananas":['Shattuck, L','Lindholm, H'],'Diverted Flights':['Bieber, T','Bieber, J'],'180 Proof':['Zeitz, K','Dirks, J'],
'Dart Chocolate':['Jallah, K','Nelson, P'],'Total Bulls Hit':['Meyer, M','Lapthorne, A'],'One Hit Wonders':['Baluja, L','Shaw, J'],
'2 Nips 1 Tip':['Cordrey, S','Rensch, S'],'Savages Vol 2':['Larson, D','Guerrero, B'],'Darts & Recreation':['Shaw, L','Chavez, J'],
'Love at First Flight':['Mejeur, S','Mejeur, M'],'The Dartly Couple':['Rosenbach, A','Rosenbach, K']}
subs=['Cochran, S']
teams={p:t for t,ps in roster.items() for p in ps}
for _p in subs: teams[_p]='Sub'
legs={'Shattuck, L':18,'Lindholm, H':9,'Cochran, S':9,'Bieber, T':18,'Bieber, J':18,'Dirks, J':18,'Zeitz, K':18,'Jallah, K':18,'Nelson, P':18,
'Meyer, M':18,'Lapthorne, A':18,'Baluja, L':18,'Shaw, J':18,'Cordrey, S':18,'Rensch, S':18,'Larson, D':9,'Guerrero, B':9,'Shaw, L':18,
'Chavez, J':18,'Mejeur, S':18,'Mejeur, M':18,'Rosenbach, A':9,'Rosenbach, K':9}

wb=Workbook()
# ---------------- README
ws=wb.active; ws.title='How It Works'
ws['A1']='How It Works – Cheers Monday Handicap & Ranking'; ws['A1'].font=TITLE
lines=[
('START HERE',''),
('Bar Card','Pick the HOME and AWAY team (yellow cells). The card shows the 501 doubles start, cricket marks and the four singles starts. Each team\'s captain is Player 1 and the other regular is Player 2 by default. Only fill in the lineup cells if the captains set a different order, or a sub plays.'),
('Lineup and subs','Player 1 / Player 2 are chosen by the captains each night and decide the singles matchups: game 4 = Home P1 v Away P2, game 6 = Home P2 v Away P1, game 8 = P1 v P1, game 10 = P2 v P2. Each team has two regulars. A sub is not on a roster and can sub for any team; regulars cannot sub for other teams. The absent regular simply does not play, so they have no stats that week. Throwing order: the home team has the first ‘option’; the loser of the prior leg has the ‘option’ thereafter. ‘Option’ = the choice of who corks (throws one dart at the bull) first; closest to the bull throws first; on a tie, cork again in the reverse order. Throwing order does not change handicaps.'),
('Hover / tap for answers','Cells with a small marker have a note explaining exactly how the number is calculated and what was assumed. Hover (computer) or tap the cell (phone) to read it.'),
('THE THREE HANDICAPS',''),
('1. Team 501 start','Add the two players\' 3DA, divide by 2 = team average. Difference between the two team averages is looked up in the Pre-Match worksheet chart (Settings): 0–2.9 = no spot, 3.0–5.9 = 25 points, then +25 per 3 points of difference, capped at 175. The LOWER team starts at 501 minus the spot (or the higher team starts at 501 plus the spot).'),
('2. Team cricket marks','Add the two players\' MPR (not averaged) for each team. Difference between the two combined MPRs is looked up in the Pre-Match worksheet chart: 0–0.49 = 0 marks, 0.50–0.99 = 1, 1.00–1.49 = 2, 1.50–1.99 = 3, 2.00–2.49 = 4, 2.50+ = 5 (cap). The lower team receives the marks. They count toward closing numbers only; no points are scored with them.'),
('3. Singles 501 start','For each singles game: lower player\'s 3DA ÷ higher player\'s 3DA = a percentage. The lower player starts at 501 × that percentage (rounded); the higher player starts at 501. ASSUMPTION: the Match Night worksheet gives the percentage but not how to convert it to a start score.'),
('Game 11','701 doubles, scratch. No handicap.'),
('NEW PLAYERS AND SUBS',''),
('Rule used','A player with any league stats uses their own 3DA and MPR. A new player or sub with no league stats must have their 3DA and MPR entered before handicaps are shown: type them in the yellow "Entered 3DA / Entered MPR" cells on the Handicap tab (C7:D10), or on the Players tab (columns J and K) so they are kept for later weeks. The numbers can come from another league or the player\'s DartConnect profile. Until both are entered, the Bar Card says so instead of showing handicaps.'),
('Same as the phone card','The phone Bar Card (kzeitz.github.io/cheers-bar-card) uses exactly this rule. On the phone, pick "New player…" and enter the name, 3DA and MPR.'),
('PLAYER RANKING AND PLAYER OF THE WEEK',''),
('Cheers Rating','100 × [w × (3DA ÷ league 3DA) + (1 − w) × (MPR ÷ league MPR)]. 100 = league average. w = 50% (Settings). Uses all of a player\'s darts: singles and doubles 01 legs for 3DA, doubles cricket for MPR. Handicap marks and spots are not thrown darts and are not counted.'),
('Provisional','Players under the minimum legs for a full ranking (Settings B11, 18) are flagged Provisional.'),
('Player of the Week','Week score = 100 × [w × (week 3DA ÷ baseline 3DA) + (1 − w) × (week MPR ÷ baseline MPR)], then pulled toward 100 by darts ÷ (darts + K) so a lucky short sample does not win. Baseline is each player\'s own season average, so a 30-average and a 55-average player compete fairly. Needs the minimum darts (Settings B9).'),
('DATA AND UPDATES',''),
('Source','DartConnect League Leaderboard (3DA = All \'01, MPR = Doubles Cricket), DCTV standings (rosters, legs played, league averages 39.90 / 1.35). Pulled 24 Nov 2026 for Winter/Spring 2026/27, 23 players, 2 weeks played.'),
('Updating','Not automatic. After each night paste the new leaderboard numbers into the Players tab (blue = typed-in data) and update "Stats as of" on Settings. Snapshot baselines before refreshing.'),
('Colour key','Yellow = you fill in. Blue text = typed-in data or assumption. Black = formula.'),
('OPEN ASSUMPTIONS TO CONFIRM',''),
('1','Doubles handicaps use the Pre-Match worksheet charts. The Match Night sheet\'s "lower ÷ higher = %" layout for doubles is shown on the Handicap tab for comparison only.'),
('2','Singles start-score conversion (above).'),
('3','Rating weights 50/50, shrink constant K = 60 darts, minimum darts 60, and minimum legs 18 are placeholders (Settings).'),
('4','Where the two worksheets came from is unknown; they do not appear to be DartConnect documents. Confirm the league\'s official rules before relying on them.'),
]
r=3
for a,b in lines:
    if b=='' and a:
        ws.cell(row=r,column=1,value=a).font=BOLD; ws.cell(row=r,column=1).fill=SEC; ws.cell(row=r,column=2).fill=SEC
    else:
        ws.cell(row=r,column=1,value=a).font=BOLD
        c=ws.cell(row=r,column=2,value=b); c.font=BLK; c.alignment=WRAP
    r+=1
ws.column_dimensions['A'].width=26; ws.column_dimensions['B'].width=120

# ---------------- Settings
st=wb.create_sheet('Settings')
st['A1']='Settings & Lookup Charts'; st['A1'].font=TITLE
hdr(st,3,['Setting','Value','Note'])
S=[('League average 3DA (\'01)',39.90,'0.00','DCTV Division Average, standings page, 24 Nov 2026'),
('League average MPR (Cricket)',1.35,'0.00','DCTV Division Average, standings page, 24 Nov 2026'),
('Weight on 3DA',0.5,'0%','Assumption – adjust'),
('Weight on MPR','=1-B6','0%','1 minus weight on 3DA'),
('Shrink constant K (darts)',60,'0','Assumption – darts of evidence needed to give a week half weight'),
('Minimum darts for Player of the Week',60,'0','Assumption'),
('Standard 501 start',501,'0','Worksheet'),
('Minimum legs for a full ranking',18,'0','Assumption – players below this are flagged Provisional (18 = two full nights)')]
for i,(a,b,f,n) in enumerate(S):
    rr=4+i
    put(st,f'A{rr}',a); 
    put(st,f'B{rr}',b,BLUE if not str(b).startswith('=') else BLK,f)
    put(st,f'C{rr}',n,font(italic=True,color='595959'))
st['B6'].fill=YEL; st['B11'].fill=YEL
# rows: 4 L3DA,5 LMPR,6 w3DA,7 wMPR,8 K,9 minDarts,10 start501
hdr(st,3,['Data','Value','Note'],7)
put(st,'G9','Stats as of'); put(st,'H9','24 Nov 2026',BLUE,fill=YEL); put(st,'I9','Date the Players data was pulled from DartConnect. Update when you refresh.',font(italic=True,color='595959'))
for c,w in zip('GHI',(30,14,70)): st.column_dimensions[c].width=w
sec(st,12,'Pre-Match worksheet charts (source: your "Pre-Match Handicap Worksheet (Doubles 501 & Cricket)")',6)
hdr(st,13,['3DA diff (from)','Spot (pts)','','Combined MPR diff (from)','Starting marks'],1)
spots=[(0,0),(3,25),(6,50),(9,75),(12,100),(15,125),(18,150),(21,175)]
for i,(b,s) in enumerate(spots):
    put(st,f'A{14+i}',b,BLUE,'0.0'); put(st,f'B{14+i}',s,BLUE,'0')
marks=[(0,0),(0.5,1),(1,2),(1.5,3),(2,4),(2.5,5)]
for i,(b,m) in enumerate(marks):
    put(st,f'D{14+i}',b,BLUE,'0.00'); put(st,f'E{14+i}',m,BLUE,'0')
st['A23']='Bands: 0.0–2.9, 3.0–5.9 … 21.0+ (cap 175). Marks: 0.00–0.49 … 2.50+ (cap 5). Lookups use the lower bound of each band.'
st['A23'].font=font(italic=True,color='595959')
for c,w in zip('ABCDE',(38,14,62,26,16)): st.column_dimensions[c].width=w

# ---------------- Players
pl=wb.create_sheet('Players')
pl['A1']='Players – Season Stats & Cheers Rating'; pl['A1'].font=TITLE
pl['A2']='Entered 3DA/MPR/source (yellow, optional): for a new player or sub with no league stats yet, e.g. from another league or their DartConnect profile. Used only while they have no league stats. Source: DartConnect League Leaderboard (stats) and DCTV standings (legs played), pulled 24 Nov 2026. Add new players in the blank rows (up to row 40). Rating = 100 × [w × 3DA ÷ league 3DA + (1−w) × MPR ÷ league MPR].'
pl['A2'].font=font(italic=True,color='595959')
hdr(pl,4,['Player','Team',"3DA ('01)",'MPR (Cricket)','Cheers Rating','Rank','SortKey','Legs played','Status','Entered 3DA','Entered MPR','Entered source'])
dv3p=DataValidation(type='decimal',operator='between',formula1='10',formula2='100',allow_blank=True,showErrorMessage=True,errorTitle='3DA out of range',error='Enter a 3DA between 10 and 100, for example 38.5.')
dvmp=DataValidation(type='decimal',operator='between',formula1='0.1',formula2='6',allow_blank=True,showErrorMessage=True,errorTitle='MPR out of range',error='Enter an MPR between 0.10 and 6.00, for example 1.40.')
pl.add_data_validation(dv3p); pl.add_data_validation(dvmp); dv3p.add('J5:J40'); dvmp.add('K5:K40')
for i in range(36):
    rr=5+i
    if i<len(rows):
        r_=rows[i]
        put(pl,f'A{rr}',r_['Player'],BLUE); put(pl,f'B{rr}',teams.get(r_['Player'],''),BLUE,fill=YEL if r_['Player'] not in teams else None)
        put(pl,f'C{rr}',float(r_['3DA_01']),BLUE,'0.00'); put(pl,f'D{rr}',float(r_['MPR_Cricket']),BLUE,'0.00')
    else:
        for col in 'ABCD': put(pl,f'{col}{rr}',None,BLUE,fill=YEL)
        pl[f'C{rr}'].number_format='0.00'; pl[f'D{rr}'].number_format='0.00'
    put(pl,f'E{rr}',f'=IF(OR(C{rr}="",D{rr}=""),"",100*(Settings!$B$6*C{rr}/Settings!$B$4+Settings!$B$7*D{rr}/Settings!$B$5))',BLK,'0.0')
    put(pl,f'F{rr}',f'=IF(E{rr}="","",RANK(E{rr},$E$5:$E$40))',BLK,'0')
    put(pl,f'G{rr}',f'=IF(E{rr}="","",E{rr}+ROW()*0.000000001)',BLK,'0.000000')
    pn=rows[i]['Player'] if i<len(rows) else None
    put(pl,f'H{rr}',legs.get(pn),BLUE,'0',YEL if pn is None else None)
    put(pl,f'I{rr}',f'=IF(OR(E{rr}="",H{rr}=""),"",IF(H{rr}<Settings!$B$11,"Provisional","OK"))')
    put(pl,f'J{rr}',None,BLUE,'0.00',YEL); put(pl,f'K{rr}',None,BLUE,'0.00',YEL); put(pl,f'L{rr}',None,BLUE,None,YEL)
pl.column_dimensions['G'].hidden=True
for c,w in zip('ABCDEFHIJKL',(22,18,12,14,14,8,12,13,12,12,34)): pl.column_dimensions[c].width=w
pl.freeze_panes='A5'

# ---------------- Season Ranking
sr=wb.create_sheet('Season Ranking')
sr['A1']='Season Player Ranking – Cheers Rating'; sr['A1'].font=TITLE
sr['A2']='100 = league average. Sorted automatically from the Players sheet. Provisional = fewer legs than the minimum on Settings (stats still unsettled).'; sr['A2'].font=font(italic=True,color='595959')
hdr(sr,4,['Rank','Player','Team',"3DA ('01)",'MPR (Cricket)','Cheers Rating','idx','Legs played','Status','Entered 3DA','Entered MPR','Entered source'])
for i in range(36):
    rr=5+i
    put(sr,f'A{rr}',f'=IF(ROWS($A$5:A{rr})>COUNT(Players!$G$5:$G$40),"",ROWS($A$5:A{rr}))',BLK,'0')
    put(sr,f'G{rr}',f'=IF(A{rr}="","",MATCH(LARGE(Players!$G$5:$G$40,A{rr}),Players!$G$5:$G$40,0))',BLK,'0')
    put(sr,f'B{rr}',f'=IF(A{rr}="","",INDEX(Players!$A$5:$A$40,G{rr}))')
    put(sr,f'C{rr}',f'=IF(A{rr}="","",INDEX(Players!$B$5:$B$40,G{rr})&"")')
    put(sr,f'D{rr}',f'=IF(A{rr}="","",INDEX(Players!$C$5:$C$40,G{rr}))',BLK,'0.00')
    put(sr,f'E{rr}',f'=IF(A{rr}="","",INDEX(Players!$D$5:$D$40,G{rr}))',BLK,'0.00')
    put(sr,f'F{rr}',f'=IF(A{rr}="","",INDEX(Players!$E$5:$E$40,G{rr}))',BLK,'0.0')
for rr in range(5,41):
    put(sr,f'H{rr}',f'=IF(A{rr}="","",INDEX(Players!$H$5:$H$40,G{rr}))',BLK,'0')
    put(sr,f'I{rr}',f'=IF(A{rr}="","",INDEX(Players!$I$5:$I$40,G{rr}))')
sr.column_dimensions['G'].hidden=True
sr.column_dimensions['H'].width=12; sr.column_dimensions['I'].width=13
for c,w in zip('ABCDEF',(8,22,18,12,14,14)): sr.column_dimensions[c].width=w
sr.freeze_panes='A5'

# ---------------- Handicap
hc=wb.create_sheet('Handicap',1)
hc['A1']='Match Handicap Calculator'; hc['A1'].font=TITLE
put(hc,'A3','Date',BOLD); put(hc,'B3','2026-11-30',BLUE,fill=YEL)
put(hc,'C3','Round / Week',BOLD); put(hc,'D3',3,BLUE,fill=YEL)
put(hc,'A4','Home team',BOLD); put(hc,'B4',"='Bar Card'!B4",GREEN)
put(hc,'C4','Away team',BOLD); put(hc,'D4',"='Bar Card'!B5",GREEN)
put(hc,'E3','Roster',BOLD,border=False); put(hc,'F3','=IFERROR(INDEX(Teams!$D$5:$D$20,MATCH(B4,Teams!$A$5:$A$20,0)),"(team not in Teams sheet)")',BLK,border=False)
put(hc,'E4','Roster',BOLD,border=False); put(hc,'F4','=IFERROR(INDEX(Teams!$D$5:$D$20,MATCH(D4,Teams!$A$5:$A$20,0)),"(team not in Teams sheet)")',BLK,border=False)
hdr(hc,6,['Slot','Player (from Bar Card)','Entered 3DA (new player)','Entered MPR (new player)','3DA used','MPR used','Basis','Legs played','Players-tab 3DA','Players-tab MPR'])
slots=[('Home Player 1','Zeitz, K'),('Home Player 2','Dirks, J'),('Away Player 1','Cordrey, S'),('Away Player 2','Rensch, S')]
put(hc,'L2','Captain (P1 default)',BOLD,border=False); put(hc,'M2','Other regular',BOLD,border=False)
put(hc,'L3','=IFERROR(INDEX(Teams!$B$5:$B$20,MATCH($B$4,Teams!$A$5:$A$20,0)),"")',GREEN); put(hc,'M3','=IFERROR(INDEX(Teams!$C$5:$C$20,MATCH($B$4,Teams!$A$5:$A$20,0)),"")',GREEN)
put(hc,'L4','=IFERROR(INDEX(Teams!$B$5:$B$20,MATCH($D$4,Teams!$A$5:$A$20,0)),"")',GREEN); put(hc,'M4','=IFERROR(INDEX(Teams!$C$5:$C$20,MATCH($D$4,Teams!$A$5:$A$20,0)),"")',GREEN)
dv=DataValidation(type='list',formula1='=Players!$A$5:$A$40',allow_blank=True,showErrorMessage=False); hc.add_data_validation(dv)
for i,(s,p) in enumerate(slots):
    rr=7+i
    L,M=('$L$3','$M$3') if i<2 else ('$L$4','$M$4')
    if i==0:   f_="=IF('Bar Card'!B9<>\"\",'Bar Card'!B9,IF('Bar Card'!B10="+L+","+M+","+L+"))"
    elif i==1: f_="=IF('Bar Card'!B10<>\"\",'Bar Card'!B10,IF(B7="+L+","+M+",IF(B7="+M+","+L+","+M+")))"
    elif i==2: f_="=IF('Bar Card'!B11<>\"\",'Bar Card'!B11,IF('Bar Card'!B12="+L+","+M+","+L+"))"
    else:      f_="=IF('Bar Card'!B12<>\"\",'Bar Card'!B12,IF(B9="+L+","+M+",IF(B9="+M+","+L+","+M+")))"
    put(hc,f'A{rr}',s_ if False else slots[i][0],BOLD); put(hc,f'B{rr}',f_,GREEN)
    put(hc,f'C{rr}',None,BLUE,'0.00',YEL); put(hc,f'D{rr}',None,BLUE,'0.00',YEL)
    put(hc,f'H{rr}',f'=IFERROR(INDEX(Players!$H$5:$H$40,MATCH(B{rr},Players!$A$5:$A$40,0)),0)',BLK,'0')
    put(hc,f'I{rr}',f'=IFERROR(IF(ISNUMBER(INDEX(Players!$J$5:$J$40,MATCH(B{rr},Players!$A$5:$A$40,0))),INDEX(Players!$J$5:$J$40,MATCH(B{rr},Players!$A$5:$A$40,0)),""),"")',BLK,'0.00')
    put(hc,f'J{rr}',f'=IFERROR(IF(ISNUMBER(INDEX(Players!$K$5:$K$40,MATCH(B{rr},Players!$A$5:$A$40,0))),INDEX(Players!$K$5:$K$40,MATCH(B{rr},Players!$A$5:$A$40,0)),""),"")',BLK,'0.00')
    put(hc,f'G{rr}',f'=IF(AND(ISNUMBER(C{rr}),ISNUMBER(D{rr})),"Entered",IF(AND(H{rr}>0,ISNUMBER(IFERROR(INDEX(Players!$C$5:$C$40,MATCH(B{rr},Players!$A$5:$A$40,0)),""))),"Own stats",IF(AND(ISNUMBER(I{rr}),ISNUMBER(J{rr})),"Entered (Players tab)","Needs 3DA & MPR")))')
    put(hc,f'E{rr}',f'=IF(G{rr}="Entered",C{rr},IF(G{rr}="Own stats",INDEX(Players!$C$5:$C$40,MATCH(B{rr},Players!$A$5:$A$40,0)),IF(G{rr}="Entered (Players tab)",I{rr},"")))',BLK,'0.00')
    put(hc,f'F{rr}',f'=IF(G{rr}="Entered",D{rr},IF(G{rr}="Own stats",INDEX(Players!$D$5:$D$40,MATCH(B{rr},Players!$A$5:$A$40,0)),IF(G{rr}="Entered (Players tab)",J{rr},"")))',BLK,'0.00')
dv3=DataValidation(type='decimal',operator='between',formula1='10',formula2='100',allow_blank=True,showErrorMessage=True,errorTitle='3DA out of range',error='Enter a 3DA between 10 and 100, for example 38.5.')
dvm=DataValidation(type='decimal',operator='between',formula1='0.1',formula2='6',allow_blank=True,showErrorMessage=True,errorTitle='MPR out of range',error='Enter an MPR between 0.10 and 6.00, for example 1.40.')
hc.add_data_validation(dv3); hc.add_data_validation(dvm); dv3.add('C7:C10'); dvm.add('D7:D10')
hc['A11']='Lineup comes from the Bar Card tab: blank = captain is Player 1, other regular is Player 2. Players with league stats use them. A new player with no league stats needs their 3DA and MPR typed in the yellow Entered cells (or on the Players tab); until then no handicaps are shown.'
hc['A11'].font=font(italic=True,color='595959')

sec(hc,13,'PART 1 – 501 DOUBLES (games 3 and 7): team average 3DA difference',6)
P1=[('Home combined 3DA (A+B)','=E7+E8','0.00'),('Away combined 3DA (A+B)','=E9+E10','0.00'),
('Home team average 3DA (÷ 2)','=B14/2','0.00'),('Away team average 3DA (÷ 2)','=B15/2','0.00'),
('Higher team average','=MAX(B16,B17)','0.00'),('Lower team average','=MIN(B16,B17)','0.00'),
('Team average 3DA difference','=B18-B19','0.00'),
('Spot (points)','=INDEX(Settings!$B$14:$B$21,MATCH(B20,Settings!$A$14:$A$21,1))','0'),
('Team receiving spot','=IF(B21=0,"None – straight 501",IF(B16<B17,B4,D4))',None),
('Lower team starts at','=Settings!$B$10-B21','0'),
('(Alt) higher team starts at','=Settings!$B$10+B21','0')]
for i,(a,f,nf) in enumerate(P1):
    rr=14+i; put(hc,f'A{rr}',a,BOLD if 'Spot' in a or 'receiving' in a else BLK); put(hc,f'B{rr}','=IFERROR('+f[1:]+',"")',BLK,nf)

sec(hc,26,'PART 2 – CRICKET DOUBLES (games 1, 2, 5 and 9): combined MPR difference',6)
P2=[('Home combined MPR (A+B)','=F7+F8','0.00'),('Away combined MPR (A+B)','=F9+F10','0.00'),
('Higher team combined MPR','=MAX(B27,B28)','0.00'),('Lower team combined MPR','=MIN(B27,B28)','0.00'),
('Combined MPR difference','=B29-B30','0.00'),
('Starting marks awarded','=INDEX(Settings!$E$14:$E$19,MATCH(B31,Settings!$D$14:$D$19,1))','0'),
('Team receiving marks','=IF(B32=0,"None – straight up",IF(B27<B28,B4,D4))',None)]
for i,(a,f,nf) in enumerate(P2):
    rr=27+i; put(hc,f'A{rr}',a,BOLD if 'Starting' in a or 'receiving' in a else BLK); put(hc,f'B{rr}','=IFERROR('+f[1:]+',"")',BLK,nf)
hc['A34']='Handicap marks count toward closing numbers only; no points are awarded before Dart 1 (per worksheet).'
hc['A34'].font=font(italic=True,color='595959')

sec(hc,36,'PART 3 – 501 SINGLES (games 4, 6, 8, 10): percentage handicapping. Game 11 = 701 doubles, scratch (no handicap).',9)
hdr(hc,37,['Game','Matchup','Home player','Home 3DA','Away player','Away 3DA','Lower ÷ Higher','Lower player (receives)','Lower player starts at (higher = 501)'])
sg=[(4,'H1 v A2','B7','E7','B10','E10'),(6,'H2 v A1',None,None,None,None),(8,'H1 v A1',None,None,None,None),(10,'H2 v A2',None,None,None,None)]
m={4:('B7','E7','B10','E10'),6:('B8','E8','B9','E9'),8:('B7','E7','B9','E9'),10:('B8','E8','B10','E10')}
bs={4:('G7','G10'),6:('G8','G9'),8:('G7','G9'),10:('G8','G10')}
# note: slot rows: H1=7, H2=8, A1=9, A2=10
for i,(g,mu) in enumerate([(4,'H1 v A2'),(6,'H2 v A1'),(8,'H1 v A1'),(10,'H2 v A2')]):
    rr=38+i; hp,h3,ap,a3=m[g]
    put(hc,f'A{rr}',g,BLK,'0'); put(hc,f'B{rr}',mu)
    put(hc,f'C{rr}',f'={hp}'); put(hc,f'D{rr}',f'={h3}',BLK,'0.00')
    put(hc,f'E{rr}',f'={ap}'); put(hc,f'F{rr}',f'={a3}',BLK,'0.00')
    b1,b2=bs[g]
    put(hc,f'G{rr}',f'=IFERROR(MIN(D{rr},F{rr})/MAX(D{rr},F{rr}),"")',BLK,'0.0%')
    put(hc,f'H{rr}',f'=IF(G{rr}="","",IF(G{rr}=1,"Even (both 501)",IF(D{rr}<F{rr},C{rr},E{rr})))')
    put(hc,f'I{rr}',f'=IF(G{rr}="","",ROUND(Settings!$B$10*G{rr},0))',BLK,'0')
hc['A42']='Assumption: lower player starts at 501 × percentage (rounded); the worksheet gives the percentage but not the conversion.'
hc['A42'].font=font(italic=True,color='595959')

sec(hc,44,'FOR COMPARISON ONLY – "lower ÷ higher = %" layout from the Match Night sheet, Step 2 (conversion to spot/marks not defined)',6)
put(hc,'A45','501 doubles: lower ÷ higher team avg 3DA'); put(hc,'B45','=IFERROR(B19/B18,"")',BLK,'0.0%')
put(hc,'A46','Cricket doubles: lower ÷ higher combined MPR'); put(hc,'B46','=IFERROR(B30/B29,"")',BLK,'0.0%')
for c,w in zip('ABCDEFGHIJ',(44,24,22,18,16,12,16,24,26,12)): hc.column_dimensions[c].width=w
for rr in (37,): hc.row_dimensions[rr].height=42

# ---------------- Teams
tm=wb.create_sheet('Teams')
tm['A1']='Team Rosters and Subs'; tm['A1'].font=TITLE
tm['A2']='Each team has two regulars. The captain is Player 1 by default. Source: DartConnect schedule (captains) and DCTV standings (players), 24 Nov 2026. Names match the Players sheet.'
tm['A2'].font=font(italic=True,color='595959')
hdr(tm,4,['Team','Captain (Player 1 default)','Other regular','Roster'])
for i,(t,ps) in enumerate(roster.items()):
    rr=5+i
    put(tm,f'A{rr}',t,BLUE); put(tm,f'B{rr}',ps[0],BLUE); put(tm,f'C{rr}',ps[1],BLUE)
    put(tm,f'D{rr}',f'=B{rr}&IF(C{rr}<>"","; "&C{rr},"")')
for rr in range(5+len(roster),21):
    for col in 'ABC': put(tm,f'{col}{rr}',None,BLUE,fill=YEL)
    put(tm,f'D{rr}',f'=IF(A{rr}="","",B{rr}&IF(C{rr}<>"","; "&C{rr},""))')
hdr(tm,4,['Subs (not on any roster)'],6)
for i in range(20):
    put(tm,f'F{5+i}',subs[i] if i<len(subs) else None,BLUE,fill=YEL if i>=len(subs) else None)
tm['F26']='A sub is not on a team roster and can sub for any team. Regulars cannot sub for another team.'
tm['F26'].font=font(italic=True,color='595959')
for c,w in zip('ABCDEF',(24,26,18,40,3,28)): tm.column_dimensions[c].width=w

# ---------------- Weekly POTW
wk=wb.create_sheet('Weekly POTW')
wk['A1']='Player of the Week'; wk['A1'].font=TITLE
put(wk,'A2','Week of',BOLD); put(wk,'B2',None,BLUE,fill=YEL)
wk['A3']=('Score = 100 × [w × (week 3DA ÷ baseline 3DA) + (1−w) × (week MPR ÷ baseline MPR)], then pulled toward 100 by darts/(darts+K). '
          'Baseline = season stats on Players (snapshot as values before refreshing). Yellow = enter each week. Row 7 is an EXAMPLE – clear it.')
wk['A3'].font=font(italic=True,color='595959'); wk['A3'].alignment=WRAP; wk.merge_cells('A3:J3'); wk.row_dimensions[3].height=40
hdr(wk,6,['Player','Baseline 3DA','Baseline MPR','Week 3DA','Week MPR','Darts thrown','Raw score','Adjusted score','Eligible?','key'])
for i in range(36):
    rr=7+i; pr=5+i
    put(wk,f'A{rr}',f'=IF(Players!A{pr}="","",Players!A{pr})')
    put(wk,f'B{rr}',f'=IF(A{rr}="","",Players!C{pr})',BLK,'0.00')
    put(wk,f'C{rr}',f'=IF(A{rr}="","",Players!D{pr})',BLK,'0.00')
    for col,fm in (('D','0.00'),('E','0.00'),('F','0')):
        put(wk,f'{col}{rr}',None,BLUE,fm,YEL)
    put(wk,f'G{rr}',f'=IF(OR(A{rr}="",D{rr}="",E{rr}=""),"",100*(Settings!$B$6*D{rr}/B{rr}+Settings!$B$7*E{rr}/C{rr}))',BLK,'0.0')
    put(wk,f'H{rr}',f'=IF(OR(G{rr}="",F{rr}=""),"",100+(G{rr}-100)*F{rr}/(F{rr}+Settings!$B$8))',BLK,'0.0')
    put(wk,f'I{rr}',f'=IF(H{rr}="","",IF(F{rr}>=Settings!$B$9,"Yes","No"))')
    put(wk,f'J{rr}',f'=IF(I{rr}="Yes",H{rr}+ROW()*0.000000001,"")',BLK,'0.000000')
wk.column_dimensions['J'].hidden=True
# example row (Zeitz is first player)
wk['D7']=55.0; wk['E7']=2.10; wk['F7']=110
wk['K7']='EXAMPLE (illustrative values) – clear D7:F7'; wk['K7'].font=font(bold=True,color='C00000')
hdr(wk,6,['Top 3','Player','Adjusted score'],12)
for k in (1,2,3):
    rr=6+k
    put(wk,f'L{rr}',k,BLK,'0')
    put(wk,f'M{rr}',f'=IF({k}>COUNT($J$7:$J$42),"",INDEX($A$7:$A$42,MATCH(LARGE($J$7:$J$42,{k}),$J$7:$J$42,0)))')
    put(wk,f'N{rr}',f'=IF(M{rr}="","",INDEX($H$7:$H$42,MATCH(LARGE($J$7:$J$42,{k}),$J$7:$J$42,0)))',BLK,'0.0')
for c,w in zip('ABCDEFGHIKLMN',(22,12,12,11,11,12,11,14,10,40,8,22,14)): wk.column_dimensions[c].width=w
wk.freeze_panes='A7'


# ================= BAR CARD =================
bc=wb.create_sheet('Bar Card')
bc.sheet_view.showGridLines=False
def bset(ref,val,f=BLK,fill=None,fmt=None,al=None,border=True):
    c=bc[ref]; c.value=val; c.font=f
    if fill: c.fill=fill
    if border: c.border=BOX
    c.alignment=al or Alignment(wrap_text=True,vertical='center')
    return c
bc.merge_cells('A1:B1'); bset('A1','Cheers Monday – Handicap Card',font(bold=True,size=18),border=False)
bc.merge_cells('A2:B2'); bset('A2','Pick the two teams. Captains are Player 1 by default.',font(italic=True,size=11,color='595959'),border=False)
bset('A4','HOME team',font(bold=True,size=13)); bset('B4','180 Proof',font(bold=True,size=14,color='0000FF'),YEL)
bset('A5','AWAY team',font(bold=True,size=13)); bset('B5','2 Nips 1 Tip',font(bold=True,size=14,color='0000FF'),YEL)
bset('A6','Home lineup',font(size=11)); bset('B6','=Handicap!B7&" (P1) & "&Handicap!B8&" (P2)"',font(size=11))
bset('A7','Away lineup',font(size=11)); bset('B7','=Handicap!B9&" (P1) & "&Handicap!B10&" (P2)"',font(size=11))
bc.merge_cells('A8:B8'); bset('A8','Lineup tonight – optional. Blank = captain is Player 1, other regular is Player 2.',font(bold=True,size=10,color='595959'),SEC)
for k,(lab) in enumerate(['Home Player 1','Home Player 2','Away Player 1','Away Player 2']):
    bset(f'A{9+k}',lab,font(size=11)); bset(f'B{9+k}',None,font(size=12,color='0000FF'),YEL)
bc.merge_cells('A13:B13'); bset('A13','YOUR STARTS',HDRF,HDR,al=Alignment(horizontal='center',vertical='center'))
R=font(bold=True,size=14)
bset('A14','Cricket doubles\n(games 1, 2, 5, 9)',font(bold=True,size=12)); bset('B14','=IF(COUNTIF(Handicap!$G$7:$G$10,"Needs*")>0,"Enter 3DA and MPR for every new player (Handicap tab, yellow Entered cells)",IF(Handicap!B32=0,"Straight up – no free marks",Handicap!B33&" gets "&Handicap!B32&IF(Handicap!B32=1," free mark"," free marks")))',R,GREY)
bset('A15','501 doubles\n(games 3 & 7)',font(bold=True,size=12)); bset('B15','=IF(COUNTIF(Handicap!$G$7:$G$10,"Needs*")>0,"Enter 3DA and MPR for every new player (Handicap tab, yellow Entered cells)",IF(Handicap!B21=0,"Both teams start at 501",Handicap!B22&" starts at "&Handicap!B23&" (other team 501)"))',R,GREY)
for k,(r_,g) in enumerate(((38,4),(39,6),(40,8),(41,10))):
    rr=16+k
    bset(f'A{rr}',f'="Game {g}"&CHAR(10)&Handicap!C{r_}&" v "&Handicap!E{r_}',font(bold=True,size=11))
    bset(f'B{rr}',f'=IF(COUNTIF(Handicap!$G$7:$G$10,"Needs*")>0,"Enter 3DA and MPR for every new player (Handicap tab, yellow Entered cells)",IF(Handicap!G{r_}=1,"Both start at 501",Handicap!H{r_}&" starts at "&Handicap!I{r_}&" (other 501)"))',R,GREY)
bset('A20','Game 11\n701 doubles',font(bold=True,size=12)); bset('B20','Scratch – no handicap',R,GREY)
bc.merge_cells('A21:B21'); bset('A21','Home team has the first ‘option’. Loser of the prior leg has the ‘option’ thereafter. (See How It Works.)',font(italic=True,size=10,color='595959'),border=False)
bc.row_dimensions[21].height=30
bc.merge_cells('A22:B22')
bset('A22','=IF(COUNTIF(Handicap!G7:G10,"Needs*")>0,"NOTE: a new player has no league stats. Enter their 3DA and MPR on the Handicap tab (yellow Entered cells) to see handicaps.",IF(COUNTIF(Handicap!G7:G10,"Entered*")>0,"NOTE: a new player is using entered 3DA and MPR.",""))',font(bold=True,size=11,color='C00000'),border=False)
bc.merge_cells('A23:B23')
bset('A23','=IF(OR(AND(COUNTIF(Teams!$B$5:$C$20,Handicap!B7)>0,Handicap!B7<>Handicap!$L$3,Handicap!B7<>Handicap!$M$3),AND(COUNTIF(Teams!$B$5:$C$20,Handicap!B8)>0,Handicap!B8<>Handicap!$L$3,Handicap!B8<>Handicap!$M$3),AND(COUNTIF(Teams!$B$5:$C$20,Handicap!B9)>0,Handicap!B9<>Handicap!$L$4,Handicap!B9<>Handicap!$M$4),AND(COUNTIF(Teams!$B$5:$C$20,Handicap!B10)>0,Handicap!B10<>Handicap!$L$4,Handicap!B10<>Handicap!$M$4)),"NOTE: a regular from another team is in a lineup. Regulars cannot sub for other teams.","")',font(bold=True,size=11,color='C00000'),border=False)
bc.row_dimensions[23].height=36
bset('A24','Stats as of',font(size=10,color='595959')); bset('B24','=Settings!H9',font(size=10,color='595959'))
bc.merge_cells('A25:B25'); bset('A25','Questions? Hover or tap any result for how it was calculated. Full step-by-step is on the Handicap tab.',font(italic=True,size=10,color='595959'),border=False)
bc.column_dimensions['A'].width=24; bc.column_dimensions['B'].width=42
for rr in (14,15,16,17,18,19,20): bc.row_dimensions[rr].height=46
for rr in (4,5): bc.row_dimensions[rr].height=30
bc.row_dimensions[22].height=48; bc.row_dimensions[25].height=32; bc.row_dimensions[1].height=30
bc['D8']='Home list'; bc['E8']='Away list'
bc['D9']='=Handicap!L3'; bc['D10']='=Handicap!M3'; bc['E9']='=Handicap!L4'; bc['E10']='=Handicap!M4'
for _i in range(10):
    bc[f'D{11+_i}']=f'=IF(Teams!F{5+_i}="","",Teams!F{5+_i})'; bc[f'E{11+_i}']=f'=IF(Teams!F{5+_i}="","",Teams!F{5+_i})'
bc.column_dimensions['D'].hidden=True; bc.column_dimensions['E'].hidden=True
dvb=DataValidation(type='list',formula1='=Teams!$A$5:$A$20',allow_blank=False); bc.add_data_validation(dvb); dvb.add('B4'); dvb.add('B5')
dvp=DataValidation(type='list',formula1='=$D$9:$D$20',allow_blank=True,showErrorMessage=False); bc.add_data_validation(dvp); dvp.add('B9:B10')
dvq=DataValidation(type='list',formula1='=$E$9:$E$20',allow_blank=True,showErrorMessage=False); bc.add_data_validation(dvq); dvq.add('B11:B12')

# ================= NOTES (hover / tap explanations) =================
NOTES={}
def note(ws,ref,text,w=380,h=170):
    NOTES.setdefault(ws.title,{})[ref]=text
note(bc,'B4','Home team = the team listed first on the DartConnect schedule. Choose from the list. Team rosters come from the Teams tab.',300,90)
note(bc,'B9','Player 1 and Player 2 are set by the captains each night and decide the singles matchups (game 4 = Home P1 v Away P2, game 6 = P2 v P1, game 8 = P1 v P1, game 10 = P2 v P2). Leave blank = captain is Player 1 and the other regular is Player 2. Pick a different order or a sub from the list; a brand-new sub can be typed in, then enter their 3DA and MPR on the Handicap tab (yellow Entered cells). If a sub is Player 1 and Player 2 is blank, Player 2 defaults to the non-captain regular, so set Player 2 yourself if the sub replaces that regular.',380,230)
note(bc,'B14','Pre-Match worksheet, Part 2. Combined MPR = the two players\' MPR added (not averaged). The difference is looked up in the marks chart (Settings). The lower team receives the marks. Full steps: Handicap tab, Part 2.',360,150)
note(bc,'B15','Pre-Match worksheet, Part 1. Team average 3DA = the two players\' 3DA added and divided by 2. The difference between the two team averages is looked up in the spot chart (Settings). The lower team starts at 501 minus the spot. Full steps: Handicap tab, Part 1.',360,150)
note(bc,'B16','Game 4 = Home Player 1 v Away Player 2. Singles percentage handicapping. Lower 3DA ÷ higher 3DA = %. Lower player starts at 501 × % (rounded); higher player starts at 501. ASSUMPTION: the worksheet does not say how to convert the percentage to a start score. Full steps: Handicap tab, Part 3.',360,160)
note(bc,'B20','Game 11 is 701 doubles, scratch, per the Match Night worksheet.',300,70)

S=wb['Settings']
note(S,'B4','League-average 3DA used to scale the Cheers Rating (3DA ÷ this). Source: Division Average on the DCTV standings page, 24 Nov 2026. Update when you refresh stats.',340,110)
note(S,'B5','League-average MPR used to scale the Cheers Rating. Source: Division Average on the DCTV standings page, 24 Nov 2026.',340,90)
note(S,'B6','ASSUMPTION. How much of the rating comes from 3DA vs MPR. 50% = equal weight. Cricket is 4 of 11 games; 501 doubles and singles are 6 of 11 games; you could weight by that.',340,120)
note(S,'B8','ASSUMPTION. Player of the Week shrink constant K. A week\'s score is pulled toward 100 by darts ÷ (darts + K). With K = 60, a 60-dart week counts half. Raise it to be more cautious about small samples.',340,130)
note(S,'B9','ASSUMPTION. Minimum darts thrown in a week to be eligible for Player of the Week.',300,80)
note(S,'B11','ASSUMPTION. Players with fewer legs than this are flagged Provisional on the ranking. 18 = two full nights (9 legs per player per night).',320,100)
note(S,'A14','Source: your "Pre-Match Handicap Worksheet (Doubles 501 & Cricket)", Part 1 Step 3. Each band applies from its lower bound up to the next one. Spot is subtracted from 501 for the lower team.',340,110)
note(S,'D14','Source: your "Pre-Match Handicap Worksheet", Part 2 Step 3. Marks are given to the lower team and count toward closing numbers only.',340,100)

P=wb['Players']
note(P,'C4',"3DA = three-dart average (points per 3 darts) from the DartConnect League Leaderboard, 'All 01' view: singles and doubles 501/701 legs. Pulled 24 Nov 2026.",340,110)
note(P,'D4','MPR = marks per round from the DartConnect League Leaderboard, Doubles Cricket view. Free handicap marks are not thrown darts and are not counted.',340,110)
note(P,'E4','Cheers Rating = 100 × [w × (3DA ÷ league 3DA) + (1−w) × (MPR ÷ league MPR)]. 100 = league average. w and the league averages are on Settings.',340,110)
note(P,'E5','Example: Bieber J, 3DA 54.56, MPR 2.03 → 100 × [0.5 × (54.56 ÷ 39.90) + 0.5 × (2.03 ÷ 1.35)] = 143.6.',340,90)
note(P,'H4','Legs played from the DCTV standings (LP). Each player throws about 9 legs a night: 7 doubles + 2 singles.',320,90)
note(P,'I4','Provisional = fewer legs than the minimum on Settings B11. Numbers are still settling.',300,80)
note(P,'J4','Optional. 3DA for a player with no league stats yet, e.g. from another league or their DartConnect profile. Used only until they have league stats. Averages from other formats may not match Cheers exactly.',340,120)

H=wb['Handicap']
note(H,'C7','Entered 3DA for a new player with no league stats. Must be between 10 and 100.',300,70)
note(H,'D7','Entered MPR for a new player with no league stats. Must be between 0.10 and 6.00.',300,70)
note(H,'E7','3DA used for this player. Order: 1) Entered 3DA typed in column C (new player), 2) their own league 3DA if they have any league stats, 3) Entered 3DA on the Players tab. If none, the Bar Card asks for the numbers instead of showing handicaps.',360,140)
note(H,'G7','Basis: Entered = you typed both numbers here. Own stats = the player\'s league stats. Entered (Players tab) = numbers kept on the Players tab. Needs 3DA & MPR = new player with nothing entered yet.',340,120)
note(H,'B16','Home team average 3DA = combined 3DA of the two players ÷ 2. Pre-Match worksheet Part 1, Step 1.',320,90)
note(H,'B20','Team average difference = higher team average − lower team average. Pre-Match worksheet Part 1, Step 2.',320,90)
note(H,'B21','Spot = looked up in the chart on Settings (A14:B21). 0.0–2.9 → 0; 3.0–5.9 → 25; 6.0–8.9 → 50; 9.0–11.9 → 75; 12.0–14.9 → 100; 15.0–17.9 → 125; 18.0–20.9 → 150; 21.0+ → 175 (cap). The lower team starts at 501 minus the spot.',360,150)
note(H,'B23','Lower team starts at 501 minus the spot. The alternative on the worksheet is the higher team starting at 501 plus the spot (next row). Same difference either way.',340,110)
note(H,'B27','Combined MPR = the two players\' MPR added (not averaged). Pre-Match worksheet Part 2, Step 1.',320,90)
note(H,'B31','Combined MPR difference = higher team combined MPR − lower team combined MPR. Pre-Match worksheet Part 2, Step 2.',330,90)
note(H,'B32','Starting marks = looked up in the chart on Settings (D14:E19). 0.00–0.49 → 0; 0.50–0.99 → 1; 1.00–1.49 → 2; 1.50–1.99 → 3; 2.00–2.49 → 4; 2.50+ → 5 (cap). Given to the lower team; they count toward closing numbers only.',360,150)
note(H,'G38','Lower player\'s 3DA ÷ higher player\'s 3DA. Example: Cordrey 41.25 ÷ Zeitz 50.10 = 82.3%.',340,110)
note(H,'I38','Start score for the lower player = 501 × the percentage, rounded. The higher player starts at 501. ASSUMPTION: the Match Night worksheet gives the percentage but not this conversion.',340,120)
note(H,'B45','For comparison only. The Match Night worksheet, Step 2, shows "lower ÷ higher = %" for doubles. The Pre-Match worksheet uses difference charts instead. This workbook uses the charts.',340,110)

SR=wb['Season Ranking']
note(SR,'F4','Cheers Rating (see Players tab). 100 = league average. Combines 3DA and MPR; uses every dart a player threw (singles and doubles).',320,100)
note(SR,'I4','Provisional = fewer legs than the minimum (Settings B11).',280,70)

W=wb['Weekly POTW']
note(W,'B6','Baseline = the player\'s season 3DA on the Players tab. Snapshot (paste as values) before refreshing Players each week.',320,100)
note(W,'D6','Enter the player\'s 3DA for the night from DartConnect (all 01 legs).',280,80)
note(W,'E6','Enter the player\'s cricket MPR for the night.',260,70)
note(W,'F6','Total darts the player threw that night (used to pull a short sample toward 100).',280,80)
note(W,'G6','Raw score = 100 × [w × (week 3DA ÷ baseline 3DA) + (1−w) × (week MPR ÷ baseline MPR)]. 100 = playing exactly to their own average.',340,110)
note(W,'H6','Adjusted = 100 + (raw − 100) × darts ÷ (darts + K). Fewer darts pull the score toward 100. K is on Settings B8.',330,100)
note(W,'I6','Eligible = darts thrown at least the minimum on Settings B9.',280,70)

wb.move_sheet('Settings',offset=0)
order=['Bar Card','How It Works','Handicap','Players','Teams','Season Ranking','Weekly POTW','Settings']
wb._sheets=[wb[n] for n in order]
wb.active=0
wb.save(os.path.join(HERE,'Cheers_Monday_Handicap_and_Ranking.xlsx'))
print('ok')

import json as _json
_p=os.path.join(HERE,'setup_sheet.gs')
_g=open(_p).read(); _i=_g.index('var NOTES = ')+len('var NOTES = '); _j=_g.index(';\n  var count')
open(_p,'w').write(_g[:_i]+_json.dumps(NOTES,ensure_ascii=False,indent=1)+_g[_j:])
print('notes',sum(len(v) for v in NOTES.values()))
