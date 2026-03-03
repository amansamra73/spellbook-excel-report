from flask import Flask, request, send_file, jsonify
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
from datetime import datetime

app = Flask(__name__)

DB='1F3864'; MB='2E75B6'; LB='BDD7EE'; LG='F2F2F2'; WH='FFFFFF'
GBG='E2EFDA'; GTX='375623'; ABG='FFF2CC'; ATX='7F6000'; RBG='FCE4D6'; RTX='C00000'
PHDR={'Enterprise':'1F5C99','Commercial In-House':'1E5631','SMB Law':'7F5800','Enterprise AM':'4A235A','SMB AM':'7B1519'}
PROW={'Enterprise':'EBF3FB','Commercial In-House':'EBF7EE','SMB Law':'FDFBE6','Enterprise AM':'F5EEF8','SMB AM':'FDECEA'}
PORD=['Enterprise','Commercial In-House','SMB Law','Enterprise AM','SMB AM']
SEASON={1:0.80,2:0.92,3:2.20,4:0.80,5:0.92,6:2.20,7:0.80,8:0.92,9:2.20,10:0.80,11:0.92,12:2.20}

def sd(st='thin',co='BFBFBF'): return Side(style=st,color=co)
def bdr(): return Border(left=sd(),right=sd(),top=sd(),bottom=sd())
def fill(hex): return PatternFill('solid',fgColor=hex)
def fnt(bold=False,italic=False,fc='000000',sz=10): return Font(name='Arial',bold=bold,italic=italic,size=sz,color=fc)
def aln(h='center',wrap=False): return Alignment(horizontal=h,vertical='center',wrap_text=wrap)

def C(ws,r,c,v=None,fmt=None,bold=False,italic=False,fc='000000',bg=None,ha='center',wrap=False,sz=10):
    cell=ws.cell(row=r,column=c)
    if v is not None: cell.value=v
    cell.font=fnt(bold,italic,fc,sz)
    if bg: cell.fill=fill(bg)
    cell.alignment=aln(ha,wrap)
    if fmt: cell.number_format=fmt
    cell.border=bdr()
    return cell

def MH(ws,r,c1,c2,txt,bg=None,fc=WH,sz=14,h=34):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cell=ws.cell(row=r,column=c1,value=txt)
    cell.font=fnt(True,False,fc,sz); cell.fill=fill(bg or DB); cell.alignment=aln('center')
    ws.row_dimensions[r].height=h

def SEC(ws,r,c1,c2,txt,bg=None,h=20):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cell=ws.cell(row=r,column=c1,value=txt)
    cell.font=fnt(True,False,WH,11); cell.fill=fill(bg or MB); cell.alignment=aln('left')
    ws.row_dimensions[r].height=h

def HDR(ws,r,c,txt):
    C(ws,r,c,txt,bold=True,bg=DB,fc=WH,ha='center',wrap=True,sz=9)
    ws.row_dimensions[r].height=28

def ATT(ws,r,c,pct):
    bg=GBG if pct>=1.0 else(ABG if pct>=0.75 else RBG)
    fc=GTX if pct>=1.0 else(ATX if pct>=0.75 else RTX)
    C(ws,r,c,pct,'0.0%',bold=True,bg=bg,fc=fc)

def GAP(ws,r,h=7): ws.row_dimensions[r].height=h

def build_excel(data):
    summary=data['summary']; podStats=data['podStats']
    repSummaries=data['repSummaries']
    top10Deals=data.get('top10Deals',data.get('top5Deals',[]))
    thisWeekDeals=data.get('thisWeekDeals',[])
    RL=f"Week Ending {datetime.now().strftime('%b %-d, %Y')}"
    wb=openpyxl.Workbook()

    # ── SHEET 1: EXECUTIVE SUMMARY ────────────────────────────────────────────
    ws=wb.active; ws.title='Executive Summary'
    ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=90
    for col,w in zip('ABCD',[32,16,16,14]): ws.column_dimensions[col].width=w
    MH(ws,1,1,4,'Spellbook Legal — Sales Performance Report',DB,WH,15,38)
    MH(ws,2,1,4,RL,MB,WH,11,22); GAP(ws,3)
    SEC(ws,4,1,4,'  COMPANY OVERVIEW — YTD 2026')
    for i,h in enumerate(['Metric','YTD Actual','YTD Target','Attainment']): HDR(ws,5,i+1,h)
    tot_tgt=summary['totalNBTarget']+summary['totalExpTarget']
    for i,(lbl,act,tgt,att) in enumerate([
        ('Total Revenue',summary['totalRevenue'],tot_tgt,summary['totalRevenue']/tot_tgt if tot_tgt else 0),
        ('New Business',summary['totalNB'],summary['totalNBTarget'],summary['totalNB']/summary['totalNBTarget'] if summary['totalNBTarget'] else 0),
        ('Expansion Revenue',summary['totalExp'],summary['totalExpTarget'],summary['totalExp']/summary['totalExpTarget'] if summary['totalExpTarget'] else 0),
        ('Total Deals',summary['totalDeals'],None,None)]):
        r=6+i; bg=LG if i%2==0 else WH
        C(ws,r,1,lbl,bold=True,bg=bg,ha='left'); C(ws,r,2,act,'$#,##0' if tgt else '#,##0',bg=bg)
        C(ws,r,3,tgt if tgt else '—','$#,##0' if tgt else None,bg=bg)
        if att is not None: ATT(ws,r,4,att)
        else: C(ws,r,4,'—',bg=bg)
        ws.row_dimensions[r].height=18
    GAP(ws,10)
    SEC(ws,11,1,4,'  YEAR-OVER-YEAR — Same Period Est.')
    for i,h in enumerate(['Metric','2026 YTD','2025 Est.','YoY Change']): HDR(ws,12,i+1,h)
    for i,(lbl,v26,v25) in enumerate([
        ('Total Revenue',summary['totalRevenue'],summary['pace2025Total']),
        ('New Business',summary['totalNB'],summary['pace2025NB']),
        ('Expansion',summary['totalExp'],summary['pace2025Exp'])]):
        r=13+i; bg=LG if i%2==0 else WH; chg=(v26-v25)/v25 if v25 else 0
        C(ws,r,1,lbl,bold=True,bg=bg,ha='left'); C(ws,r,2,v26,'$#,##0',bg=bg); C(ws,r,3,v25,'$#,##0',bg=bg)
        C(ws,r,4,chg,'+0.0%;-0.0%;-',bold=True,bg=GBG if chg>=0 else RBG,fc=GTX if chg>=0 else RTX)
        ws.row_dimensions[r].height=18
    GAP(ws,16)
    SEC(ws,17,1,4,'  2026 FULL YEAR FORECAST (Q-end surge: Mar/Jun/Sep/Dec = 2.2x)')
    for i,h in enumerate(['Metric','2025 Full Year','2026 Projected','YoY Growth']): HDR(ws,18,i+1,h)
    fy25T=summary['pace2025Total']*6; fy25N=summary['pace2025NB']*6; fy25E=summary['pace2025Exp']*6
    aR=summary['totalRevenue']/2; aN=summary['totalNB']/2; aE=summary['totalExp']/2
    sv=sum(SEASON.values())
    pR=aR*sv*0.6+fy25T*1.15*0.4; pN=aN*sv*0.6+fy25N*1.15*0.4; pE=aE*sv*0.6+fy25E*1.20*0.4
    for i,(lbl,fy,pj) in enumerate([('Total Revenue',fy25T,pR),('New Business',fy25N,pN),('Expansion',fy25E,pE)]):
        r=19+i; bg=LG if i%2==0 else WH; chg=(pj-fy)/fy if fy else 0
        C(ws,r,1,lbl,bold=True,bg=bg,ha='left'); C(ws,r,2,round(fy),'$#,##0',bg=bg); C(ws,r,3,round(pj),'$#,##0',bg=bg)
        C(ws,r,4,chg,'+0.0%;-0.0%;-',bold=True,bg=GBG if chg>=0 else RBG,fc=GTX if chg>=0 else RTX)
        ws.row_dimensions[r].height=18

    # ── SHEET 2: POD PERFORMANCE ──────────────────────────────────────────────
    ws2=wb.create_sheet('Pod Performance')
    ws2.sheet_view.showGridLines=False; ws2.sheet_view.zoomScale=90
    for col,w in zip('ABCDEFGHIJK',[24,14,14,12,14,14,12,14,14,12,10]): ws2.column_dimensions[col].width=w
    MH(ws2,1,1,11,f'Pod Performance — {RL}',DB,WH,14,34); GAP(ws2,2)

    # NB section
    SEC(ws2,3,1,11,'  NEW BUSINESS BY POD — YTD 2026')
    for i,h in enumerate(['Pod','Jan NB','Jan Target','Jan Att%','Feb NB','Feb Target','Feb Att%','Mar NB (partial)','Mar Target','Mar Att%','YTD NB']): HDR(ws2,4,i+1,h)
    for i,pod in enumerate(['Enterprise','Commercial In-House','SMB Law']):
        r=5+i; bg=LG if i%2==0 else WH; ps=podStats[pod]
        jt=ps.get('janNBTarget',0); ft=ps.get('febNBTarget',0); mt=ps.get('marNBTarget',0)
        C(ws2,r,1,pod,bold=True,bg=bg,ha='left')
        C(ws2,r,2,round(ps.get('janNB',0)),'$#,##0',bg=bg)
        C(ws2,r,3,jt,'$#,##0',bg=bg)
        ATT(ws2,r,4,ps.get('janNB',0)/jt if jt else 0)
        C(ws2,r,5,round(ps.get('febNB',0)),'$#,##0',bg=bg)
        C(ws2,r,6,ft,'$#,##0',bg=bg)
        ATT(ws2,r,7,ps.get('febNB',0)/ft if ft else 0)
        C(ws2,r,8,round(ps.get('marNB',0)),'$#,##0',bg=bg)
        C(ws2,r,9,mt,'$#,##0',bg=bg)
        ATT(ws2,r,10,ps.get('marNB',0)/mt if mt else 0)
        C(ws2,r,11,round(ps['newBiz']),'$#,##0',bold=True,bg=bg)
        ws2.row_dimensions[r].height=18

    # Exp section
    GAP(ws2,8)
    SEC(ws2,9,1,11,'  EXPANSION BY POD — YTD 2026')
    for i,h in enumerate(['Pod','Jan Exp','Jan Target','Jan Att%','Feb Exp','Feb Target','Feb Att%','Mar Exp (partial)','Mar Target','Mar Att%','YTD Exp']): HDR(ws2,10,i+1,h)
    for i,(pod,amPod) in enumerate([('Enterprise','Enterprise AM'),('SMB Law','SMB AM')]):
        r=11+i; bg=LG if i%2==0 else WH; ps=podStats[pod]; am=podStats[amPod]
        totalJanExp=ps.get('janExp',0)+am.get('janExp',0)
        totalFebExp=ps.get('febExp',0)+am.get('febExp',0)
        totalMarExp=ps.get('marExp',0)+am.get('marExp',0)
        totalExp=ps['expansion']+am['expansion']
        jt=ps.get('janExpTarget',0); ft=ps.get('febExpTarget',0); mt=ps.get('marExpTarget',0)
        C(ws2,r,1,pod,bold=True,bg=bg,ha='left')
        C(ws2,r,2,round(totalJanExp),'$#,##0',bg=bg); C(ws2,r,3,jt,'$#,##0',bg=bg)
        ATT(ws2,r,4,totalJanExp/jt if jt else 0)
        C(ws2,r,5,round(totalFebExp),'$#,##0',bg=bg); C(ws2,r,6,ft,'$#,##0',bg=bg)
        ATT(ws2,r,7,totalFebExp/ft if ft else 0)
        C(ws2,r,8,round(totalMarExp),'$#,##0',bg=bg); C(ws2,r,9,mt,'$#,##0',bg=bg)
        ATT(ws2,r,10,totalMarExp/mt if mt else 0)
        C(ws2,r,11,round(totalExp),'$#,##0',bold=True,bg=bg)
        ws2.row_dimensions[r].height=18

    # ── SHEET 3: INDIVIDUAL PERFORMANCE ──────────────────────────────────────
    ws3=wb.create_sheet('Individual Performance')
    ws3.sheet_view.showGridLines=False; ws3.sheet_view.zoomScale=75
    cols3='ABCDEFGHIJKLMNOPQRSTU'
    widths3=[20,26,16,12,12,10,8,10,10,10,9,10,10,9,10,10,9,10,10,12,12]
    for col,w in zip(cols3,widths3): ws3.column_dimensions[col].width=w
    MH(ws3,1,1,21,f'Individual Rep Performance — {RL}',DB,WH,14,34); GAP(ws3,2)
    hdrs3=['Rep','Role','Pod','YTD Quota','YTD Revenue','Att%','Deals','Avg Deal',
           'Jan Quota','Jan Rev','Jan Att%','Feb Quota','Feb Rev','Feb Att%',
           'Mar Quota','Mar Rev','Mar Att%','New Biz','Expansion','2025 FY','YoY']
    for i,h in enumerate(hdrs3): HDR(ws3,3,i+1,h)
    ws3.freeze_panes='A4'

    row=4
    for pod in PORD:
        reps=sorted([r for r in repSummaries if r['pod']==pod],key=lambda x:x['ytdRevenue'],reverse=True)
        if not reps: continue
        ws3.merge_cells(start_row=row,start_column=1,end_row=row,end_column=21)
        hc=ws3.cell(row=row,column=1,value=f'  {pod.upper()}')
        hc.font=fnt(True,False,WH,10); hc.fill=fill(PHDR.get(pod,MB)); hc.alignment=aln('left')
        ws3.row_dimensions[row].height=18; row+=1
        bg=PROW.get(pod,WH); pR=pQ=pN=pE=pD=0
        for rep in reps:
            att=rep['ytdRevenue']/rep['ytdQuota'] if rep['ytdQuota'] else None
            janA=rep['janRevenue']/rep['janQuota'] if rep.get('janQuota') else None
            febA=rep['febRevenue']/rep['febQuota'] if rep.get('febQuota') else None
            marA=rep['marRevenue']/rep['marQuota'] if rep.get('marQuota') else None
            est25=rep['fy2025Revenue']*2/12 if rep.get('fy2025Revenue') else None
            yoy=(rep['ytdRevenue']-est25)/est25 if est25 else None
            C(ws3,row,1,rep['rep'],bold=True,bg=bg,ha='left',sz=9)
            C(ws3,row,2,rep['role'],bg=bg,ha='left',sz=8)
            C(ws3,row,3,pod,bg=bg,ha='left',sz=9)
            C(ws3,row,4,rep['ytdQuota'],'$#,##0',bg=bg)
            C(ws3,row,5,rep['ytdRevenue'],'$#,##0',bg=bg)
            if att is not None: ATT(ws3,row,6,att)
            else: C(ws3,row,6,'—',bg=bg)
            C(ws3,row,7,rep['ytdDeals'],'#,##0',bg=bg)
            C(ws3,row,8,rep['avgDeal'],'$#,##0',bg=bg)
            C(ws3,row,9,rep.get('janQuota',0),'$#,##0',bg=bg)
            C(ws3,row,10,rep.get('janRevenue',0),'$#,##0',bg=bg)
            if janA is not None: ATT(ws3,row,11,janA)
            else: C(ws3,row,11,'—',bg=bg)
            C(ws3,row,12,rep.get('febQuota',0),'$#,##0',bg=bg)
            C(ws3,row,13,rep.get('febRevenue',0),'$#,##0',bg=bg)
            if febA is not None: ATT(ws3,row,14,febA)
            else: C(ws3,row,14,'—',bg=bg)
            C(ws3,row,15,rep.get('marQuota',0),'$#,##0',bg=bg)
            C(ws3,row,16,rep.get('marRevenue',0),'$#,##0',bg=bg)
            if marA is not None: ATT(ws3,row,17,marA)
            else: C(ws3,row,17,'—',bg=bg)
            C(ws3,row,18,rep['ytdNewBiz'],'$#,##0',bg=bg)
            C(ws3,row,19,rep['ytdExpansion'],'$#,##0',bg=bg)
            C(ws3,row,20,rep['fy2025Revenue'] if rep.get('fy2025Revenue') else '—','$#,##0' if rep.get('fy2025Revenue') else None,bg=bg)
            if yoy is not None: C(ws3,row,21,yoy,'+0.0%;-0.0%;-',bold=True,bg=GBG if yoy>=0 else RBG,fc=GTX if yoy>=0 else RTX)
            else: C(ws3,row,21,'N/A',bg=bg)
            pR+=rep['ytdRevenue']; pQ+=rep['ytdQuota']; pN+=rep['ytdNewBiz']; pE+=rep['ytdExpansion']; pD+=rep['ytdDeals']
            ws3.row_dimensions[row].height=17; row+=1
        C(ws3,row,1,f'{pod} SUBTOTAL',bold=True,bg=LB,ha='left')
        for ci in [2,3]: C(ws3,row,ci,'',bg=LB)
        C(ws3,row,4,pQ,'$#,##0',bold=True,bg=LB); C(ws3,row,5,pR,'$#,##0',bold=True,bg=LB)
        ATT(ws3,row,6,pR/pQ if pQ else 0)
        C(ws3,row,7,pD,'#,##0',bold=True,bg=LB)
        for ci in range(8,22): C(ws3,row,ci,'',bg=LB)
        ws3.row_dimensions[row].height=20; row+=2

    # ── SHEET 4: TOP 10 DEALS YTD ─────────────────────────────────────────────
    ws4=wb.create_sheet('Top 10 Deals YTD')
    ws4.sheet_view.showGridLines=False
    for col,w in zip('ABCDE',[50,22,20,16,20]): ws4.column_dimensions[col].width=w
    MH(ws4,1,1,5,f'Top 10 Deals YTD — {RL}',DB,WH,14,34); GAP(ws4,2)
    for i,h in enumerate(['Deal Name','Owner','Pipeline','Amount','Revenue Start Date']): HDR(ws4,3,i+1,h)
    for i,deal in enumerate(top10Deals):
        r=4+i; bg=LG if i%2==0 else WH
        C(ws4,r,1,deal['dealname'],bg=bg,ha='left')
        C(ws4,r,2,deal['owner'],bg=bg); C(ws4,r,3,deal['pipeline'],bg=bg)
        C(ws4,r,4,deal['amount'],'$#,##0',bold=True,bg=bg)
        C(ws4,r,5,deal['revenue_start_date'],bg=bg)
        ws4.row_dimensions[r].height=18

    # ── SHEET 5: THIS WEEK'S DEALS ────────────────────────────────────────────
    ws5=wb.create_sheet("This Week's Deals")
    ws5.sheet_view.showGridLines=False
    for col,w in zip('ABCDE',[50,22,20,16,20]): ws5.column_dimensions[col].width=w
    MH(ws5,1,1,5,f"This Week's Deals — {RL}",DB,WH,14,34); GAP(ws5,2)
    for i,h in enumerate(['Deal Name','Owner','Pipeline','Amount','Revenue Start Date']): HDR(ws5,3,i+1,h)
    if thisWeekDeals:
        for i,deal in enumerate(thisWeekDeals):
            r=4+i; bg=LG if i%2==0 else WH
            C(ws5,r,1,deal['dealname'],bg=bg,ha='left')
            C(ws5,r,2,deal['owner'],bg=bg); C(ws5,r,3,deal['pipeline'],bg=bg)
            C(ws5,r,4,deal['amount'],'$#,##0',bold=True,bg=bg)
            C(ws5,r,5,deal['revenue_start_date'],bg=bg)
            ws5.row_dimensions[r].height=18
    else:
        ws5.merge_cells('A4:E4')
        c=ws5.cell(row=4,column=1,value='No deals with revenue start date in the last 7 days.')
        c.font=fnt(False,True,'888888',10); c.alignment=aln('left')
        ws5.row_dimensions[4].height=20

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

@app.route('/health',methods=['GET'])
def health():
    return jsonify({'status':'ok','service':'spellbook-excel-report'})

@app.route('/generate-report',methods=['POST'])
def generate_report():
    try:
        data=request.get_json(force=True)
        if not data: return jsonify({'error':'No JSON body received'}),400
        buf=build_excel(data)
        filename=f"Sales_Report_{datetime.now().strftime('%b%-d_%Y')}.xlsx"
        return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,download_name=filename)
    except Exception as e:
        import traceback
        return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

if __name__=='__main__':
    app.run(host='0.0.0.0',port=8080)
