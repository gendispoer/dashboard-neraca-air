from __future__ import annotations
from datetime import date, datetime
from io import BytesIO
import re
from typing import Any
from openpyxl import load_workbook

MONTHS = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Ags","Sep","Okt","Nov","Des"]
ALIASES = {"jan":1,"feb":2,"mar":3,"apr":4,"mei":5,"may":5,"jun":6,"jul":7,
           "ags":8,"agt":8,"agu":8,"aug":8,"sep":9,"okt":10,"oct":10,
           "nov":11,"des":12,"dec":12}
G = 9.81
VOLSRC = ["Umbulan","Sumber Air","IP Ngagel I","IP Ngagel II","IP Ngagel III","IP KP I","IP KP II","IP KP III"]
OPS = ["psb","bukaan","tutupan","aanboring","pindah","rendah","tinggi","gantiReg","stopKran","kopling"]
GMB_SIZES = [1,1.5,2,3,4,6,8,16,18]
CM_SIZES = [1,1.5,2,3,4,6,8]
TB_KEYS = ["tkGratis","pmk","asumsiPmk",*OPS,"gantiBesar","cuciMeter","cuciStrainer","flushing","tersier","penghubung"]
LINE_KEYS = ["input","rekening","suplisi","restitusi","tkKom","tkSos","sweeping",*TB_KEYS,
             "curi","airTerjual","salah","bocorSR","bocorTD","luapTangki","luapReservoir"]

def parse_period(period: str) -> tuple[int,int]:
    s=period.strip()
    m=re.fullmatch(r"(\\d{4})-(\\d{1,2})",s)
    if m: y,mo=int(m.group(1)),int(m.group(2))
    else:
        m=re.fullmatch(r"(?:WB_)?([A-Za-z]{3})_(\\d{2}|\\d{4})(?:\\.xlsx)?",s,re.I)
        if not m or m.group(1).lower() not in ALIASES:
            raise ValueError("Periode harus YYYY-MM atau WB_Bulan_Tahun, contoh 2026-08")
        mo=ALIASES[m.group(1).lower()]; yy=int(m.group(2)); y=2000+yy if yy<100 else yy
    if not 1 <= mo <= 12: raise ValueError("Bulan tidak valid")
    return y,mo

def canonical_filename(year:int, month:int)->str:
    return f"WB_{MONTHS[month-1]}_{year%100:02d}.xlsx"

def number(v):
    try:
        if v is None or v == "": return 0.0
        return float(v)
    except (TypeError,ValueError): return 0.0

def _cell(ws, ref): return number(ws[ref].value)

def parse_wb_workbook(content:bytes)->dict[str,Any]:
    wb=load_workbook(BytesIO(content),data_only=False,read_only=False)
    names={n.strip().upper():n for n in wb.sheetnames}
    if "WB" not in names: raise ValueError('Sheet "WB" tidak ditemukan')
    ws=wb[names["WB"]]
    v=lambda ref:_cell(ws,ref)
    def row(r): return {"ptt":v(f"D{r}"),"ptb":v(f"E{r}"),"p":v(f"F{r}"),"d":v(f"G{r}"),"t":v(f"H{r}")}
    def pair(r): return {"distim":v(f"D{r}"),"disbar":v(f"E{r}")}
    inp={
      "volume":[v(f"D{56+i}") for i in range(8)],
      "rekening":v("D73"),"suplisi":v("D74"),"restitusi":v("D75"),
      "tkKomersial":v("D80"),"tkSosial":v("D81"),"sweeping":v("D82"),
      "tkGratis":v("D103"),"pmk":v("D104"),"asumsiPmk":v("D102"),
      "ops":{k:row(92+i) for i,k in enumerate(OPS)},
      "gmb":[row(109+i) for i in range(9)],
      "cm":[row(119+i) for i in range(7)],
      "cs":[row(127+i) for i in range(7)],
      "flush":pair(139),"tersier":pair(140),"penghubung":pair(141),
      "curi":{"timur":v("D147"),"barat":v("E147")},
      "salahPct":v("D152"),"bocorSR":pair(162),"bocorTD":pair(163),
      "luapTangki":v("F164"),"luapReservoir":v("F165")
    }
    wb.close()
    return inp

def write_input_to_workbook(content:bytes, inp:dict[str,Any])->bytes:
    wb=load_workbook(BytesIO(content),data_only=False)
    names={n.strip().upper():n for n in wb.sheetnames}
    if "WB" not in names: raise ValueError('Sheet "WB" tidak ditemukan')
    ws=wb[names["WB"]]
    def put(ref,val): ws[ref]=val
    for i,val in enumerate(inp.get("volume",[])[:8]): put(f"D{56+i}",val)
    for key,ref in [("rekening","D73"),("suplisi","D74"),("restitusi","D75"),
                    ("tkKomersial","D80"),("tkSosial","D81"),("sweeping","D82"),
                    ("tkGratis","D103"),("pmk","D104"),("asumsiPmk","D102"),
                    ("salahPct","D152"),("luapTangki","F164"),("luapReservoir","F165")]:
        put(ref,inp.get(key,0))
    def putrow(r,obj):
        for f,col in [("ptt","D"),("ptb","E"),("p","F"),("d","G"),("t","H")]:
            put(f"{col}{r}",obj.get(f,0))
    for i,k in enumerate(OPS): putrow(92+i,inp.get("ops",{}).get(k,{}))
    for key,start,count in [("gmb",109,9),("cm",119,7),("cs",127,7)]:
        for i,obj in enumerate(inp.get(key,[])[:count]): putrow(start+i,obj)
    for key,r in [("flush",139),("tersier",140),("penghubung",141),("bocorSR",162),("bocorTD",163)]:
        obj=inp.get(key,{})
        put(f"D{r}",obj.get("distim",0)); put(f"E{r}",obj.get("disbar",0))
    curi=inp.get("curi",{})
    put("D147",curi.get("timur",0)); put("E147",curi.get("barat",0))
    out=BytesIO(); wb.save(out); wb.close()
    return out.getvalue()

def _qv(ptt,ptb,p,diameter,time,inch=False):
    dm=number(diameter)*(0.0254 if inch else 1)
    return 0.6*(number(ptt)+number(ptb))*(max(0,2*G*number(p))**0.5)*(0.75*3.141592653589793*dm*dm)*(number(time)*60)

def calc_lines(i):
    L={k:0.0 for k in LINE_KEYS}
    L["input"]=sum(map(number,i.get("volume",[])))
    L["rekening"]=number(i.get("rekening")); L["suplisi"]=number(i.get("suplisi"))
    L["restitusi"]=-number(i.get("restitusi"))
    L["tkKom"]=number(i.get("tkKomersial")); L["tkSos"]=number(i.get("tkSosial"))
    L["sweeping"]=number(i.get("sweeping")); L["tkGratis"]=number(i.get("tkGratis"))
    L["pmk"]=number(i.get("pmk")); L["asumsiPmk"]=number(i.get("asumsiPmk"))
    for k in OPS:
        r=i.get("ops",{}).get(k,{})
        L[k]=_qv(r.get("ptt"),r.get("ptb"),r.get("p"),r.get("d"),r.get("t"))
    L["gantiBesar"]=sum(_qv(**{ "ptt":r.get("ptt"),"ptb":r.get("ptb"),"p":r.get("p"),"diameter":r.get("d"),"time":r.get("t"),"inch":True}) for r in i.get("gmb",[])[:7])
    L["cuciMeter"]=sum(_qv(r.get("ptt"),r.get("ptb"),r.get("p"),r.get("d"),r.get("t"),True) for r in i.get("cm",[])[:6])
    L["cuciStrainer"]=sum(_qv(r.get("ptt"),r.get("ptb"),r.get("p"),r.get("d"),r.get("t"),True) for r in i.get("cs",[])[:6])
    for k in ("flush","tersier","penghubung"):
        obj=i.get(k,{})
        L["flushing" if k=="flush" else k]=number(obj.get("distim"))+number(obj.get("disbar"))
    curi=i.get("curi",{})
    L["curi"]=number(curi.get("timur"))+number(curi.get("barat"))
    L["airTerjual"]=L["rekening"]+L["suplisi"]-L["restitusi"]
    L["salah"]=number(i.get("salahPct"))*L["airTerjual"]/100
    for k in ("bocorSR","bocorTD"):
        obj=i.get(k,{})
        L[k]=number(obj.get("distim"))+number(obj.get("disbar"))
    L["luapTangki"]=number(i.get("luapTangki")); L["luapReservoir"]=number(i.get("luapReservoir"))
    return L

def summarize(L):
    div=lambda a,b:a/b if b else 0.0
    air=L["rekening"]+L["suplisi"]+L["restitusi"]+L["tkKom"]+L["tkSos"]+L["sweeping"]
    tb=sum(L[k] for k in TB_KEYS)
    nrw=L["input"]-air
    nf=L["curi"]+L["salah"]
    fs=L["bocorSR"]+L["bocorTD"]+L["luapTangki"]+L["luapReservoir"]
    resmi=air+tb
    keh=L["input"]-resmi
    belum=keh-nf+fs
    return {"input":L["input"],"airBer":air,"nrw":nrw,"nrwPct":div(nrw,L["input"])*100,
            "revPct":(100-div(nrw,L["input"])*100) if L["input"] else 0,
            "tb":tb,"tbPct":div(tb,nrw)*100,"resmi":resmi,"nf":nf,"nfPct":div(nf,nrw)*100,
            "fs":fs,"fsPct":div(fs,nrw)*100,"keh":keh,"belum":belum,"belumPct":div(belum,nrw)*100}

def aggregate_months(months:list[dict[str,Any]])->dict[str,Any]:
    line_rows=[]
    for m in months:
        lines=calc_lines(m["input"])
        line_rows.append({"period":m["period"],"lines":lines,"summary":summarize(lines)})
    total={k:sum(x["lines"][k] for x in line_rows) for k in LINE_KEYS}
    summary=summarize(total)
    return {"periods":[x["period"] for x in line_rows],"month_count":len(line_rows),
            "monthly":line_rows,"total":{"lines":total,"summary":summary}}
