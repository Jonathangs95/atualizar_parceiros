# -*- coding: utf-8 -*-
"""Atualizacao de parceiros residencial - Windows, Excel desktop e pywin32.

Instalar: python -m pip install pywin32
Simular: python atualizar_parceiros.py --pasta-base "C:\caminho\Residencial"
Aplicar: python atualizar_parceiros.py --pasta-base "C:\caminho\Residencial" --aplicar

A pasta deve conter atualizarParceiros.xlsm, New_Dash_Fixa_DATA.xlsb,
os arquivos dos parceiros e Base\BaseDash.xlsx (ou BaseDash.xlsx).
Por seguranca, somente substitui o mes de referencia quando ele esta no
final da Base; se houver meses posteriores/intercalados, interrompe.
"""
from __future__ import annotations
import argparse
from collections import Counter
from datetime import date, datetime, timedelta
import logging
from pathlib import Path
import sys
import time
import unicodedata

LOG = logging.getLogger("parceiros")
VERSAO = "5.0"
LOTE = 10000
XL_UP = -4162
XL_LEFT = -4159
XL_MANUAL = -4135
XL_AUTO = -4105
XL_FORMATS = -4122

# Chave = destino; valor = origem. Nao copiar NM_EQUIPE_VENDA.
MAPA = {
    "ANO_MES": "NR_ANO_MES", "DT": "DT",
    "INDICADOR": "NM_INDICADOR_NEGOCIO", "MARCA": "NM_PERFIL_CLIENTE",
    "PERFIL": "NM_LINHA_NEGOCIO", "PROD": "NM_TIPO_ASS_DOMICILIO",
    "NR_CONTRATO": "NR_CONTRATO", "NM_CANAL_VENDA_GRUPO": "NM_CANAL_VENDA_GRUPO",
    "COD_MUNICIPIO": "CD_MUNICIPIO", "PRODUTO_ATUAL": "PRODUTO_ATUAL",
    "QT": "QT", "CD_LOGIN_VENDEDOR": "CD_LOGIN_VENDEDOR",
    "DDD MUNICÍPIO": "DDD_MUN", "COD": "COD_PDV", "CIDADE": "NM_CIDADE",
    "GR": "DS_RAZAO_SOCIAL", "PLANO AGRUPADO": "PLANO_AGRUPADO",
    "CANAL": "CANAL", "GN": "NM_GN", "Tipo": "TP_LOCALIDADE",
    "CID CRITICA": "CIDADE_CRITICA", "MERC": "MERCADO",
    "HC": "d_PDV.HC Atual", "Semana": "Semana do Mês", "Tipo Cid": "TP_CIDADE",
}
# Campos complementares: somente formula historica uniforme ou historico vazio.
# LOJA nao era preenchida pela macro original; nao inventar regra.
OPCIONAIS = (
    "SUBTIPO", "Recente Gpon?", "OFERTA", "CLASSIF", "TELECINE",
    "LOJA", "Meta", "Orç", "TP Mplay", "CIDADE PDV", "META PME",
)

def norm(v):
    s = unicodedata.normalize("NFKD", str(v or "").strip().casefold())
    return "".join(c for c in s if c.isalnum() and not unicodedata.combining(c))

def chave(v):
    return str(v or "").strip().casefold()

def mes(v):
    if isinstance(v, (int, float)) and int(v) == v:
        return str(int(v))
    return str(v or "").strip()

def vazio(v):
    return v is None or v == ""

def retry(fn, desc):
    for i in range(60):
        try:
            return fn()
        except Exception as e:
            if getattr(e, "hresult", None) not in (-2147418111, -2147417846) or i == 59:
                raise
            if i == 0:
                LOG.warning("Excel ocupado: %s", desc)
            import pythoncom
            pythoncom.PumpWaitingMessages()
            time.sleep(1)

def rg(ws, r1, c1, r2, c2):
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))

def matriz(v, n, k):
    if n == k == 1:
        if not isinstance(v, (tuple, list)):
            return ((v,),)
        if not isinstance(v[0], (tuple, list)):
            return (tuple(v),)
    return v

def ler(ws, a, b, c, d, prop="Value2"):
    return matriz(getattr(rg(ws, a, b, c, d), prop), c-a+1, d-b+1)

def indices(headers):
    out = {}
    for i, h in enumerate(headers, 1):
        k = norm(h)
        if not k or k in out:
            raise ValueError("Cabecalho vazio ou duplicado: %r" % h)
        out[k] = i
    return out

def ultimo(ws):
    return int(ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row)

def headers(ws):
    n = int(ws.Cells(1, ws.Columns.Count).End(XL_LEFT).Column)
    return list(ler(ws, 1, 1, 1, n)[0])

def abrir(excel, path, readonly=True):
    path = Path(path).resolve()
    if not path.is_file():
        raise FileNotFoundError(str(path))
    LOG.info("Abrindo %s", path)
    wb = retry(lambda: excel.Workbooks.Open(Filename=str(path), UpdateLinks=0,
               ReadOnly=readonly, IgnoreReadOnlyRecommended=True,
               AddToMru=False, Notify=False), "abrir " + path.name)
    try:
        if not readonly:
            if retry(lambda: wb.ReadOnly, "ReadOnly"):
                raise PermissionError("Arquivo em uso: " + str(path))
            try:
                autosave = retry(lambda: wb.AutoSaveOn, "AutoSave")
            except AttributeError:
                autosave = False
            except Exception as exc:
                if getattr(exc, "hresult", None) not in (-2147352573, -2147352570):
                    raise
                autosave = False
            if autosave:
                retry(lambda: setattr(wb, "AutoSaveOn", False), "desativar AutoSave")
                if wb.AutoSaveOn:
                    raise RuntimeError("Nao foi possivel desativar AutoSave.")
        return wb
    except Exception:
        fechar(wb)
        raise

def fechar(wb):
    if wb is not None:
        retry(lambda: wb.Close(SaveChanges=False), "fechar workbook")

def parceiros(ws, pasta, ref):
    out, vistos, nomes = [], set(), set()
    for lin in range(2, 5):
        a,b,c,d,e,f,g,h,i,j,k = ler(ws,lin,1,lin,11)[0]
        if a and b and d:
            e = f"{a}_{b}_{ref:%Y-%m-%d}{d}"
        if not e or not k:
            raise ValueError(f"Controle incompleto na linha {lin} (E/K).")
        path = (pasta / str(e).strip()).resolve()
        if path.parent != pasta or path.suffix.lower() != ".xlsb":
            raise ValueError("Caminho de parceiro inesperado: " + str(path))
        if path in vistos or chave(k) in nomes:
            raise ValueError("Parceiro ou arquivo duplicado no controle.")
        vistos.add(path)
        nomes.add(chave(k))
        out.append((lin, str(k).strip(), path))
    return out

def carregar(ws, ref, ps):
    tab = ws.ListObjects.Item("f_Assinaturas")
    if tab.HeaderRowRange.Row != 1 or tab.Range.Column != 1:
        raise ValueError("f_Assinaturas deve iniciar em A1.")
    n = int(tab.ListColumns.Count)
    hs = list(ler(ws,1,1,1,n)[0])
    ix = indices(hs)
    missing = [h for h in MAPA.values() if norm(h) not in ix]
    if missing:
        raise ValueError("Origem sem campos: " + ", ".join(missing))
    LOG.info("Origem: %s colunas | %s", n, hs)
    out = {chave(p): [] for _,p,_ in ps}
    body = tab.DataBodyRange
    if body is None:
        raise ValueError("Origem vazia.")
    a = int(body.Row)
    end = a + int(body.Rows.Count) - 1
    im, ip, iw = (ix[norm(x)]-1 for x in ("NR_ANO_MES","DS_RAZAO_SOCIAL","Semana do Mês"))
    for start in range(a,end+1,LOTE):
        for row in ler(ws,start,1,min(end,start+LOTE-1),n):
            if mes(row[im]) == ref and chave(row[ip]) in out:
                if str(row[iw]).strip() not in {"S1","S2","S3","S4","S5","S6"}:
                    raise ValueError("Semana do Mes invalida na origem.")
                out[chave(row[ip])].append(tuple(row))
    for p, rows in out.items():
        LOG.info("Origem filtrada | %s | %s registros | semanas %s",
                 p,len(rows),dict(Counter(r[iw] for r in rows)))
        if not rows:
            raise ValueError("Origem sem registros de " + p)
    return ix,out

def plano(ws, src, ref, nome, rows):
    hs = headers(ws)
    dest = indices(hs)
    if dest.get(norm("ANO_MES")) != 1:
        raise ValueError("A Base deve iniciar com ANO_MES em A1.")
    if any(norm(h) not in dest for h in MAPA):
        raise ValueError("Destino sem campos obrigatorios: " +
                         ", ".join(h for h in MAPA if norm(h) not in dest))
    mapa = {dest[norm(d)]:src[norm(s)]-1 for d,s in MAPA.items()}
    end = ultimo(ws)
    if end < 2:
        raise ValueError("Base sem historico.")
    first = None
    ip = dest[norm("GR")]
    # Valida o mes e o GR de todo o historico preservado, em blocos.
    for a in range(2,end+1,LOTE):
        b = min(end,a+LOTE-1)
        ms = ler(ws,a,1,b,1)
        gs = ler(ws,a,ip,b,ip)
        for j,(m,g) in enumerate(zip(ms,gs),a):
            mm = mes(m[0])
            if mm == ref:
                if first is None:
                    first = j
            else:
                if first is not None:
                    raise ValueError("Mes de referencia intercalado; historico preservado.")
                if not mm.isdigit() or len(mm)!=6 or mm > ref:
                    raise ValueError(f"Mes invalido ou posterior na linha {j}: {mm}")
                if chave(g[0]) != chave(nome):
                    raise ValueError(f"GR historico divergente na linha {j}; arquivo preservado.")
    start = first if first is not None else end+1
    if start <= 2:
        raise ValueError("Sem historico anterior para validar o parceiro.")
    nold = end-start+1
    final = start+len(rows)-1
    if final > int(ws.Rows.Count):
        raise ValueError("Limite de linhas Excel excedido.")
    if final > end:
        # Nao sobrescrever dados ou formulas fora da faixa de registros.
        tail = ler(ws,end+1,1,final,len(hs),"FormulaR1C1")
        if any(not vazio(v) for r in tail for v in r):
            raise ValueError("Existem dados abaixo da Base; expansao insegura.")
    tab = None
    for t in ws.ListObjects:
        if t.Range.Column == 1 and t.HeaderRowRange.Row == 1:
            tab = t
            if t.ShowTotals or int(t.Range.Columns.Count) != len(hs) or int(t.Range.Rows.Count) != end:
                raise ValueError("Tabela destino possui estrutura diferente da Base.")
        elif int(t.Range.Row) <= max(end,final):
            raise ValueError("Outra tabela na area de atualizacao.")
    formulas, blanks = {}, []
    permitidos = {norm(x) for x in OPCIONAIS}
    for j,h in enumerate(hs,1):
        if j in mapa:
            LOG.info("MAPA | destino %s %s <- origem %s",j,h,mapa[j]+1)
            continue
        if norm(h) not in permitidos:
            raise ValueError("Campo destino sem regra: " + str(h))
        lo = max(2,start-100)
        fs = [r[0] for r in ler(ws,lo,j,start-1,j,"FormulaR1C1")]
        if all(vazio(x) for x in fs):
            blanks.append(j)
        elif all(isinstance(x,str) and x.startswith("=") for x in fs) and len(set(fs)) == 1:
            formulas[j] = fs[0]
        else:
            raise ValueError("Campo complementar com regra historica indefinida: " + str(h))
    LOG.info("Plano | %s | %s antigos | %s novos | linhas %s:%s | %s formulas | %s vazias",
             nome,nold,len(rows),start,final,len(formulas),len(blanks))
    return hs,mapa,formulas,blanks,start,end,final,tab

def blocos(cols):
    out = []
    for c in sorted(cols):
        if out and c == out[-1][-1]+1:
            out[-1].append(c)
        else:
            out.append([c])
    return out

def gravar(ws, rows, p):
    hs,mapa,formulas,blanks,start,end,final,tab = p
    if tab is not None:
        tab.Resize(rg(ws,1,1,final,len(hs)))
    rg(ws,start-1,1,start-1,len(hs)).Copy()
    rg(ws,start,1,final,len(hs)).PasteSpecial(Paste=XL_FORMATS)
    ws.Application.CutCopyMode = False
    if end > final:
        rg(ws,final+1,1,end,len(hs)).ClearContents()
    for cs in blocos(mapa):
        for off in range(0,len(rows),LOTE):
            batch = rows[off:off+LOTE]
            a,b = start+off,start+off+len(batch)-1
            vals = tuple(tuple(r[mapa[c]] for c in cs) for r in batch)
            rg(ws,a,cs[0],b,cs[-1]).Value2 = vals
            got = ler(ws,a,cs[0],b,cs[-1])
            if len(got) != len(vals):
                raise RuntimeError("Quantidade gravada divergente.")
            for i,(x,y) in enumerate(zip(got,vals),a):
                if len(x)!=len(y) or any(u!=v and not (vazio(u) and vazio(v)) for u,v in zip(x,y)):
                    raise RuntimeError(f"Conferencia divergente na linha {i}, colunas {cs}.")
    for cs in blocos(blanks):
        rg(ws,start,cs[0],final,cs[-1]).ClearContents()
    for c,f in formulas.items():
        rg(ws,start,c,final,c).FormulaR1C1 = f
    if ultimo(ws) != final:
        raise RuntimeError("Ultima linha apos gravacao divergente.")
    # Confere novamente os campos mapeados depois das formulas e da limpeza.
    for cs in blocos(mapa):
        for off in range(0,len(rows),LOTE):
            batch = rows[off:off+LOTE]
            a,b = start+off,start+off+len(batch)-1
            got = ler(ws,a,cs[0],b,cs[-1])
            expected = tuple(tuple(r[mapa[c]] for c in cs) for r in batch)
            if len(got)!=len(expected) or any(len(x)!=len(y) or any(u!=v and not(vazio(u) and vazio(v)) for u,v in zip(x,y)) for x,y in zip(got,expected)):
                raise RuntimeError("Conferencia final divergente.")
    LOG.info("Conferencia final: %s linhas e %s campos",len(rows),len(mapa))

def horario(ws, lin, col):
    now = datetime.now()
    cell = ws.Cells(lin,col)
    cell.Value2 = (now.hour*3600+now.minute*60+now.second)/86400
    cell.NumberFormat = "hh:mm:ss"

def executar(args, downloads):
    import pythoncom
    import win32com.client
    pasta = Path(args.pasta_base).resolve()
    ref = date.fromisoformat(args.data_ref) if args.data_ref else date.today()-timedelta(days=1)
    ym = ref.strftime("%Y%m")
    control = Path(args.controle).resolve() if args.controle else pasta/"atualizarParceiros.xlsm"
    source = Path(args.origem).resolve() if args.origem else pasta/"Base"/"BaseDash.xlsx"
    if not args.origem and not source.is_file():
        source = pasta/"BaseDash.xlsx"
    trend = Path(args.tendencia).resolve() if args.tendencia else pasta/f"New_Dash_Fixa_{ref:%Y-%m-%d}.xlsb"
    LOG.info("Versao %s | %s | %s",VERSAO,ym,"APLICAR" if args.aplicar else "SIMULACAO")
    for path in (control,source,trend):
        if not path.is_file():
            raise FileNotFoundError(str(path))
    backup = downloads/"Backup_Atualizar_Parceiros"/datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pythoncom.CoInitialize()
    excel = wc = wb = wt = None
    auto_fill = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        wc = abrir(excel,control,not args.aplicar)
        retry(lambda: setattr(excel,"Calculation",XL_MANUAL),"calculo manual")
        try:
            auto_fill = excel.AutoCorrect.AutoFillFormulasInLists
            excel.AutoCorrect.AutoFillFormulasInLists = False
        except Exception:
            LOG.warning("Nao foi possivel desativar preenchimento automatico de tabelas.")
        ctrl = wc.Worksheets.Item("Parceiro")
        ps = parceiros(ctrl,pasta,ref)
        if any(p in (control,source,trend) for _,_,p in ps):
            raise ValueError("Arquivo de parceiro coincide com a origem/controle.")
        wb = abrir(excel,source)
        src, grupos = carregar(wb.Worksheets.Item("f_Assinaturas"),ym,ps)
        fechar(wb)
        wb = None
        wt = abrir(excel,trend)
        values = wt.Worksheets.Item("Projeção").Range("S2:S5").Value2
        if len(values)!=4 or any(vazio(r[0]) for r in values):
            raise ValueError("Tendencia S2:S5 incompleta.")
        fechar(wt)
        wt = None
        # Preflight dos tres arquivos antes de qualquer gravacao.
        for lin,nome,path in ps:
            w = abrir(excel,path)
            try:
                w.Worksheets.Item("Aux")
                w.Worksheets.Item("LOJA")
                plano(w.Worksheets.Item("Base"),src,ym,nome,grupos[chave(nome)])
            finally:
                fechar(w)
        if not args.aplicar:
            LOG.info("SIMULACAO CONCLUIDA: todos os parceiros validados; nenhum arquivo salvo.")
            return 0
        backup.mkdir(parents=True,exist_ok=False)
        retry(lambda: wc.SaveCopyAs(str(backup/control.name)),"backup controle")
        erros = 0
        for lin,nome,path in ps:
            w = None
            salvo = False
            horario(ctrl,lin,6)
            ctrl.Cells(lin,7).Value2 = "Em andamento"
            try:
                w = abrir(excel,path,False)
                base = w.Worksheets.Item("Base")
                p = plano(base,src,ym,nome,grupos[chave(nome)])
                retry(lambda: w.SaveCopyAs(str(backup/path.name)),"backup parceiro")
                if base.FilterMode:
                    base.ShowAllData()
                gravar(base,grupos[chave(nome)],p)
                aux = w.Worksheets.Item("Aux")
                aux.Visible = -1
                aux.Range("K5:K8").Value2 = values
                w.Activate()
                w.Worksheets.Item("LOJA").Activate()
                aux.Visible = 0
                retry(lambda: excel.Calculate(),"recalcular")
                retry(lambda: setattr(excel,"Calculation",XL_AUTO),"calculo automatico")
                retry(lambda: w.Save(),"salvar parceiro")
                salvo = True
                ctrl.Cells(lin,7).Value2 = "Concluído"
                LOG.info("CONCLUIDO | %s | %s linhas",nome,len(grupos[chave(nome)]))
            except Exception:
                erros += 1
                ctrl.Cells(lin,7).Value2 = "Erro após salvar" if salvo else "Erro - verificar log"
                LOG.exception("Falha | %s",nome)
            finally:
                fechar(w)
                retry(lambda: setattr(excel,"Calculation",XL_MANUAL),"calculo manual")
                horario(ctrl,lin,9)
        retry(lambda: setattr(excel,"Calculation",XL_AUTO),"calculo automatico")
        retry(lambda: wc.Save(),"salvar controle")
        LOG.info("RESUMO | %s sucesso(s) | %s erro(s) | backup %s",len(ps)-erros,erros,backup)
        return 1 if erros else 0
    finally:
        for w in (wt,wb,wc):
            try:
                fechar(w)
            except Exception:
                LOG.exception("Falha ao fechar arquivo.")
        if excel is not None:
            try:
                if auto_fill is not None:
                    excel.AutoCorrect.AutoFillFormulasInLists = auto_fill
            except Exception:
                LOG.warning("Falha ao restaurar AutoFill.")
            try:
                retry(lambda: excel.Quit(),"encerrar Excel")
            except Exception:
                LOG.exception("Falha ao encerrar Excel.")
        pythoncom.CoUninitialize()

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pasta-base",default=str(Path.cwd()))
    parser.add_argument("--controle")
    parser.add_argument("--origem")
    parser.add_argument("--tendencia")
    parser.add_argument("--data-ref",help="AAAA-MM-DD; padrao D-1")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--simular",action="store_true",help="Padrao: nao salva")
    mode.add_argument("--aplicar",action="store_true",help="Autoriza gravacao apos preflight e backups")
    args = parser.parse_args()
    downloads = Path.home()/"Downloads"
    downloads.mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO,format="%(asctime)s | %(levelname)s | %(message)s",
                        handlers=[logging.StreamHandler(sys.stdout),
                                  logging.FileHandler(downloads/"Atualizar_Parceiros.log",encoding="utf-8")])
    try:
        return executar(args,downloads)
    except Exception:
        LOG.exception("Falha geral. Consulte o log.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
