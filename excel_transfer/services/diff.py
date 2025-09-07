# excel_transfer/services/diff.py
import os, datetime, hashlib, tempfile
from pathlib import Path
from typing import Dict, Any, List, Tuple
import xlwings as xw

from models.dto import DiffRequest, LogFn
from utils.excel import open_app, safe_kill, used_range_2d_values, normalize_2d

COLOR_HEADER = (217,225,242); COLOR_DIFF=(248,203,173); COLOR_ADD=(198,239,206); COLOR_DEL=(255,199,206); COLOR_CENTER=(255,230,153)

def _align(a: List[List[str]], b: List[List[str]]) -> Tuple[List[List[str]], List[List[str]]]:
    rows = max(len(a), len(b))
    cols = max(max((len(r) for r in a), default=0), max((len(r) for r in b), default=0))
    def pad(mat):
        out=[]; 
        for r in range(rows):
            row = mat[r] if r < len(mat) else []
            row = row + [""]*(cols-len(row)); out.append(row)
        return out
    return pad(a), pad(b)

def _a1(col_idx: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(col_idx)

def _diff_by_position(a_body, b_body, sheet_name):
    A,B = _align(a_body,b_body); diffs=[]
    for r in range(len(A)):
        for c in range(len(A[r])):
            if A[r][c]!=B[r][c]:
                addr = f"{sheet_name}!{_a1(c+1)}{r+2}"
                diffs.append((r+2,c+1,A[r][c],B[r][c],addr))
    return diffs

def _head_idx(header: List[str]): 
    m={}; 
    for i,h in enumerate(header):
        k=str(h).strip()
        if k and k not in m: m[k]=i
    return m

def _diff_by_keys(a_table, b_table, key_cols: List[str]):
    if not a_table or not b_table: 
        return {"added": [], "deleted": [], "changed": [], "columns": []}
    ah=[str(h) for h in a_table[0]]; bh=[str(h) for h in b_table[0]]
    cols = list(dict.fromkeys(ah+bh)) if ah!=bh else ah
    ai=_head_idx(ah); bi=_head_idx(bh)
    def rowd(row, idx):
        d={}; 
        for c in cols:
            j=idx.get(c); d[c]=row[j] if (j is not None and j < len(row)) else ""
        return d
    def key_of(d): return tuple(d.get(k,"") for k in key_cols)
    am={ key_of(d): d for d in (rowd(r,ai) for r in a_table[1:])}
    bm={ key_of(d): d for d in (rowd(r,bi) for r in b_table[1:])}
    added=[(k,d) for k,d in bm.items() if k not in am]
    deleted=[(k,d) for k,d in am.items() if k not in bm]
    changed=[]
    for k in set(am)&set(bm):
        da,db=am[k],bm[k]; diff_cols={}
        for c in cols:
            if da.get(c,"")!=db.get(c,""): diff_cols[c]=(da.get(c,""),db.get(c,""))
        if diff_cols: changed.append((k,diff_cols))
    return {"added":added,"deleted":deleted,"changed":changed,"columns":cols}

def _extract_context(vals: List[List[str]], r: int, c: int, radius: int) -> List[List[str]]:
    ctx=[]; r0=max(1, r-radius); r1=r+radius; c0=max(1, c-radius); c1=c+radius
    for rr in range(r0, r1+1):
        row=[]
        for cc in range(c0, c1+1):
            try: row.append(vals[rr-1][cc-1])
            except: row.append("")
        ctx.append(row)
    return ctx

def _paste_context_sheet(wb: xw.Book, name: str, ctx: List[List[str]], center_rc=(0,0)):
    sh = wb.sheets.add(name)
    if ctx:
        sh.range(1,1).value = ctx
        rr,cc=center_rc
        try: sh.range(1+rr, 1+cc).color = COLOR_CENTER
        except: pass
    sh.autofit()

def _fp_shape(s):
    try: typ = s.api.Type
    except: typ = "?"
    return (typ, s.name, int(s.left), int(s.top), int(s.width), int(s.height))

def _pic_hash(p, tmpdir) -> str:
    path = os.path.join(tmpdir, f"{p.name}.png")
    try:
        p.api.Export(path)
        with open(path, "rb") as f: return hashlib.md5(f.read()).hexdigest()
    except Exception:
        return ""

def _compare_shapes(file_a: str, file_b: str) -> List[List[Any]]:
    rows=[]
    tmp = tempfile.mkdtemp()
    app = open_app()
    try:
        maps={}
        for tag, path in [("A",file_a),("B",file_b)]:
            wb = app.books.open(path)
            shmap={}
            for sh in wb.sheets:
                ary=[]
                for s in sh.shapes: ary.append(("shape", _fp_shape(s), ""))  # 図形
                for p in sh.pictures: ary.append(("picture", _fp_shape(p), _pic_hash(p,tmp)))  # 画像
                shmap[sh.name]=ary
            maps[tag]=shmap
            wb.close()
        sheets = sorted(set(maps["A"].keys()) | set(maps["B"].keys()))
        for s in sheets:
            A = {(k,fp,hs) for (k,fp,hs) in maps["A"].get(s, [])}
            B = {(k,fp,hs) for (k,fp,hs) in maps["B"].get(s, [])}
            for (k,fp,hs) in sorted(B - A): rows.append([s,"ADDED",   k, fp[1],fp[2],fp[3],fp[4],hs])
            for (k,fp,hs) in sorted(A - B): rows.append([s,"DELETED", k, fp[1],fp[2],fp[3],fp[4],hs])
        return rows
    finally:
        safe_kill(app)

def _write_report(file_a: Path, file_b: Path, diffs: Dict[str, Dict[str, Any]], by_keys: bool,
                  include_context: bool, context_items: List[Tuple[str,int,int,List[List[str]],List[List[str]]]],
                  shapes_rows: List[List[Any]], out_path:str) -> str:
    app = open_app()
    try:
        wb = app.books.add()
        sh = wb.sheets[0]; sh.name="Summary"
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sh.range(1,1).value = [["Diff Report",""],["Generated",now],["File A",str(Path(file_a).resolve())],["File B",str(Path(file_b).resolve())],["Mode","Key-based" if by_keys else "Position-based"]]
        sh.range(1,1).expand().color = COLOR_HEADER
        row=7; sh.range(row,1).value=[["Sheet","Changed","Added","Deleted"]]; sh.range(row,1).expand().color=COLOR_HEADER; row+=1

        for sheet, det in diffs.items():
            if by_keys:
                ch, ad, de = det.get("changed",[]), det.get("added",[]), det.get("deleted",[])
                cols = det.get("columns",[])
                cws = wb.sheets.add(f"Changed_{sheet[:20]}"); cws.range(1,1).value=[["Key","Column","A","B"]]; cws.range(1,1).expand().color=COLOR_HEADER
                out=[]
                for ktuple, dcols in ch:
                    key=" | ".join(ktuple)
                    for col,(av,bv) in dcols.items(): out.append([key,col,av,bv])
                if out: cws.range(2,1).value=out; cws.range(2,1).expand().color=COLOR_DIFF
                cws.autofit()

                aws = wb.sheets.add(f"Added_{sheet[:21]}"); aws.range(1,1).value=[["Key"]+cols]; aws.range(1,1).expand().color=COLOR_HEADER
                out_a=[[" | ".join(k)]+[rowd.get(c,"") for c in cols] for k,rowd in det.get("added",[])]
                if out_a: aws.range(2,1).value=out_a; aws.range(2,1).expand().color=COLOR_ADD
                aws.autofit()

                dws = wb.sheets.add(f"Deleted_{sheet[:21]}"); dws.range(1,1).value=[["Key"]+cols]; dws.range(1,1).expand().color=COLOR_HEADER
                out_d=[[" | ".join(k)]+[rowd.get(c,"") for c in cols] for k,rowd in det.get("deleted",[])]
                if out_d: dws.range(2,1).value=out_d; dws.range(2,1).expand().color=COLOR_DEL
                dws.autofit()

                sh.range(row,1).value=[[sheet,len(ch),len(det.get("added",[])),len(det.get("deleted",[]))]]; row+=1

            else:
                cells = det.get("changed_cells", [])
                ws = wb.sheets.add(f"Diff_{sheet[:24]}")
                ws.range(1,1).value = [["Address","Row","Col","A","B","Open A","Open B"]]
                ws.range(1,1).expand().color = COLOR_HEADER
                r=[]
                for (rr,cc,av,bv,addr) in cells:
                    link_a = f'=HYPERLINK("[{file_a}]{sheet}!{_a1(cc)}{rr}","Open A")'
                    link_b = f'=HYPERLINK("[{file_b}]{sheet}!{_a1(cc)}{rr}","Open B")'
                    r.append([addr, rr, cc, av, bv, link_a, link_b])
                if r:
                    ws.range(2,1).value=r; ws.range(2,1).expand().color=COLOR_DIFF
                ws.autofit()
                sh.range(row,1).value=[[sheet,len(cells),0,0]]; row+=1

        if include_context and context_items:
            for i,(sheet, rr, cc, ctxA, ctxB) in enumerate(context_items, start=1):
                _paste_context_sheet(wb, f"CtxA_{sheet[:12]}_{i}", ctxA, center_rc=(2,2))
                _paste_context_sheet(wb, f"CtxB_{sheet[:12]}_{i}", ctxB, center_rc=(2,2))

        if shapes_rows:
            sps = wb.sheets.add("ShapesDiff")
            sps.range(1,1).value=[["Sheet","Change","Kind","Name","Left","Top","Width","Height","Hash"]]
            sps.range(1,1).expand().color=COLOR_HEADER
            sps.range(2,1).value=shapes_rows
            sps.autofit()

        wb.save(out_path); wb.close()
        return out_path
    finally:
        safe_kill(app)

def run_diff(req: DiffRequest, ctx, logger, append_log: LogFn) -> str:
    if not os.path.isfile(req.file_a): raise ValueError("File A が不正です。")
    if not os.path.isfile(req.file_b): raise ValueError("File B が不正です。")

    append_log("=== Diff開始 ===")
    app = open_app(); diffs_per_sheet: Dict[str, Dict[str, Any]] = {}; contexts=[]
    try:
        wb_a = app.books.open(req.file_a); wb_b = app.books.open(req.file_b)
        try:
            sa = [s.name for s in wb_a.sheets]; sb = [s.name for s in wb_b.sheets]
            all_sheets = sorted(set(sa)|set(sb))
            for s in all_sheets:
                sha = wb_a.sheets[s] if s in sa else None
                shb = wb_b.sheets[s] if s in sb else None
                a = normalize_2d(used_range_2d_values(sha, as_formula=req.compare_formula)) if sha else []
                b = normalize_2d(used_range_2d_values(shb, as_formula=req.compare_formula)) if shb else []

                if req.key_cols:
                    diffs_per_sheet[s] = _diff_by_keys(a,b,req.key_cols)
                else:
                    a_body = a[1:] if a else []; b_body = b[1:] if b else []
                    changed = _diff_by_position(a_body,b_body,s)
                    diffs_per_sheet[s] = {"changed_cells": changed}
                    if req.include_context and changed:
                        for (rr,cc,_,_,_) in changed[:req.max_context_items]:
                            ctxA = _extract_context(a, rr, cc, req.context_radius)
                            ctxB = _extract_context(b, rr, cc, req.context_radius)
                            contexts.append((s, rr, cc, ctxA, ctxB))

                append_log(f"比較: {s}")
        finally:
            wb_a.close(save=False); wb_b.close(save=False)
    finally:
        safe_kill(app)

    shapes_rows=[]
    if req.compare_shapes:
        append_log("図・画像の比較中…")
        shapes_rows = _compare_shapes(req.file_a, req.file_b)

    out = os.path.join(ctx.output_dir, "diff_report.xlsx")
    return _write_report(req.file_a, req.file_b, diffs_per_sheet, by_keys=bool(req.key_cols),
                         include_context=req.include_context, context_items=contexts,
                         shapes_rows=shapes_rows, out_path=out)
