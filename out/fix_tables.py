# -*- coding: utf-8 -*-
"""memo_builder가 만든 표의 w:tblW / w:tcW 중복을 제거하고 OOXML 스키마 순서로 재배치한다."""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TBLPR_ORDER = ["tblStyle","tblpPr","tblOverlap","bidiVisual","tblStyleRowBandSize",
               "tblStyleColBandSize","tblW","jc","tblCellSpacing","tblInd","tblBorders",
               "shd","tblLayout","tblCellMar","tblLook","tblCaption","tblDescription"]
TCPR_ORDER = ["cnfStyle","tcW","gridSpan","hMerge","vMerge","tcBorders","shd","noWrap",
              "tcMar","textDirection","tcFitText","vAlign","hideMark"]

def _insert_ordered(parent, el, order):
    name = el.tag.split("}")[1]
    idx = order.index(name)
    for child in parent:
        cname = child.tag.split("}")[1]
        if cname in order and order.index(cname) > idx:
            child.addprevious(el)
            return
    parent.append(el)

def _replace_single(parent, tag, order, **attrs):
    for old in parent.findall(qn("w:" + tag)):
        parent.remove(old)
    el = OxmlElement("w:" + tag)
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    _insert_ordered(parent, el, order)
    return el

def fix_tables(doc, table_width_dxa=9747):
    n_tbl = n_cell = 0
    for tbl in doc.element.body.iter(qn("w:tbl")):
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            continue
        # 표 전체 너비: 중복 제거 후 본문 폭으로 단일 지정
        _replace_single(tblPr, "tblW", TBLPR_ORDER, w=table_width_dxa, type="dxa")
        # 고정 레이아웃: 셀 너비 지정이 실제로 반영되도록
        _replace_single(tblPr, "tblLayout", TBLPR_ORDER, type="fixed")
        n_tbl += 1
        # 셀 너비: python-docx가 남긴 w=0 요소를 제거하고 dxa 값만 남김
        for tc in tbl.iter(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                continue
            tcWs = tcPr.findall(qn("w:tcW"))
            if not tcWs:
                continue
            keep = None
            for el in tcWs:
                if el.get(qn("w:type")) == "dxa" and el.get(qn("w:w")) not in (None, "0"):
                    keep = el
            if keep is None:
                keep = tcWs[-1]
            w, t = keep.get(qn("w:w")), keep.get(qn("w:type"))
            _replace_single(tcPr, "tcW", TCPR_ORDER, w=w, type=t)
            n_cell += 1
    return n_tbl, n_cell
