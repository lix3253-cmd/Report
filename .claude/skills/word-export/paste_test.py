# -*- coding: utf-8 -*-
"""붙여넣기 시험: 기본 단락이 '줄 간격 고정 8pt'인 Word 파일에 결과 docx를 붙여넣은 상황을 만든다.

사용:
    python3 paste_test.py <결과.docx> <출력_폴더>
→ <출력_폴더>/paste_test.pdf, paste_test-1.png ... 생성 (LibreOffice, pdftoppm 필요)
결과 PNG를 열어 번호(1. / 1) / ① / a)), 표 행 높이, 글자 겹침이 없는지 확인한다.
Word의 "원본 서식 유지" 붙여넣기처럼 본문과 함께 '보고 N단계' 스타일, 번호 정의도 옮긴다.
"""
import copy
import os
import subprocess
import sys

import docx
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt


def main(src, outdir):
    os.makedirs(outdir, exist_ok=True)
    s = docx.Document(src)
    t = docx.Document()
    pf = t.styles["Normal"].paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(8)
    t.add_paragraph("대상 문서 본문 (줄 간격 고정 8pt)")
    for st in s.styles.element.findall(qn("w:style")):
        if st.get(qn("w:styleId"), "").startswith("Report"):
            t.styles.element.append(copy.deepcopy(st))
    tn, sn = t.part.numbering_part.element, s.part.numbering_part.element
    for a in sn.findall(qn("w:abstractNum")):
        if a.get(qn("w:abstractNumId")) == "90":
            first = tn.find(qn("w:num"))
            (first.addprevious(copy.deepcopy(a)) if first is not None else tn.append(copy.deepcopy(a)))
    for n in sn.findall(qn("w:num")):
        if n.get(qn("w:numId")) == "90":
            tn.append(copy.deepcopy(n))
    body = t.element.body
    for el in list(s.element.body):
        if not el.tag.endswith("sectPr"):
            body.insert(len(body) - 1, copy.deepcopy(el))
    path = os.path.join(outdir, "paste_test.docx")
    t.save(path)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, path],
                   check=True, capture_output=True)
    subprocess.run(["pdftoppm", "-r", "70", "-png", os.path.join(outdir, "paste_test.pdf"),
                    os.path.join(outdir, "paste_test")], check=True)
    print("rendered:", sorted(f for f in os.listdir(outdir) if f.endswith(".png")))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
