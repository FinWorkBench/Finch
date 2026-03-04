#!/usr/bin/env python
# -*- coding: utf-8 -*-


import re
import json
import shutil
from pathlib import Path

import pandas as pd

# python-docx， docx 
try:
    from docx import Document
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
except ImportError:
    Document = None
    RT = None

# PyMuPDF， pdf -> png 
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


DATA_DIR = Path("data")

MODEL_OUTPUT_DIR = DATA_DIR / ""

ANNOTATION_XLSX = DATA_DIR / ""
SET_OUTPUT_ROOT = DATA_DIR / ""


# ：DATA_DIR / "library" 
LIBRARY_PATH = DATA_DIR / ""

# Excel （，） 
ID_COL = ""
SRC_COL = ""
# OTHER_SRC_COL = ""  
ANS_COL = ""
DESC_COL = ""
LIMIT_COL = ""
TASK_CLASS_COL = ""  


def parse_multi_value(cell_value):
    if cell_value is None or (isinstance(cell_value, float) and pd.isna(cell_value)):
        return []
    s = str(cell_value).strip()
    if not s:
        return []
    parts = re.split(r"[;；]", s)
    return [p.strip() for p in parts if p.strip()]


def build_file_index(root: Path):
    index = {}
    if not root.exists():
        return index
    for p in root.rglob("*"):
        if p.is_file():
            key = p.name.lower()
            index.setdefault(key, []).append(p)
    return index


def resolve_file(name: str, index: dict):
    # ， "[foo.png]" -> "foo.png" 
    
    name = name.strip()
    if name.startswith("[") and name.endswith("]") and len(name) > 2:
        name = name[1:-1].strip()

    key = Path(name).name.lower()
    candidates = index.get(key, [])
    if not candidates:
        print(f"[WARN] File not found: {name}")
        return None
    if len(candidates) > 1:
        print(f"[WARN] Multiple matches for {name}; using first: {candidates[0]}")
    return candidates[0]



def resolve_library_file(name: str, library_index: dict):
    name = name.strip()
    if name.startswith("[") and name.endswith("]") and len(name) > 2:
        name = name[1:-1].strip()
    if not name:
        return None
    key = Path(name).name.lower()
    candidates = library_index.get(key, [])
    if not candidates:
        return None
    if len(candidates) > 1:
        print(f"[WARN] Multiple matches for {name}; using first: {candidates[0]}")
    return candidates[0]


def extract_id_from_filename(filename: str) -> str | None:
    stem = Path(filename).stem
    m = re.search(r"(\d+)", stem)
    if not m:
        return None
    return m.group(1)


def infer_model_name_for_root_file(file_path: Path, model_output_dir: Path) -> str | None:
    """
    Infer model name for files directly under MODEL_OUTPUT_DIR.

    Legacy naming:
    - gptpro123.xlsx
    - claudesonnet456.txt

    New naming:
    - 123.xlsx
    - 140-1.csv
    In this case, use the directory name as model name (e.g. gptpro5.1).
    """
    lower_name = file_path.name.lower()
    for legacy_prefix in ("gptpro", "claudesonnet"):
        if lower_name.startswith(legacy_prefix):
            return legacy_prefix

    if re.match(r"^\d", file_path.stem):
        return model_output_dir.name

    return None


# ========== .pdf / .jsx -> .png ========== 


def convert_pdf_to_png(
    src_path: Path,
    dst_png: Path,
    special_cases: list[str],
    model_name: str,
    id_str: str,
    kind: str,
):
    if fitz is None:
        print(f"[WARN] PyMuPDF is not installed; cannot convert PDF to PNG: {src_path}")
        special_cases.append(
            f"[NO_PYMUPDF]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={src_path.name}"
        )
        return

    try:
        doc = fitz.open(src_path)
        page_count = doc.page_count
        if page_count == 0:
            print(f"[WARN] PDF has no pages: {src_path}")
            special_cases.append(
                f"[EMPTY_PDF]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={src_path.name}"
            )
            return

        # dst_png 
        # dst_png = eval_dir/input.png -> stem = "input" 
        stem = dst_png.stem
        out_dir = dst_png.parent
        out_dir.mkdir(parents=True, exist_ok=True)

        for i in range(page_count):
            page = doc.load_page(i)
            pix = page.get_pixmap()
            out_name = f"{stem}_page{i+1}.png"
            out_path = out_dir / out_name
            pix.save(str(out_path))
            print(f"[OK] {model_name} / {id_str}: {kind} PDF {src_path} page {i+1} -> PNG {out_path}")

    except Exception as e:
        print(f"[WARN] PDF-to-PNG conversion failed: {src_path}, error: {e}")
        special_cases.append(
            f"[PDF_CONVERT_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={src_path.name}\terror={e}"
        )




def convert_file_to_png_if_needed(
    dst_path: Path,
    special_cases: list[str],
    model_name: str,
    id_str: str,
    kind: str,
):
    suffix = dst_path.suffix.lower()

    # 1) PDF： PNG 
    if suffix == ".pdf":
        dst_png = dst_path.with_suffix(".png")
        convert_pdf_to_png(dst_path, dst_png, special_cases, model_name, id_str, kind)
        return

    # 2) ： 
    excel_exts = {".xlsx", ".xls", ".xlsm"}
    image_exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg"}
    rich_exts = {".docx", ".md", ".markdown"}
    text_exts = {".txt"}
    
    if suffix in excel_exts | image_exts | rich_exts | text_exts:
        
        return

    # 3) ： .txt 
    txt_path = dst_path.with_suffix(".txt")
    try:
        try:
            content = dst_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = dst_path.read_text(encoding="utf-8", errors="ignore")

        txt_path.write_text(content, encoding="utf-8")
        print(f"[OK] {model_name} / {id_str}: {kind} {dst_path.name} exported as text {txt_path.name}")
    except Exception as e:
        print(f"[WARN] Failed to export {dst_path} as text: {e}")
        special_cases.append(
            f"[GENERIC_TEXT_EXPORT_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={dst_path.name}\terror={e}"
        )



# ========== （markdown / docx） ========== 

def process_markdown_file(
    md_path: Path,
    out_dir: Path,
    library_index: dict,
    special_cases: list[str],
    model_name: str,
    id_str: str,
    kind: str,
    text_index_start: int = 0,
    pic_index_start: int = 0,
) -> tuple[int, int]:
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        content = md_path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        content = md_path.read_text(encoding="utf-8", errors="ignore")

    img_pattern = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
    pos = 0
    text_idx = text_index_start
    pic_idx = pic_index_start

    for m in img_pattern.finditer(content):
        # 1. 
        before = content[pos:m.start()]
        if before:
            text_idx += 1
            text_path = out_dir / f"text{text_idx}.txt"
            with text_path.open("w", encoding="utf-8") as f:
                f.write(before)
            print(f"[OK] {model_name} / {id_str}: {kind} markdown -> {text_path}")

        # 2. 
        img_path_raw = m.group(1).strip()
        img_fp = resolve_library_file(img_path_raw, library_index)
        if img_fp:
            pic_idx += 1
            ext = img_fp.suffix or ".png"
            pic_path = out_dir / f"pic{pic_idx}{ext}"
            shutil.copy2(img_fp, pic_path)
            print(f"[OK] {model_name} / {id_str}: {kind} markdown {img_fp} -> {pic_path}")
        else:
            
            text_idx += 1
            text_path = out_dir / f"text{text_idx}.txt"
            snippet = content[m.start(): m.end()]
            with text_path.open("w", encoding="utf-8") as f:
                f.write(snippet)
            special_cases.append(
                f"[MISSING_MD_IMAGE]\tmodel={model_name}\tid={id_str}\tkind={kind}\tref={img_path_raw}"
            )
            print(
                f"[WARN] {model_name} / {id_str}: markdown image {img_path_raw} was not found in the library; saved as text {text_path}"
            )

        pos = m.end()

    # 3. 
    tail = content[pos:]
    if tail:
        text_idx += 1
        text_path = out_dir / f"text{text_idx}.txt"
        with text_path.open("w", encoding="utf-8") as f:
            f.write(tail)
        print(f"[OK] {model_name} / {id_str}: {kind} markdown -> {text_path}")

    return text_idx, pic_idx





def process_docx_file(
    docx_path: Path,
    out_dir: Path,
    special_cases: list[str],
    model_name: str,
    id_str: str,
    kind: str,
    text_index_start: int = 0,
    pic_index_start: int = 0,
) -> tuple[int, int]:
    out_dir.mkdir(parents=True, exist_ok=True)

    if Document is None or RT is None:
        print(f"[WARN] python-docx is not installed; cannot process docx: {docx_path}")
        special_cases.append(
            f"[NO_PYTHON_DOCX]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={docx_path.name}"
        )
        return text_index_start, pic_index_start

    doc = Document(docx_path)

    text_idx = text_index_start
    pic_idx = pic_index_start

    # 1. ： textN.txt 
    texts = []
    for para in doc.paragraphs:
        if para.text:
            texts.append(para.text)
    if texts:
        text_idx += 1
        text_path = out_dir / f"text{text_idx}.txt"
        with text_path.open("w", encoding="utf-8") as f:
            f.write("\n".join(texts))
        print(f"[OK] {model_name} / {id_str}: {kind} docx -> {text_path}")

    # 2. ： relationship part 
    for rel in doc.part.rels.values():
        if rel.reltype != RT.IMAGE:
            continue

        img_part = getattr(rel, "target_part", None)
        if img_part is None:
            t = getattr(rel, "target", None)
            if t is not None and hasattr(t, "partname"):
                img_part = t

        if img_part is None:
            print(
                f"[WARN] {model_name} / {id_str}: could not resolve image part from relationship "
                f"(rel={rel}); skipping this image."
            )
            special_cases.append(
                f"[DOCX_REL_NO_TARGET]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={docx_path.name}"
            )
            continue

        try:
            pic_idx += 1
            partname = str(img_part.partname)
            ext = Path(partname).suffix or ".png"
            pic_path = out_dir / f"pic{pic_idx}{ext}"
            with pic_path.open("wb") as f:
                f.write(img_part.blob)
            print(f"[OK] {model_name} / {id_str}: {kind} docx -> {pic_path}")
        except Exception as e:
            print(f"[WARN] Failed to process docx image: {docx_path}, part={img_part}, error: {e}")
            special_cases.append(
                f"[DOCX_IMAGE_EXTRACT_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={docx_path.name}\terror={e}"
            )

    return text_idx, pic_idx


def get_next_text_pic_index(out_dir: Path) -> tuple[int, int]:
    max_text = 0
    max_pic = 0

    if not out_dir.exists():
        return 0, 0

    for p in out_dir.iterdir():
        if not p.is_file():
            continue

        m_txt = re.match(r"text(\d+)\.txt$", p.name, re.IGNORECASE)
        if m_txt:
            i = int(m_txt.group(1))
            if i > max_text:
                max_text = i

        m_pic = re.match(r"pic(\d+)\.", p.name, re.IGNORECASE)
        if m_pic:
            i = int(m_pic.group(1))
            if i > max_pic:
                max_pic = i

    return max_text, max_pic



def process_rich_file_if_needed(
    dst_path: Path,
    id_dir: Path,
    kind: str,
    library_index: dict,
    special_cases: list[str],
    model_name: str,
    id_str: str,
):
    suffix = dst_path.suffix.lower()
    if suffix not in (".md", ".markdown", ".docx"):
        return

    out_dir = id_dir / kind
    # textN/picN， 
    text_start, pic_start = get_next_text_pic_index(out_dir)

    if suffix in (".md", ".markdown"):
        process_markdown_file(
            dst_path,
            out_dir,
            library_index,
            special_cases,
            model_name,
            id_str,
            kind,
            text_index_start=text_start,
            pic_index_start=pic_start,
        )
    elif suffix == ".docx":
        process_docx_file(
            dst_path,
            out_dir,
            special_cases,
            model_name,
            id_str,
            kind,
            text_index_start=text_start,
            pic_index_start=pic_start,
        )



def process_qa_answer_cell(
    ans_raw,
    id_dir: Path,
    library_index: dict,
    special_cases: list[str],
    model_name: str,
    id_str: str,
):
    folder = id_dir / "answer"
    folder.mkdir(parents=True, exist_ok=True)

    # （ NaN） 
    if ans_raw is None or (isinstance(ans_raw, float) and pd.isna(ans_raw)):
        s = ""
    else:
        s = str(ans_raw)

    pattern = re.compile(r"\[([^\]]+)\]")
    pos = 0
    text_idx = 0
    pic_idx = 0

    for m in pattern.finditer(s):
        # 1. 
        before = s[pos:m.start()]
        if before:
            text_idx += 1
            text_path = folder / f"text{text_idx}.txt"
            with text_path.open("w", encoding="utf-8") as f:
                f.write(before)
            print(f"[OK] {model_name} / {id_str}: QA -> {text_path}")

        inner = m.group(1).strip()
        img_fp = resolve_library_file(inner, library_index)

        if img_fp:
            # 2a. 
            pic_idx += 1
            ext = img_fp.suffix or ".png"
            pic_path = folder / f"pic{pic_idx}{ext}"
            shutil.copy2(img_fp, pic_path)
            print(f"[OK] {model_name} / {id_str}: QA {img_fp} -> {pic_path}")
        else:
            # 2b. ： [xxx] 
            text_idx += 1
            text_path = folder / f"text{text_idx}.txt"
            snippet = s[m.start(): m.end()]  
            with text_path.open("w", encoding="utf-8") as f:
                f.write(snippet)
            print(
                f"[INFO] {model_name} / {id_str}: no matching image found for QA segment {snippet!r}; saved as text -> {text_path}"
            )

        pos = m.end()

    # 3. 
    tail = s[pos:]
    if tail:
        text_idx += 1
        text_path = folder / f"text{text_idx}.txt"
        with text_path.open("w", encoding="utf-8") as f:
            f.write(tail)
        print(f"[OK] {model_name} / {id_str}: QA -> {text_path}")

    
    # full_txt = folder / "answer_raw.txt"
    # with full_txt.open("w", encoding="utf-8") as f:
    #     f.write(s)
    # print(f"[OK] {model_name} / {id_str}: QA -> {full_txt}") 


# ========== ==========


def main():
    if not ANNOTATION_XLSX.exists():
        raise FileNotFoundError(f"Annotation workbook not found: {ANNOTATION_XLSX}")

    if not MODEL_OUTPUT_DIR.exists():
        raise FileNotFoundError(f"Model output directory not found: {MODEL_OUTPUT_DIR}")

    # eval_dataset （） 
    SET_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    
    df = pd.read_excel(ANNOTATION_XLSX, dtype={ID_COL: str})
    df[ID_COL] = df[ID_COL].astype(str).str.strip()
    records = df.set_index(ID_COL).to_dict(orient="index")

    # data （） 
    print(f"[INFO] Scanning all files under {DATA_DIR} to build index. This may take a while...")
    file_index = build_file_index(DATA_DIR)
    print(f"[INFO] Index built: {len(file_index)} unique filenames.")

    # library_path 
    library_index = build_file_index(LIBRARY_PATH)
    
    
    if library_index:
        print(f"[INFO] Image library index built under {LIBRARY_PATH}: {len(library_index)} filenames.")
    else:
        print(f"[WARN] Library directory {LIBRARY_PATH} does not exist or is empty; images in QA/markdown may not be resolvable.")

    
    special_cases: list[str] = []

    # (model_name, id) 
    groups: dict[tuple[str, str], list[Path]] = {}

    skipped = 0

    # ---------- 1. files directly under MODEL_OUTPUT_DIR ----------
    # ：web_model_output/gptpro123.xlsx 
    for f in MODEL_OUTPUT_DIR.iterdir():
        if not f.is_file():
            continue
        model_name = infer_model_name_for_root_file(f, MODEL_OUTPUT_DIR)
        if not model_name:
            continue

        name = f.name
        id_str = extract_id_from_filename(name)
        if not id_str:
            print(f"[WARN] No numeric ID found in file {name}; skipped.")
            special_cases.append(
                f"[BAD_FILENAME]\tmodel={model_name}\tid=UNKNOWN\tfile={name}\treason=NO_NUMERIC_ID"
            )
            skipped += 1
            continue

        groups.setdefault((model_name, id_str), []).append(f)

    # ---------- 2. api model ---------- 
    for model_dir in MODEL_OUTPUT_DIR.iterdir():
        if not model_dir.is_dir():
            continue

        model_name = model_dir.name  # finch_anthropic_claude-sonnet-4.5 

        for f in model_dir.iterdir():
            if not f.is_file():
                continue

            name = f.name
            id_str = extract_id_from_filename(name)
            if not id_str:
                print(f"[WARN] No numeric ID found in file {name}; skipped.")
                special_cases.append(
                    f"[BAD_FILENAME]\tmodel={model_name}\tid=UNKNOWN\tfile={name}\treason=NO_NUMERIC_ID"
                )
                skipped += 1
                continue

            groups.setdefault((model_name, id_str), []).append(f)

    processed = 0

    # (model_name, id) 
    for (model_name, id_str), file_list in groups.items():
        if id_str not in records:
            print(f"[WARN] ID {id_str} was not found in {ANNOTATION_XLSX}; skipping group {model_name}.")
            special_cases.append(
                f"[NO_ANNOTATION]\tmodel={model_name}\tid={id_str}\tfiles={[p.name for p in file_list]}"
            )
            skipped += 1
            continue

        row = records[id_str]
        src_list = parse_multi_value(row.get(SRC_COL))
        ans_raw = row.get(ANS_COL)
        desc = str(row.get(DESC_COL) or "").strip()
        limit_text = str(row.get(LIMIT_COL) or "").strip()
        task_class_raw = str(row.get(TASK_CLASS_COL) or "").strip()

        # task_classification， modify / generate / qa 
        task_class_norm = task_class_raw.lower()
        if task_class_norm in ("modify", "generate", "qa"):
            task_class = task_class_norm
        else:
            # ：， modify 
            task_class = "modify"
            special_cases.append(
                f"[UNKNOWN_TASK_CLASS]\tmodel={model_name}\tid={id_str}\tvalue={task_class_raw}"
            )

        # QA 
        ans_list = parse_multi_value(ans_raw) if task_class != "qa" else []

        # ：eval_dataset//ID 
        model_root = SET_OUTPUT_ROOT / model_name
        id_dir = model_root / id_str
        id_dir.mkdir(parents=True, exist_ok=True)

        # 1. SRC_COL -> input+ 
        # ：， input+，，。 
        if src_list:
            if len(src_list) > 1:
                special_cases.append(
                    f"[SRC_MULTI]\tmodel={model_name}\tid={id_str}\tcol={SRC_COL}\tvalues={';'.join(src_list)}"
                )
            for idx, src_name in enumerate(src_list, start=1):
                src_path = resolve_file(src_name, file_index)
                if not src_path:
                    special_cases.append(
                        f"[MISSING_SRC_FILE]\tmodel={model_name}\tid={id_str}\tcol={SRC_COL}\tfile={src_name}"
                    )
                    continue
                suffix = src_path.suffix
                if idx == 1:
                    dst_path = id_dir / f"input{suffix}"
                else:
                    
                    dst_path = id_dir / src_path.name
                shutil.copy2(src_path, dst_path)
                print(f"[OK] {model_name} / {id_str}: {src_path} -> {dst_path}")

                # input markdown / docx， textN / picN（ input ） 
                process_rich_file_if_needed(
                    dst_path,
                    id_dir,
                    kind="input",
                    library_index=library_index,
                    special_cases=special_cases,
                    model_name=model_name,
                    id_str=id_str,
                )


                # input pdf / jsx， png 
                convert_file_to_png_if_needed(
                    dst_path,
                    special_cases,
                    model_name,
                    id_str,
                    kind="input",
                )
        else:
            print(f"[WARN] {model_name} / ID={id_str}: '{SRC_COL}' is empty in the annotation row.")
            special_cases.append(
                f"[EMPTY_SRC_COL]\tmodel={model_name}\tid={id_str}\tcol={SRC_COL}"
            )

        # 2. ANS_COL 
        # - qa：ANS_COL []， textN.txt / picN.png 
        # - ：， answer+ 
        if task_class == "qa":
            # QA：“ + []” 
            process_qa_answer_cell(
                ans_raw,
                id_dir=id_dir,
                library_index=library_index,
                special_cases=special_cases,
                model_name=model_name,
                id_str=id_str,
            )
        else:
            # QA：“” 
            # answer+，（）。 
            if ans_list:
                if len(ans_list) > 1:
                    special_cases.append(
                        f"[ANS_MULTI]\tmodel={model_name}\tid={id_str}\tcol={ANS_COL}\tvalues={';'.join(ans_list)}"
                    )
                for idx, ans_name in enumerate(ans_list, start=1):
                    ans_path = resolve_file(ans_name, file_index)
                    if not ans_path:
                        special_cases.append(
                            f"[MISSING_ANS_FILE]\tmodel={model_name}\tid={id_str}\tcol={ANS_COL}\tfile={ans_name}"
                        )
                        continue
                    suffix = ans_path.suffix
                    if idx == 1:
                        dst_path = id_dir / f"answer{suffix}"
                    else:
                        
                        dst_path = id_dir / ans_path.name
                    shutil.copy2(ans_path, dst_path)
                    print(f"[OK] {model_name} / {id_str}: {ans_path} -> {dst_path}")

                    # answer markdown / docx，（ answer ） 
                    process_rich_file_if_needed(
                        dst_path,
                        id_dir,
                        kind="answer",
                        library_index=library_index,
                        special_cases=special_cases,
                        model_name=model_name,
                        id_str=id_str,
                    )


                    # answer pdf / jsx， png 
                    convert_file_to_png_if_needed(
                        dst_path,
                        special_cases,
                        model_name,
                        id_str,
                        kind="answer",
                    )
            else:
                print(f"[WARN] {model_name} / ID={id_str}: '{ANS_COL}' is empty in the annotation row.")
                special_cases.append(
                    f"[EMPTY_ANS_COL]\tmodel={model_name}\tid={id_str}\tcol={ANS_COL}"
                )

        # 3. 
        
        # - ：output+ 
        
        # * .xlsx output.xlsx 
        # * output_（1）+ 
        all_outputs = file_list
        xlsx_outputs = [p for p in all_outputs if p.suffix.lower() == ".xlsx"]
        other_outputs = [p for p in all_outputs if p.suffix.lower() != ".xlsx"]
        txt_outputs = [p for p in other_outputs if p.suffix.lower() == ".txt"]
        non_txt_other_outputs = [p for p in other_outputs if p.suffix.lower() != ".txt"]
        if not xlsx_outputs:
            print(f"[WARN] {model_name} / {id_str}: No xlsx model output file found.")
            special_cases.append(
                f"[NO_XLSX_OUTPUT]\tmodel={model_name}\tid={id_str}\tfiles={[p.name for p in all_outputs]}"
            )

        if all_outputs:
            if len(all_outputs) == 1:
                # ： output+ 
                only = all_outputs[0]
                suffix = only.suffix
                dst_output = id_dir / f"output{suffix}"
                if dst_output.exists():
                    print(f"[WARN] {model_name} / {id_str}: {dst_output.name} already exists and will be overwritten.")
                    special_cases.append(
                        f"[OUTPUT_OVERWRITE]\tmodel={model_name}\tid={id_str}\tfile={only.name}"
                    )
                shutil.copy2(only, dst_output)
                print(f"[OK] {model_name} / {id_str}: model output {only} -> {dst_output}")

                # markdown / docx， 
                process_rich_file_if_needed(
                    dst_output,
                    id_dir,
                    kind="output",
                    library_index=library_index,
                    special_cases=special_cases,
                    model_name=model_name,
                    id_str=id_str,
                )

                # pdf / jsx， png 
                convert_file_to_png_if_needed(
                    dst_output,
                    special_cases,
                    model_name,
                    id_str,
                    kind="output",
                )
            else:
                
                # 3.1 xlsx ： output.xlsx 
                main_output_path = None
                if xlsx_outputs:
                    main_xlsx = xlsx_outputs[0]
                    dst_output = id_dir / "output.xlsx"
                    main_output_path = dst_output
                    if dst_output.exists():
                        print(f"[WARN] {model_name} / {id_str}: output.xlsx already exists and will be overwritten.")
                        special_cases.append(
                            f"[OUTPUT_OVERWRITE]\tmodel={model_name}\tid={id_str}\tfile={main_xlsx.name}"
                        )
                    shutil.copy2(main_xlsx, dst_output)
                    print(f"[OK] {model_name} / {id_str}: model xlsx output {main_xlsx} -> {dst_output}")

                    if len(xlsx_outputs) > 1:
                        # xlsx， 
                        print(f"[WARN] {model_name} / {id_str}: multiple xlsx outputs found; non-primary files keep original names.")
                        special_cases.append(
                            f"[MULTI_XLSX_OUTPUT]\tmodel={model_name}\tid={id_str}\tfiles={[p.name for p in xlsx_outputs]}"
                        )
                        for extra in xlsx_outputs[1:]:
                            extra_dst = id_dir / extra.name
                            shutil.copy2(extra, extra_dst)
                            print(f"[OK] {model_name} / {id_str}: extra xlsx output {extra} -> {extra_dst}")

                # 3.2 txt ： output.txt 
                if txt_outputs:
                    primary_txt = txt_outputs[0]
                    dst_txt = id_dir / "output.txt"
                    if dst_txt.exists():
                        print(f"[WARN] {model_name} / {id_str}: output.txt already exists and will be overwritten.")
                        special_cases.append(
                            f"[OUTPUT_OVERWRITE]\tmodel={model_name}\tid={id_str}\tfile={primary_txt.name}"
                        )
                    shutil.copy2(primary_txt, dst_txt)
                    print(f"[OK] {model_name} / {id_str}: txt output {primary_txt} -> {dst_txt}")

                    # txt： output_{k}.txt， log 
                    if len(txt_outputs) > 1:
                        print(f"[WARN] {model_name} / {id_str}: multiple txt outputs found; extras are named output_{{k}}.txt")
                        special_cases.append(
                            f"[MULTI_TXT_OUTPUT]\tmodel={model_name}\tid={id_str}\tfiles={[p.name for p in txt_outputs]}"
                        )
                        for k, extra_txt in enumerate(txt_outputs[1:], start=1):
                            extra_dst = id_dir / f"output_{k}.txt"
                            if extra_dst.exists():
                                print(f"[WARN] {model_name} / {id_str}: {extra_dst.name} already exists and will be overwritten.")
                                special_cases.append(
                                    f"[OUTPUT_OVERWRITE]\tmodel={model_name}\tid={id_str}\tfile={extra_txt.name}"
                                )
                            shutil.copy2(extra_txt, extra_dst)
                            print(f"[OK] {model_name} / {id_str}: extra txt output {extra_txt} -> {extra_dst}")

                            # txt rich / png ，（ .txt skip） 
                            convert_file_to_png_if_needed(
                                extra_dst, special_cases, model_name, id_str, kind="output"
                            )

                # 3.3 xlsx、 txt ： output_+ 
                for idx, of in enumerate(non_txt_other_outputs, start=1):
                    suffix = of.suffix
                    dst = id_dir / f"output_{idx}{suffix}"
                    if dst.exists():
                        print(f"[WARN] {model_name} / {id_str}: {dst.name} already exists and will be overwritten.")
                        special_cases.append(
                            f"[OUTPUT_OVERWRITE]\tmodel={model_name}\tid={id_str}\tfile={of.name}"
                        )
                    shutil.copy2(of, dst)
                    print(f"[OK] {model_name} / {id_str}: other output file {of} -> {dst}")

                    process_rich_file_if_needed(
                        dst,
                        id_dir,
                        kind="output",
                        library_index=library_index,
                        special_cases=special_cases,
                        model_name=model_name,
                        id_str=id_str,
                    )
                    convert_file_to_png_if_needed(
                        dst,
                        special_cases,
                        model_name,
                        id_str,
                        kind="output",
                    )

        # 4. query.txt （ + ） 
        query_lines = []
        if desc:
            query_lines.append(desc)
        if limit_text:
            query_lines.append(limit_text)
        query_text = "\n\n".join(query_lines)

        query_file = id_dir / "query.txt"
        with query_file.open("w", encoding="utf-8") as qf:
            qf.write(query_text)
        print(f"[OK] {model_name} / {id_str}: query.txt")

        # 5. property.json（task_classification） 
        property_path = id_dir / "property.json"
        with property_path.open("w", encoding="utf-8") as pf:
            json.dump(
                {"task_classification": task_class},
                pf,
                ensure_ascii=False,
                indent=2,
            )
        print(f"[OK] {model_name} / {id_str}: property.json (task_classification={task_class})")

        processed += 1

    # post_process.log 
    log_path = SET_OUTPUT_ROOT / "post_process.log"
    with log_path.open("w", encoding="utf-8") as logf:
        for line in special_cases:
            logf.write(line + "\n")

    print(f"\n[SUMMARY] Processed {processed} groups (model, ID), skipped {skipped} groups/files.")
    print(f"[INFO] Special-case log written to: {log_path.resolve()}")
    print(f"Generated eval directory: {SET_OUTPUT_ROOT.resolve()}")

    print(library_index)

if __name__ == "__main__":
    main()
