# gpt_judge_gpt_helpers.py
"""
Helpers that talk to Azure OpenAI:
- local_image_to_data_url
- call_gpt_judge (final evaluation)
- call_gpt_select_sheets
- call_gpt_pick_important_cells
- detect_important_cells_with_gpt
"""

import json
import sys
import traceback
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Set, Sequence
from mimetypes import guess_type
import base64
import pandas as pd

from openpyxl import load_workbook

from src.config import MODEL_NAME, client
from .gpt_judge_excel import (
    parse_addr,
    _index_to_col_letters,
    build_sheet_cells_text_for_region,# excel ， 
    expand_cell_range
)
from .gpt_judge_prompts import SHEET_SELECTION_PROMPT, HEADER_REGION_PROMPT


def local_image_to_data_url(image_path: Path) -> str:
    """Convert a local image into a data URL for Azure image_url."""
    mime_type, _ = guess_type(str(image_path))
    if mime_type is None:
        mime_type = "application/octet-stream"

    with image_path.open("rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")

    return f"data:{mime_type};base64,{b64}"


def call_gpt_judge(
    prompt: str,
    input_image: Optional[Path] = None,
    answer_image: Optional[Path] = None,
    output_image: Optional[Path] = None,
    extra_output_images: Optional[Sequence[Path]] = None,
) -> Dict[str, Any]:
    try:
        # 1. prompt 
        content_parts: List[Dict[str, Any]] = [
            {
                "type": "text",
                "text": prompt,
            }
        ]

        # 2. ：input / answer / output（modify ） 
        images: List[Tuple[str, Optional[Path]]] = [
            ("input", input_image),
            ("answer", answer_image),
            ("output", output_image),
        ]

        for name, img_path in images:
            if img_path is None:
                continue
            try:
                data_url = local_image_to_data_url(img_path)
            except Exception as e:
                print(f"[WARN] Failed to load {name} image {img_path}: {e}")
                continue

            
            if name == "input":
                caption = f"[Image: {img_path.name}] The image below is a screenshot of the original input workbook input.xlsx (possibly a concatenation of multiple sheets)."
            elif name == "answer":
                caption = f"[Image: {img_path.name}] The image below is a screenshot of the ground-truth workbook answer.xlsx (possibly a concatenation of multiple sheets)."
            elif name == "output":
                caption = f"[Image: {img_path.name}] The image below is a screenshot of the model output workbook output.xlsx (possibly a concatenation of multiple sheets)."
            else:
                caption = f"[Image: {img_path.name}] The image below is related to this task."


            
            content_parts.append(
                {
                    "type": "text",
                    "text": caption,
                }
            )
            content_parts.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": data_url,
                    },
                }
            )

        # 3. （modify “extra_output_images”， 
        # generate answer.png / answer_*.png / output.png / output_*.png / ） 
        if extra_output_images:
            for idx, img_path in enumerate(extra_output_images):
                if img_path is None:
                    continue
                try:
                    data_url = local_image_to_data_url(img_path)
                except Exception as e:
                    print(f"[WARN] Failed to load extra output image {img_path}: {e}")
                    continue

                fname = img_path.name.lower()

                # generate 
                if fname == "_input.png":
                    caption = f"[Image: {img_path.name}] The image below is a screenshot of the original input workbook input.xlsx."
                elif fname.startswith("input"):
                    caption = f"[Image: {img_path.name}] The image below is an additional screenshot related to the original input workbook input.xlsx."
                elif fname == "_answer.png":
                    caption = f"[Image: {img_path.name}] The image below is a screenshot of the ground-truth workbook answer.xlsx."
                elif fname.startswith("answer"):
                    caption = f"[Image: {img_path.name}] The image below is an additional screenshot related to the ground-truth workbook answer.xlsx."
                elif fname == "_output.png":
                    caption = f"[Image: {img_path.name}] The image below is a screenshot of the model output workbook output.xlsx."
                elif fname.startswith("output"):
                    caption = f"[Image: {img_path.name}] The image below is an additional screenshot related to the model output workbook output.xlsx."
                else:
                    caption = f"[Image: {img_path.name}] The image below is an additional image related to this task."


                content_parts.append(
                    {
                        "type": "text",
                        "text": caption,
                    }
                )
                content_parts.append(
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": data_url,
                        },
                    }
                )

        # 4. Azure OpenAI 
        rsp = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": content_parts,
                }
            ],
            #max_tokens=2048,
            max_completion_tokens=128000
        )

        raw = rsp.choices[0].message.content or ""
        data = json.loads(raw)

        score = data.get("score")
        detail = str(data.get("detailed_analysis", "") or "")

        if score not in (0, 1):
            return {
                "score": None,
                "detailed_analysis": detail,
                "llm_call_status": "error",
                "error_message": f"Invalid score in JSON: {score}",
            }

        return {
            "score": int(score),
            "detailed_analysis": detail,
            "llm_call_status": "success",
            "error_message": "",
        }

    except Exception as e:
        return {
            "score": None,
            "detailed_analysis": "",
            "llm_call_status": "error",
            "error_message": str(e),
        }


def call_gpt_select_sheets(
    query: str,
    sheet_names: Sequence[str],
    case_name: Optional[str] = None,
) -> Tuple[List[str], Optional[Dict[str, Any]]]:
    if not sheet_names:
        return [], None

    sheet_list_str = "\n".join(f"- {name}" for name in sheet_names)

    prompt = SHEET_SELECTION_PROMPT.format(
        query=query,
        sheet_name_list=sheet_list_str,
    )

    content_parts: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt}
    ]

    raw: Optional[str] = None
    data: Optional[Dict[str, Any]] = None
    error_text: Optional[str] = None
    traceback_text: Optional[str] = None

    try:
        rsp = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": content_parts,
                }
            ],
            #max_tokens=4096,
            max_completion_tokens=128000,
        )
        raw = rsp.choices[0].message.content or ""
        data = json.loads(raw)

    except Exception as e:
        error_text = repr(e)
        traceback_text = traceback.format_exc()
        print(f"[GPT-SHEET] call_gpt_select_sheets failed on case {case_name}: {error_text}")
        print("[GPT-SHEET] full traceback:")
        traceback.print_exc(file=sys.stdout)

    # debug （） 
    debug_info: Optional[Dict[str, Any]] = {
        "case": case_name,
        "sheet_names": list(sheet_names),
        "query": query,
        "prompt": prompt,
        "raw_text": raw,
        "raw_json": data,
        "error": error_text,
        "traceback": traceback_text,
    }

    if data is None:
        return [], debug_info

    # important_sheets 
    important_sheets_raw = data.get("important_sheets") or data.get("sheets") or []
    selected: List[str] = []
    if isinstance(important_sheets_raw, list):
        # sheet_names ， GPT 
        sheet_set = set(sheet_names)
        for name in important_sheets_raw:
            if isinstance(name, str):
                name = name.strip()
                if name in sheet_set:
                    selected.append(name)

    # ，“ sheet” 
    return selected, debug_info

def call_gpt_pick_important_cells(
    query: str,
    sheet_name: str,
    sheet_items: List[Dict[str, Any]],
    sheet_image: Optional[Path] = None,
    max_sheet_chars: int = 8000,
    case_name: Optional[str] = None,   # case （folder_name） 
) -> Tuple[Set[str], Optional[Dict[str, Any]]]:
    print("[GPT-HEADER] call_gpt_pick_important_cells ENTERED", sheet_name, case_name)

    # ====== sheet & prompt ====== 
    sheet_cells_text = build_sheet_cells_text_for_region(
        sheet_items,
        max_chars=max_sheet_chars,
    )

    prompt = HEADER_REGION_PROMPT.format(
        query=query,
        sheet_name=sheet_name,
        sheet_cells_text=sheet_cells_text,
    )

    content_parts: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt}
    ]

    if sheet_image is not None:
        try:
            data_url = local_image_to_data_url(sheet_image)
            content_parts.append(
                {
                    "type": "image_url",
                    "image_url": {"url": data_url},
                }
            )
        except Exception as e:
            print(f"[WARN] Failed to load sheet image for {sheet_name}: {e}")

    # ====== jsonl ====== 
    raw: Optional[str] = None
    data: Optional[Dict[str, Any]] = None
    error_text: Optional[str] = None
    traceback_text: Optional[str] = None

    # ====== Azure OpenAI（， finally ） ====== 
    try:
        rsp = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": content_parts,
                }
            ],
            # max_tokens， SDK 
            #max_tokens=4096,
            max_completion_tokens=128000
        )
        print(rsp.choices[0].message.content)
        raw = rsp.choices[0].message.content or ""
        data = json.loads(raw)

    except Exception as e:
        
        error_text = repr(e)
        traceback_text = traceback.format_exc()
        print(f"[WARN] call_gpt_pick_important_cells failed on sheet {sheet_name}: {error_text}")
        print("[WARN] exception type:", type(e))
        print("[WARN] full traceback:")
        traceback.print_exc(file=sys.stdout)

    # ====== ：， jsonl ====== 
    try:
        ctx_dir = Path.cwd() / "llm_judge_ctx"
        ctx_dir.mkdir(parents=True, exist_ok=True)
        api_raw_path = ctx_dir / "gpt_header_api_raw.jsonl"

        record = {
            "case": case_name,
            "sheet_name": sheet_name,
            "query": query,
            "prompt": prompt,
            "has_image": sheet_image is not None,
            "image_path": str(sheet_image) if sheet_image is not None else None,
            "raw_text": raw,        # ， None 
            "raw_json": data,       # JSON， None 
            "error": error_text,    # None， repr(e) 
            "traceback": traceback_text,  # None， traceback 
        }

        with api_raw_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")

        print(f"[GPT-HEADER] logged header API call to: {api_raw_path}")

    except Exception as e2:
        print(f"[GPT-HEADER][WARN] failed to append gpt_header_api_raw.jsonl: {e2}")

    # ====== API （data is None）， ====== 
    if data is None:
        return set(), None

    # ====== important_regions -> ====== 
    cells: Set[str] = set()
    regions = data.get("important_regions") or data.get("regions") or []
    if isinstance(regions, list):
        for region in regions:
            if isinstance(region, dict):
                rtype = str(region.get("type", "range")).lower()
                if rtype in ("range", "cell_range"):
                    start = region.get("start") or region.get("from")
                    end = region.get("end") or region.get("to")
                    if isinstance(start, str) and isinstance(end, str):
                        cells |= expand_cell_range(start.strip(), end.strip())
                elif rtype in ("cells", "cell_list"):
                    cell_list = region.get("cells") or region.get("addresses") or []
                    if isinstance(cell_list, list):
                        for addr in cell_list:
                            if not isinstance(addr, str):
                                continue
                            addr = addr.strip()
                            if parse_addr(addr):
                                cells.add(addr)
            elif isinstance(region, str):
                s = region.strip()
                if ":" in s:
                    start, end = s.split(":", 1)
                    cells |= expand_cell_range(start.strip(), end.strip())
                else:
                    if parse_addr(s):
                        cells.add(s)

    debug_info: Dict[str, Any] = {
        "raw_text": raw,
        "parsed": data,
    }
    return cells, debug_info

def build_input_sheet_image_map(case_dir: Path, input_workbook: Path) -> Dict[str, Path]:
    mapping: Dict[str, Path] = {}
    if not input_workbook.is_file():
        return mapping

    try:
        wb = load_workbook(filename=str(input_workbook), data_only=False)
    except Exception as e:
        print(f"[GPT-HEADER] Failed to open workbook {input_workbook} for sheet-image mapping: {e}")
        return mapping

    stem = input_workbook.stem  # "input" 
    for idx, ws in enumerate(wb.worksheets):
        img_name = f"_{stem}_{idx}.png"
        img_path = case_dir / img_name
        if img_path.is_file():
            mapping[ws.title] = img_path

    return mapping

def find_sheet_image(case_dir: Path, sheet_name: str) -> Optional[Path]:
    sheet_lower = sheet_name.lower()
    exts = (".png", ".jpg", ".jpeg", ".bmp")
    candidates: List[Path] = []

    for p in case_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in exts:
            continue
        stem_lower = p.stem.lower()
        if stem_lower in ("_input", "_answer", "_output"):
            continue
        if sheet_lower in stem_lower:
            candidates.append(p)

    if not candidates:
        return None

    
    candidates.sort(key=lambda x: len(x.name))
    return candidates[0]

def detect_important_cells_with_gpt(
    case_dir: Path,
    query_text: str,
    before_sheets: Dict[str, List[Dict[str, Any]]],
    target_sheets: Sequence[str],
    input_workbook: Optional[Path] = None,   
) -> Tuple[Dict[str, Set[str]], Dict[str, Any]]:
    important_cells: Dict[str, Set[str]] = {}
    debug_info: Dict[str, Any] = {}

    # ★ ： sheet_name -> _input_<idx>.png 
    sheet_image_map: Dict[str, Path] = {}
    if input_workbook is not None:
        try:
            sheet_image_map = build_input_sheet_image_map(case_dir, input_workbook)
            if sheet_image_map:
                print(f"[GPT-HEADER] Built sheet-image map from {input_workbook.name}: "
                      f"{', '.join(f'{k} -> {v.name}' for k, v in sheet_image_map.items())}")
        except Exception as e:
            print(f"[GPT-HEADER][WARN] failed to build sheet-image map: {e}")
            sheet_image_map = {}

    for sheet_name in target_sheets:
        sheet_items = before_sheets.get(sheet_name, []) or []
        if not sheet_items:
            continue

        # ★ mapping 
        sheet_image = sheet_image_map.get(sheet_name)
        if sheet_image is None:
            # ：“ sheet ” 
            sheet_image = find_sheet_image(case_dir, sheet_name)

        print(f"[GPT-HEADER] Detecting important cells for sheet '{sheet_name}' "
              f"with image: {sheet_image.name if sheet_image else 'None'} ...")

        cells, dbg = call_gpt_pick_important_cells(
            query=query_text,
            sheet_name=sheet_name,
            sheet_items=sheet_items,
            sheet_image=sheet_image,
            case_name=case_dir.name,   # ⭐ ： case 
        )


        if dbg is not None:
            debug_info[sheet_name] = dbg

        if cells:
            important_cells[sheet_name] = cells
            print(f"[GPT-HEADER] sheet '{sheet_name}' -> {len(cells)} important cells")
        else:
            print(f"[GPT-HEADER] sheet '{sheet_name}' -> no important cells detected")

    return important_cells, debug_info

def call_gpt_select_sheets(
    query: str,
    sheet_names: Sequence[str],
    case_name: Optional[str] = None,
) -> Tuple[List[str], Optional[Dict[str, Any]]]:
    if not sheet_names:
        return [], None

    sheet_list_str = "\n".join(f"- {name}" for name in sheet_names)

    prompt = SHEET_SELECTION_PROMPT.format(
        query=query,
        sheet_name_list=sheet_list_str,
    )

    content_parts: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt}
    ]

    raw: Optional[str] = None
    data: Optional[Dict[str, Any]] = None
    error_text: Optional[str] = None
    traceback_text: Optional[str] = None

    try:
        rsp = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": content_parts,
                }
            ],
            #max_tokens=4096,
            max_completion_tokens=128000,
        )
        raw = rsp.choices[0].message.content or ""
        data = json.loads(raw)

    except Exception as e:
        error_text = repr(e)
        traceback_text = traceback.format_exc()
        print(f"[GPT-SHEET] call_gpt_select_sheets failed on case {case_name}: {error_text}")
        print("[GPT-SHEET] full traceback:")
        traceback.print_exc(file=sys.stdout)

    # debug （） 
    debug_info: Optional[Dict[str, Any]] = {
        "case": case_name,
        "sheet_names": list(sheet_names),
        "query": query,
        "prompt": prompt,
        "raw_text": raw,
        "raw_json": data,
        "error": error_text,
        "traceback": traceback_text,
    }

    if data is None:
        return [], debug_info

    # important_sheets 
    important_sheets_raw = data.get("important_sheets") or data.get("sheets") or []
    selected: List[str] = []
    if isinstance(important_sheets_raw, list):
        # sheet_names ， GPT 
        sheet_set = set(sheet_names)
        for name in important_sheets_raw:
            if isinstance(name, str):
                name = name.strip()
                if name in sheet_set:
                    selected.append(name)

    # ，“ sheet” 
    return selected, debug_info

def ensure_output_xlsx(case_dir: Path) -> Path:
    output_xlsx = case_dir / "output.xlsx"
    if output_xlsx.is_file():
        return output_xlsx

    output_csv = case_dir / "output.csv"
    if output_csv.is_file():
        print(f"[INFO] Converting {output_csv.name} -> {output_xlsx.name} ...")
        try:
            # UTF-8 
            df = pd.read_csv(output_csv, encoding="utf-8", engine="python")
        except UnicodeDecodeError:
            # ：cp932 / shift-jis 
            df = pd.read_csv(output_csv, encoding="cp932", engine="python")

        with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

        print(f"[INFO] Saved converted Excel: {output_xlsx}")
        return output_xlsx

    # ，“” output.xlsx， Missing files 
    return output_xlsx
