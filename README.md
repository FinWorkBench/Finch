# FinWorkBench (Finch): Benchmarking Finance & Accounting Across Spreadsheet-Centric Enterprise Workflows

> 📊 End-to-end evaluation pipeline for **FinWorkBench (Finch)**: organize task files, preprocess multimodal artifacts, build judge-ready prompts, and score model outputs.

## 🔗 Resources
- **Dataset**: [FinWorkBench/Finch on Hugging Face](https://huggingface.co/datasets/FinWorkBench/Finch)
- **Paper**: [arXiv:2512.13168](https://arxiv.org/abs/2512.13168)

This repository contains the full pipeline from labeled JSONL tasks to GPT-judge-based evaluation.

Core flow:

`JSONL task set -> organize outputs/source/reference files -> preprocess files -> build prompts -> GPT Judge scoring`

The official JSONL task set is included in the Hugging Face dataset. You can also prepare your own JSONL task set using the same schema.

---

## 🚀 Quick Start (Finch / FinWorkBench)

This quick start walks through:
**download dataset -> preprocess eval set -> build prompts -> run GPT Judge -> generate `results.xlsx`**.

If you want to reuse previously processed results for GPT-5.1 Pro, Claude Sonnet 4.5, and Claude Opus 4.5, you can download them directly from https://drive.google.com/file/d/1GMJz-gO33a8w5rlZYVhXqjxhlizLEmOM/view?usp=drive_link and skip dataset preprocessing.

---

### 0) ✅ Prerequisites

- Python 3.9+
- Install dependencies:

```bash
pip install -r requirements.txt
```

If needed, install core packages manually:

```bash
pip install pandas openpyxl pymupdf python-docx xlwings openai
```

**Notes**
- Excel preprocessing relies on **Microsoft Excel** (through `xlwings`) and is most stable on Windows.
- PDF preprocessing relies on `PyMuPDF` (`fitz`).

---

### 1) 📥 Download the Finch dataset

```bash
git clone https://huggingface.co/datasets/FinWorkBench/Finch
```

This creates a local `Finch/` directory containing JSONL tasks and source/reference files.

---

### 2) 🗂 Prepare your model output directory

Expected structure (example):

```text
YOUR_MODEL_OUTPUT/
  opus_4.5_output/
    0.xlsx
    ...
  gpt4o_output/
    0.xlsx
    ...
```

For text-only answers, save outputs as `.txt` files (for example, task 1 -> `1.txt`).

If a task has multiple files of the same type, use:
- `id_1.*`, `id_2.*`, ...
- or `id-1.*`, `id-2.*`, ...

---

### 3) 🏗 Run the 3-step pipeline (organize -> preprocess -> build prompts)

```bash
python src/prompt_build_pipeline.py --dataset-dir Finch --output-dir "YOUR_MODEL_OUTPUT" --target-dir eval_set
```

What this does:
- Reads Finch JSONL tasks
- Organizes files into `eval_set/<model>/<task_id>/...`
- Runs preprocessors (PDF/Word/Excel/Markdown/Image) and updates `metadata.json`
- Builds evaluation payloads and generates:
  - `eval_set/<model>/content_parts.jsonl`

Expected structure after this step:

```text
eval_set/
  opus_4.5_output/
    content_parts.jsonl
    0/
      metadata.json
      preprocessed/
      _cache/
      ...
  gpt4o_output/
    content_parts.jsonl
    ...
```

---

### 4) 🤖 Run GPT Judge and generate Excel results

```bash
python src/call_gpt_judge.py eval_set -o results.xlsx --api-key "<YOUR_KEY>" --azure-endpoint "<YOUR_ENDPOINT>" --api-version "<YYYY-MM-DD>" --model "<DEPLOYMENT_NAME>"
```

What happens:
- For each `eval_set/<model>/content_parts.jsonl`
- Calls your configured Azure OpenAI judge model
- Writes per-model results to:

```text
eval_set/<model>/results.xlsx
```

---

### 5) 🧩 Common options

**Evaluate selected models only**

```bash
python src/call_gpt_judge.py eval_set --models opus_4.5_output,gpt4o_output -o results.xlsx
```

**Re-run everything (do not skip processed tasks)**

```bash
python src/call_gpt_judge.py eval_set -o results.xlsx --no-skip-processed
```

---

## 🔄 Typical Workflow

1. Build or load a JSONL task set.
2. Organize source/reference/model-output files by task ID.
3. Preprocess files (extract text, screenshots, snapshots, etc.).
4. Build judge input payloads (`content_parts.jsonl`).
5. Run GPT Judge and export scoring results to Excel.

---

## 📁 Directory Structure at a Glance

- `src/organize_files.py`: reads JSONL tasks and organizes source/reference/model-output files into task directories.
- `src/preprocessor/`: preprocesses PDF/Word/Excel/Markdown/Image files and appends `preprocess_info` into `metadata.json`.
- `src/build_prompt/`: builds `content_parts.jsonl` from metadata + preprocessing outputs.
- `src/prompt_build_pipeline.py`: one-command pipeline for steps 1-3.
- `src/call_gpt_judge.py`: calls Azure OpenAI judge and writes Excel outputs.

---

## ⚙️ Script Details

### 1) `organize_files.py`

Purpose: create `target_dir/<model>/<task_id>/` and copy:
- source files
- reference files
- model output files

Then generate `metadata.json` per task.

Example:

```bash
python src/organize_files.py --dataset-dir data/workflow --output-dir data/opus --target-dir data/eval_dataset_opus
```

Key args:
- `--dataset-dir`: dataset root containing JSONL (`*.jsonl` auto-detected)
- `--output-dir`: model output root (subdirectory per model)
- `--target-dir`: organized output root
- `--log-level`: `DEBUG|INFO|WARNING|ERROR`

### 2) `preprocessor/preprocessor_main.py`

Example:

```bash
python src/preprocessor/preprocessor_main.py --root-dir data/eval_dataset_opus
```

Key args:
- `--root-dir`: pipeline root produced by `organize_files.py`
- `--models`: optional list of specific model folders

Dependencies by file type:
- PDF: `PyMuPDF`
- Word: `python-docx`
- Excel: `xlwings` + local Microsoft Excel

### 3) `build_prompt/content_builder/content_builder.py`

Example:

```bash
python -m src.build_prompt.content_builder.content_builder data/eval_dataset_opus
```

Key config: `src/build_prompt/content_builder/config.py`
- `MAX_IMAGES`, `MAX_TEXT_CHARS`
- extension sets (`EXCEL_EXTENSIONS`, `IMAGE_EXTENSIONS`, `TEXT_EXTENSIONS`, ...)
- cache settings and output filename

Outputs:
- `content_parts.jsonl` under each model directory
- `_cache/` under each task directory

### 4) `prompt_build_pipeline.py`

Runs:
1. organize files
2. preprocess files
3. generate `content_parts.jsonl`

### 5) `call_gpt_judge.py`

Supports two input modes:
- single JSONL input -> one Excel output
- root directory input -> one Excel per model directory

Examples:

```bash
# Single JSONL
python src/call_gpt_judge.py data/eval_dataset_opus/model_a/content_parts.jsonl -o results.xlsx --api-key "<YOUR_KEY>" --azure-endpoint "<YOUR_ENDPOINT>"

# Root directory (per-model Excel)
python src/call_gpt_judge.py data/eval_dataset_opus -o results.xlsx --api-key "<YOUR_KEY>" --azure-endpoint "<YOUR_ENDPOINT>"
```

---

## 📝 Notes

- In previous evaluations we used GPT-5-mini; stronger frontier models or multi-run voting can improve evaluation reliability.
- Excel-related functionality depends on `xlwings` and a local Microsoft Excel installation; Windows is usually more stable.
- Prompt length management is implemented in `src/build_prompt/content_builder/token_counter.py` and configured via `MAX_IMAGES` and `MAX_TEXT_CHARS` in `config.py`.
- If image count exceeds the limit, images are removed from the end; if text length exceeds the limit, text is truncated from the end.
- Missing `preprocess_info` usually indicates missing dependencies or unsupported/unprocessed file types.


---

## 🗂 Legacy Code

The code version used in the paper is maintained in the [`previous` branch](https://github.com/FinWorkBench/Finch/tree/previous).

This branch contains the newer and unified implementation for Modification, Generation, and QA.

---

## 📚 Citation

```bibtex
@article{dong2025finch,
  title={Finch: Benchmarking Finance \& Accounting across Spreadsheet-Centric Enterprise Workflows},
  author={Dong, Haoyu and Zhang, Pengkun and Gao, Yan and Dong, Xuanyu and Cheng, Yilin and Lu, Mingzhe and Yakefu, Adina and Zheng, Shuxin},
  journal={arXiv preprint arXiv:2512.13168},
  year={2025}
}
```
