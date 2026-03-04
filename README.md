# Finch Automated Code Judger Overview

This directory contains the complete pipeline scripts from labeled data to evaluation content generation. The core flow is:

**JSONL task set → organize model outputs + source/reference files → preprocess files → generate evaluation `content_parts` → GPT Judge scoring.**

The JSONL task set is already included in the **huggingface-Finch** dataset. You can also organize your own JSONL task set following the same conventions. JSONL task set is agreed to be in the dataset root directory.

---

## Quick Start (Finch / FinWorkBench)

This quick start walks you through: **download dataset → build eval set → generate content parts → run GPT Judge → get `results.xlsx`**.

---

### 0) Prerequisites

* Python 3.9+ recommended
* Install dependencies (adjust to your project setup):

```bash
pip install -r requirements.txt
```

If you don’t have a requirements file, you will at least need:

```bash
pip install pandas openpyxl pymupdf python-docx xlwings openai
```

**Notes**

* Excel preprocessing needs **Microsoft Excel installed** (via `xlwings`) on Windows for best stability.
* PDF preprocessing needs `PyMuPDF` (`fitz`).

---

### 1) Download the Finch dataset

```bash
git clone https://huggingface.co/datasets/FinWorkBench/Finch
```

This creates a local folder named `Finch/` containing the dataset JSONL and source/reference files.

---

### 2) Prepare your model outputs directory

You need a directory that contains your model outputs in the expected structure (typically one subfolder per model). For example:

```
YOUR_MODEL_OUTPUT/
  opus_4.5_output/
    0.xlsx
    ...
  gpt4o_output/
    0.xlsx
    ...
```
For the model’s text output, it can be saved into a `.txt` file. For example, the model output of Task 1 can be saved as `1.txt`.

If there are multiple files of the same type, name them as `id_1.*`, `id_2.*` or `id-1.*`, `id-2.*`. For example, if there are multiple PNG images, name them `id_1.png`, `id_2.png` or `id-1.png`, `id-2.png`.

---

### 3) Run the 3-step pipeline (organize → preprocess → build prompts)

```bash
python src/prompt_build_pipeline.py --dataset-dir Finch --output-dir "YOUR MODEL OUTPUT" --target-dir eval_set
```

What this does:

* Reads the Finch JSONL tasks
* Organizes files into: `eval_set/<model>/<task_id>/...`
* Runs preprocessors (PDF/Word/Excel/Markdown/Image) and writes `metadata.json`
* Builds evaluation inputs and generates:

  * `eval_set/<model>/content_parts.jsonl`

After this step, your directory should look like:

```
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

### 4) Run GPT Judge and generate Excel results

```bash
python src/call_gpt_judge.py eval_set -o results.xlsx
```

What happens:

* For each `eval_set/<model>/content_parts.jsonl`
* Calls your configured Azure OpenAI judge model
* Writes per-model results to:

```
eval_set/<model>/results.xlsx
```

So you will get something like:

```
eval_set/
  opus_4.5_output/
    results.xlsx
  gpt4o_output/
    results.xlsx
```

---

### 5) Common options (optional)

**Only evaluate some models**

```bash
python src/call_gpt_judge.py eval_set --models opus_4.5_output,gpt4o_output -o results.xlsx
```

**Re-run everything (don’t skip processed tasks)**

```bash
python src/call_gpt_judge.py eval_set -o results.xlsx --no-skip-processed
```

---



## Directory Structure at a Glance

* `organize_files.py`: Reads the JSONL task set (agreed to be in the dataset root directory) and organizes source/reference/model-output files into a unified directory structure by task ID.
* `preprocessor/`: Preprocesses files such as PDF/Word/Excel/Markdown/Image, and writes results into `metadata.json` under `preprocess_info`.
* `build_prompt/`: Generates `content_parts.jsonl` based on `metadata.json` plus preprocessing outputs.
* `call_gpt_judge.py`: Calls Azure OpenAI for evaluation and writes results into Excel.
* `prompt_build_pipeline.py`: A three-step pipeline script (paths must be checked; see below).

---

## Typical Workflow

1. Excel annotations → JSONL task set (if you want to build your own dataset)
2. Organize files using the JSONL task set (aggregate outputs/sources/references)
3. Preprocess (extract text/images/screenshots, etc.)
4. Generate evaluation content (`content_parts.jsonl`)
5. Optional: Run GPT Judge to produce a scoring Excel

Below are the key configurations and usage for each script.

---

## 1) `organize_files.py` (Organize Files)

Purpose: Reads JSONL and creates task directories at `target_dir/<model>/<id>/`, then copies:
source files, reference files, and model output files. Finally, it generates `metadata.json`.

Command line:

```bash
python src/organize_files.py --dataset-dir data/workflow --output-dir data/opus --target-dir data/eval_dataset_opus
```

Key arguments:

* `--dataset-dir`: Dataset directory containing JSONL (the script auto-detects `*.jsonl`).
* `--output-dir`: Model output directory (one subdirectory per model).
* `--target-dir`: The organized output directory (root used by preprocessing/prompt building).
* `--log-level`: `DEBUG/INFO/WARNING/ERROR`.

---

## 2) `preprocessor/` (Preprocessing)

Entry script: `preprocessor/preprocessor_main.py`

```bash
python src/preprocessor/preprocessor_main.py --root-dir data/eval_dataset_opus
```

Key arguments:

* `--root-dir`: The `target-dir` produced by `organize_files.py`.
* `--models`: Optional. Specify which model directories to process (space-separated).

Dependencies:

* PDF: `PyMuPDF`
* Word: `python-docx`
* Excel: `xlwings` (requires Microsoft Excel installed locally)

Key configuration:

* Text descriptions are defined in `preprocessor/preprocessor_base.py` via `PreprocessorConfig`.
* Preprocessing results are written to `metadata.json` under `preprocess_info`.
* Special-case logs: `preprocessing_special_cases.log`.

---

## 3) `build_prompt/` (Generate `content_parts`)

Entry module: `build_prompt/content_builder/content_builder.py`

```bash
set PYTHONPATH=src
python -m src.build_prompt.content_builder.content_builder data/eval_dataset_opus
```

Key configuration: `build_prompt/content_builder/config.py`

* Size limits: `MAX_IMAGES`, `MAX_TEXT_CHARS`
* Extension sets: `EXCEL_EXTENSIONS`, `IMAGE_EXTENSIONS`, `TEXT_EXTENSIONS`, etc.
* Cache directory name: `CACHE_DIR_NAME`
* Output filename: `OUTPUT_JSONL_NAME` (default: `content_parts.jsonl`)
* Caption templates: `Captions.*`

Outputs:

* Generates `content_parts.jsonl` under each model directory.
* Generates `_cache/` under each task directory (diff, snapshot, screenshots cache, etc.).

---

## 4) `call_gpt_judge.py` (GPT Judge)

Supports two input modes:

* Input a JSONL file (single Excel output)
* Input a root directory (one Excel output per model subdirectory)

Examples:

```bash
# Process a single JSONL
python src/call_gpt_judge.py data/eval_dataset_opus/model_a/content_parts.jsonl -o results.xlsx --api-key ... --azure-endpoint ...

# Process a root directory (one results Excel per model)
python src/call_gpt_judge.py data/eval_dataset_opus -o results.xlsx
```

Key configuration (in `APIConfig` inside the script):

* `AZURE_ENDPOINT`, `API_KEY`, `API_VERSION`
* `MODEL`: Azure deployment name
* `MAX_TOKENS`, `MAX_COMPLETION_TOKENS`, `TEMPERATURE`
* `MAX_RETRIES`, `RATE_LIMIT_DELAY`

CLI arguments can override these settings (e.g., `--api-key`, `--azure-endpoint`, `--model`).

---

## 5) `prompt_build_pipeline.py`

A three-step pipeline script:

**Organize files → Preprocess → Generate `content_parts`**

---

## Notes

* Excel-related functionality depends on `xlwings` and a local Microsoft Excel installation; Windows is typically more stable.
* Missing `preprocess_info` usually means dependencies are not installed or the file type is not included in the preprocessing chain.
* Prompt length management: In `src\build_prompt\content_builder\token_counter.py`, it is split into two parts—an image limit and a text character limit—and they are calculated separately. If the image limit exceeds the configured value, images are dropped from `content_parts` starting from the end. If the text character count exceeds the limit, text in `content_parts` is truncated from the end. The maximum number of images and the maximum text character count are configured in `src\build_prompt\content_builder\config.py`.

## Legacy code
The code used in the paper is an older version, located in the `previous` branch. Link: https://github.com/FinWorkBench/Finch/tree/previous

The new code has been optimized based on this foundation.
