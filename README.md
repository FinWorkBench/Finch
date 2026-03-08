# FinWorkBench (Finch): Benchmarking Finance & Accounting Across Spreadsheet-Centric Enterprise Workflows

[![HuggingFace Dataset](https://img.shields.io/badge/HuggingFace-Finch%20Dataset-yellow?logo=huggingface)](https://huggingface.co/datasets/FinWorkBench/Finch)
[![arXiv](https://img.shields.io/badge/arXiv-2512.13168-b31b1b?logo=arxiv)](https://arxiv.org/abs/2512.13168)

> End-to-end evaluation pipeline for **FinWorkBench (Finch)**: organize task files, preprocess multimodal artifacts, build judge-ready prompts, and score model outputs.

The evaluation pipeline follows four stages:

> **JSONL task set** &rarr; **Organize files** &rarr; **Preprocess & build prompts** &rarr; **GPT Judge scoring**

The JSONL task set is included in the [Finch HuggingFace dataset](https://huggingface.co/datasets/FinWorkBench/Finch). You can also create your own JSONL task set following the same conventions. The JSONL file should be placed in the dataset root directory.

---

## :open_file_folder: Repository Structure

```
src/
├── organize_files.py           # Organizes source/reference/model-output files by task ID
├── preprocessor/               # Preprocesses PDF, Word, Excel, Markdown, and image files
│   └── preprocessor_main.py    # Preprocessing entry point
├── build_prompt/               # Builds judge-ready prompts (content_parts.jsonl)
│   └── content_builder/
│       ├── content_builder.py  # Prompt builder entry point
│       └── config.py           # Size limits, extension sets, cache/output settings
├── prompt_build_pipeline.py    # End-to-end pipeline: organize -> preprocess -> build prompts
└── call_gpt_judge.py           # Calls Azure OpenAI judge model and writes scored results
```

---

## :rocket: Quick Start

This guide walks you through: **download dataset &rarr; preprocess &rarr; build prompts &rarr; run GPT Judge &rarr; get `results.xlsx`**.

> **Shortcut:** If you want to reuse our preprocessed results for GPT-5.1 Pro, Claude Sonnet 4.5, and Claude Opus 4.5, download them from [Google Drive](https://drive.google.com/file/d/1GMJz-gO33a8w5rlZYVhXqjxhlizLEmOM/view?usp=drive_link) and skip directly to (5) Run GPT Judge.

### 1) :gear: Prerequisites

- Python 3.9+ recommended
- Install dependencies:

```bash
pip install -r requirements.txt
```

If you do not have a requirements file, install at minimum:

```bash
pip install pandas openpyxl pymupdf python-docx xlwings openai
```

| Dependency | Required for |
|---|---|
| `xlwings` + Microsoft Excel | Excel preprocessing (Windows recommended for stability) |
| `PyMuPDF` (`fitz`) | PDF preprocessing |
| `python-docx` | Word document preprocessing |

### 2) :inbox_tray: Download the Finch dataset

```bash
git clone https://huggingface.co/datasets/FinWorkBench/Finch
```

This creates a local `Finch/` folder containing the JSONL task set and source/reference files.

### 3) :briefcase: Prepare your model outputs

Create a directory containing your model outputs, with one subfolder per model:

```text
YOUR_MODEL_OUTPUT/
├── opus_4.5_output/
│   ├── 0.xlsx
│   └── ...
└── gpt5.1_output/
    ├── 0.xlsx
    └── ...
```

**Naming conventions:**

| Scenario | Convention | Example |
|---|---|---|
| Text output for a task | `<task_id>.txt` | `1.txt` |
| Single file per task | `<task_id>.<ext>` | `5.xlsx` |
| Multiple files of the same type | `<task_id>_<n>.<ext>` or `<task_id>-<n>.<ext>` | `3_1.png`, `3_2.png` |

### 4) :hammer_and_wrench: Run the 3-step pipeline (organize, preprocess, build prompts)

```bash
python src/prompt_build_pipeline.py \
    --dataset-dir Finch \
    --output-dir "YOUR_MODEL_OUTPUT" \
    --target-dir eval_set
```

This command:

1. Reads the Finch JSONL tasks.
2. Organizes files into `eval_set/<model>/<task_id>/`.
3. Runs preprocessors (PDF, Word, Excel, Markdown, image) and writes `metadata.json`.
4. Builds prompts and generates `eval_set/<model>/content_parts.jsonl`.

After completion, the output directory looks like:

```text
eval_set/
├── opus_4.5_output/
│   ├── content_parts.jsonl
│   └── 0/
│       ├── metadata.json
│       ├── preprocessed/
│       └── _cache/
└── gpt5.1_output/
    ├── content_parts.jsonl
    └── ...
```

### 5) :bar_chart: Run GPT Judge

```bash
python src/call_gpt_judge.py eval_set -o results.xlsx \
    --api-key "<YOUR_KEY>" \
    --azure-endpoint "<YOUR_ENDPOINT>" \
    --api-version "<YYYY-MM-DD>" \
    --model "<DEPLOYMENT_NAME>"
```

This reads each `content_parts.jsonl`, calls the configured Azure OpenAI judge model, and writes a scored `results.xlsx` into each model subdirectory:

```text
eval_set/
├── opus_4.5_output/
│   └── results.xlsx
└── gpt5.1_output/
    └── results.xlsx
```

**Common options:**

```bash
# Evaluate specific models only
python src/call_gpt_judge.py eval_set --models opus_4.5_output,gpt5.1_output -o results.xlsx

# Re-run without skipping previously processed tasks
python src/call_gpt_judge.py eval_set -o results.xlsx --no-skip-processed
```

---

## :books: Detailed Script Reference

### `organize_files.py` -- Organize Files

Reads the JSONL task set and creates task directories at `target_dir/<model>/<task_id>/`, copying source files, reference files, and model output files. Generates `metadata.json` for each task.

```bash
python src/organize_files.py \
    --dataset-dir Finch \
    --output-dir YOUR_MODEL_OUTPUT \
    --target-dir eval_set
```

| Argument | Description |
|---|---|
| `--dataset-dir` | Dataset directory containing the JSONL file (auto-detected) |
| `--output-dir` | Model output directory (one subdirectory per model) |
| `--target-dir` | Organized output directory used by downstream pipeline steps |
| `--log-level` | Logging verbosity: `DEBUG`, `INFO`, `WARNING`, or `ERROR` |

### `preprocessor/` -- Preprocess Files

Processes PDF, Word, Excel, Markdown, and image files. Results are written into `metadata.json` under the `preprocess_info` field.

```bash
python src/preprocessor/preprocessor_main.py --root-dir eval_set
```

| Argument | Description |
|---|---|
| `--root-dir` | The `target-dir` produced by `organize_files.py` |
| `--models` | Optional; specify which model directories to process (space-separated) |

**Configuration:** Text descriptions are defined in `preprocessor/preprocessor_base.py` via `PreprocessorConfig`. Special-case logs are written to `preprocessing_special_cases.log`.

### `build_prompt/` -- Build Prompts

Generates `content_parts.jsonl` for each model directory, based on `metadata.json` and preprocessing outputs.

```bash
set PYTHONPATH=src
python -m src.build_prompt.content_builder.content_builder eval_set
```

**Configuration** (`build_prompt/content_builder/config.py`):

| Setting | Description |
|---|---|
| `MAX_IMAGES` | Maximum number of images per prompt |
| `MAX_TEXT_CHARS` | Maximum text character count per prompt |
| `EXCEL_EXTENSIONS`, `IMAGE_EXTENSIONS`, `TEXT_EXTENSIONS` | Recognized file extension sets |
| `CACHE_DIR_NAME` | Name of the cache directory (stores diffs, snapshots, screenshots) |
| `OUTPUT_JSONL_NAME` | Output filename (default: `content_parts.jsonl`) |
| `Captions.*` | Caption templates for generated content |

**Prompt length management:** Defined in `src/build_prompt/content_builder/token_counter.py`. Image count and text character count are tracked separately. If the image count exceeds `MAX_IMAGES`, images are dropped from the end. If the text character count exceeds `MAX_TEXT_CHARS`, text is truncated from the end. Both limits are configured in `config.py`.

### `call_gpt_judge.py` -- GPT Judge Scoring

Supports two input modes:

- **Single JSONL file** -- scores one model and writes a single `results.xlsx`.
- **Root directory** -- scores all models and writes one `results.xlsx` per model subdirectory.

```bash
# Single model
python src/call_gpt_judge.py eval_set/opus_4.5_output/content_parts.jsonl \
    -o results.xlsx --api-key "<YOUR_KEY>" --azure-endpoint "<YOUR_ENDPOINT>"

# All models under a root directory
python src/call_gpt_judge.py eval_set -o results.xlsx \
    --api-key "<YOUR_KEY>" --azure-endpoint "<YOUR_ENDPOINT>"
```

**API configuration** (set in `APIConfig` inside the script, overridable via CLI):

| Setting | CLI Override |
|---|---|
| `AZURE_ENDPOINT` | `--azure-endpoint` |
| `API_KEY` | `--api-key` |
| `API_VERSION` | `--api-version` |
| `MODEL` | `--model` |
| `MAX_TOKENS`, `MAX_COMPLETION_TOKENS`, `TEMPERATURE` | -- |
| `MAX_RETRIES`, `RATE_LIMIT_DELAY` | -- |

---

## :memo: Notes

- Excel-related functionality depends on `xlwings` and a local Microsoft Excel installation. Windows is recommended for stability.
- In our previous evaluation, we used GPT-5-mini. We have since found that using frontier models or voting across multiple runs improves evaluation reliability.
- If `preprocess_info` is missing from `metadata.json`, check that the required dependencies are installed and that the file type is included in the preprocessing chain.

---

## :file_folder: Legacy Code

The code used in the original paper is maintained in the [`previous` branch](https://github.com/FinWorkBench/Finch/tree/previous). This branch contains the newer and unified implementation for Modification, Generation, and QA.

---

## :scroll: Citation

```bibtex
@article{dong2025finch,
  title={Finch: Benchmarking Finance \& Accounting across Spreadsheet-Centric Enterprise Workflows},
  author={Dong, Haoyu and Zhang, Pengkun and Gao, Yan and Dong, Xuanyu and Cheng, Yilin and Lu, Mingzhe and Yakefu, Adina and Zheng, Shuxin},
  journal={arXiv preprint arXiv:2512.13168},
  year={2025}
}
```
