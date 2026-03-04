# Usage Instructions

## 1. Run `python src/eval_set_build.py`

This script is responsible for organizing the raw model output into a unified evaluation set directory structure (case directory).

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
If there are multiple files of the same type, name them as `id_1.*`, `id_2.*` or `id-1.*`, `id-2.*`. For example, if there are multiple PNG images, name them `id_1.png`, `id_2.png` or `id-1.png`, `id-2.png`.

### 1.1 Configure `src/eval_set_build.py` first

Focus on checking the following constants (modify them according to your actual data paths):

* `DATA_DIR`
* `MODEL_OUTPUT_DIR`
* `ANNOTATION_XLSX`
* `SET_OUTPUT_ROOT`
* `LIBRARY_PATH`
* Column name constants: `ID_COL`, `SRC_COL`, `ANS_COL`, `DESC_COL`, `LIMIT_COL`, `TASK_CLASS_COL`

The `ANNOTATION_XLSX` needs to be prepared by the user.

`ANNOTATION_XLSX` should include the following columns to contain sufficient information:  `ID`, `source_file`, `answer_file`, `task_description`, `task_classification(modify/generate/qa)`, `task_limit`
 `task_limit` refers to constraints on the task other than the `task_description` (if any).
### 1.2 Execute the command

```bash
python src/eval_set_build.py

```

### 1.3 Artifacts and Behavior

* Generates a `Model Name/ID/` directory structure under `SET_OUTPUT_ROOT`.
* Each case typically contains: `input.*`, `answer.*`, `output.*`, `query.txt`, and `property.json`.
* Includes additional processing logic for `md/docx/pdf/txt` and other extensions (text/image splitting, PDF to image conversion, text export, etc.).
* Special cases encountered during processing are logged to `post_process.log`.

## 2. Modify `src/config.py`

This file is the core configuration for the GPT grading pipeline.

**Must configure:**

* `azure_endpoint`, `api_key`, and `api_version` inside `client = AzureOpenAI(...)`
* `MODEL_NAME`


## 3. Modify and execute `pipeline.py`

`src/pipeline.py` orchestrates the entire evaluation process.

### 3.1 Configure `src/pipeline.py` first

* `DATASET_PATH`: The root directory of the evaluation set (usually the product of step 1).
* `JOBS`: The subdirectories of the models to be evaluated and their output result filenames.
* `BASE_ENV`, `JUDGE_ENV`: The conda environment names.

### 3.2 Execute the command

```bash
python src/pipeline.py

```

Or use the PowerShell version:

```powershell
pwsh src/pipeline.ps1

```

### 3.3 Actual execution stages of the pipeline

1. `src.recalc_with_xlwings`: Batch recalculation and saving of `xlsx` files.
2. `src.sheet_screenshot_generator.main`: Generates difference screenshots/metadata according to the task mode.
3. `src.image_merger`: Merges screenshots into `_input.png` / `_answer.png` / `_output.png`.
4. `src.gpt_judger.gpt_judge_eval`: Calls GPT for grading and outputs the result Excel file.

---

## `src` Module Overview

* `eval_set_build.py`: Builds the evaluation set case directories.
* `pipeline.py` / `pipeline.ps1`: The complete pipeline entry point.
* `recalc_with_xlwings.py`: Excel recalculation.
* `sheet_screenshot_generator/`: Screenshot generation, difference detection, and metadata.
* `image_merger.py`: Merges screenshots by case.
* `label_file_diff.py`: Excel cell difference comparison (JSON).
* `gpt_judger/`: Context construction, prompt organization, and calling Azure OpenAI for grading.
* `config.py`: Azure OpenAI and grading switch configuration.

