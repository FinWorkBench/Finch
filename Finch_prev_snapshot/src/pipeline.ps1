#!/usr/bin/env pwsh
# -*- coding: utf-8 -*-

<#
PowerShell equivalent of `pipeline.py`, supporting multiple
`(SUB_PATH, RESULTS)` jobs.

The workflow is:
1. Run in `base`:
   - `python -m src.recalc_with_xlwings -f DATASET_PATH`
   - `python -m src.sheet_screenshot_generator.main --dataset DIR`
   - `python -m src.image_merger --root-dir DIR`
2. Run in `spreadsheetllm`:
   - `python -m src.gpt_judger.gpt_judge_eval --root-dir DIR --out-excel RESULTS`

Notes:
- `DATASET_PATH` is the dataset root (for example `data/eval_dataset_test`).
- `JOBS` stores multiple `(SUB_PATH, RESULTS)` pairs.
  * `SUB_PATH` is relative to `DATASET_PATH`.
  * `RESULTS` is the output filename for that job.
#>

# ============== （ Python ） ============== 

# $DATASET_PATH = Join-Path "data" "eval_dataset_web"
# $DATASET_PATH = Join-Path "data" "eval_dataset_api"
$DATASET_PATH = Join-Path "data" "eval_dataset"

# (SUB_PATH, RESULTS) 
$JOBS = @(
    @{ SubPath = "gptpro";                             Results = "gptpro_webcase_gpt5mini_1212.xlsx" }
    @{ SubPath = "finch_anthropic_claude-sonnet-4.5"; Results = "anthropic_claude-sonnet-4.5_apicase_gpt5mini_1212.xlsx" }
    @{ SubPath = "finch_google_gemini-3-pro-preview";  Results = "google_gemini-3-pro-preview_apicase_gpt5mini_1212.xlsx" }
    @{ SubPath = "finch_openai_gpt-5.1";              Results = "openai_gpt-5.1_apicase_gpt5mini_1212.xlsx" }
    @{ SubPath = "finch_qwen_qwen3-max";              Results = "qwen_qwen3-max_apicase_gpt5mini_1212.xlsx" }
    @{ SubPath = "finch_x-ai_grok-4";                 Results = "x-ai_grok-4_apicase_gpt5mini_1212.xlsx" }
    @{ SubPath = "claudesonnet";                      Results = "claudesonnet_webcase_gpt5mini_1212.xlsx" }
    
    # @{ SubPath = "finch_anthropic_claude-sonnet-4.5"; Results = "anthropic_claude-sonnet-4.5_apicase_gpt5mini_1207.xlsx" }
)

# conda 
$BASE_ENV  = "base"
$JUDGE_ENV = "base"

# ============== ==============

Write-Host ">>> DATASET_PATH = $DATASET_PATH"
Write-Host ">>> JOBS:"
foreach ($job in $JOBS) {
    $subPath = $job.SubPath
    $results = $job.Results
    $dirPath = Join-Path $DATASET_PATH $subPath

    Write-Host "    - SUB_PATH = $subPath"
    Write-Host "      DIR      = $dirPath"
    Write-Host "      RESULTS  = $results"
}
Write-Host ""

if (-not $JOBS -or $JOBS.Count -eq 0) {
    Write-Host "!!! JOBS is empty. Configure at least one (SUB_PATH, RESULTS)." -ForegroundColor Red
    exit 1
}

# ================== 1 ： base ================== 
Write-Host ">>> [ENV] conda env: $BASE_ENV"

# 1. （ DATASET_PATH ） 
Write-Host ">>> [STEP 1] Running recalc_with_xlwings ..."
conda run -n $BASE_ENV python -u -m src.recalc_with_xlwings -f "$DATASET_PATH"
if ($LASTEXITCODE -ne 0) {
    Write-Host "!!! [STEP 1] recalc_with_xlwings exited with code $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host ">>> [STEP 1] Finished recalc_with_xlwings"
Write-Host ""

# (SUB_PATH, RESULTS) 
foreach ($job in $JOBS) {
    $subPath = $job.SubPath
    $results = $job.Results
    $dirPath = Join-Path $DATASET_PATH $subPath

    Write-Host ("=" * 80)
    Write-Host ">>> Start job for SUB_PATH = $subPath"
    Write-Host ">>> DIR     = $dirPath"
    Write-Host ">>> RESULTS = $results"
    Write-Host ("=" * 80)
    Write-Host ""

    # 2. Excel & CSV & metadata 
    Write-Host ">>> [STEP 2] Running sheet_screenshot_generator ..."
    conda run -n $BASE_ENV python -u -m src.sheet_screenshot_generator.main --dataset "$dirPath"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "!!! [STEP 2] sheet_screenshot_generator exited with code $LASTEXITCODE" -ForegroundColor Red
        exit $LASTEXITCODE
    }
    Write-Host ">>> [STEP 2] Finished sheet_screenshot_generator"
    Write-Host ""

    # 3. （ sheet input.png / answer.png / output.png ） 
    Write-Host ">>> [STEP 3] Running image_merger ..."
    conda run -n $BASE_ENV python -u -m src.image_merger --root-dir "$dirPath"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "!!! [STEP 3] image_merger exited with code $LASTEXITCODE" -ForegroundColor Red
        exit $LASTEXITCODE
    }
    Write-Host ">>> [STEP 3] Finished image_merger"
    Write-Host ""

    # ================== 2 ： spreadsheetllm ================== 
    Write-Host ">>> [ENV] conda env: $JUDGE_ENV"

    # 4. gpt 
    Write-Host ">>> [STEP 4] Running gpt_judger ..."
    conda run -n $JUDGE_ENV python -u -m src.gpt_judger.gpt_judge_eval --root-dir "$dirPath" --out-excel "$results"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "!!! [STEP 4] gpt_judger exited with code $LASTEXITCODE" -ForegroundColor Red
        exit $LASTEXITCODE
    }
    Write-Host ">>> [STEP 4] Finished gpt_judger"
    Write-Host ""

    Write-Host ">>> Job for SUB_PATH = $subPath finished."
    Write-Host ""
}

Write-Host ">>> All jobs done."
