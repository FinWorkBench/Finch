#!/usr/bin/env python
# -*- coding: utf-8 -*-



import os
import subprocess
import sys

# ============== （ ps1 ） ============== 

# DATASET_PATH = os.path.join("data", "eval_dataset_web")
# DATASET_PATH = os.path.join("data", "eval_dataset_api")
DATASET_PATH = os.path.join("data", "eval_dataset_api")

# (SUB_PATH, RESULTS) 
JOBS: list[tuple[str, str]] = [
    # ("gptpro", "gptpro_webcase_gpt5mini_1213.xlsx"),
    # ("finch_anthropic_claude-sonnet-4.5", "anthropic_claude-sonnet-4.5_apicase_gpt5mini_1214.xlsx"),
    # ("finch_google_gemini-3-pro-preview", "google_gemini-3-pro-preview_apicase_gpt5mini_1214.xlsx"),
    # ("finch_openai_gpt-5.1", "openai_gpt-5.1_apicase_gpt5mini_1214.xlsx"),
    # ("finch_qwen_qwen3-max", "qwen_qwen3-max_apicase_gpt5mini_1214.xlsx"),
    # ("finch_x-ai_grok-4", "x-ai_grok-4_apicase_gpt5mini_1214.xlsx"),
    # ("claudesonnet", "claudesonnet_webcase_gpt5mini_1213.xlsx"),
    
    # ("finch_anthropic_claude-sonnet-4.5", "anthropic_claude-sonnet-4.5_apicase_gpt5mini_1207.xlsx"),
]

# conda 
BASE_ENV = "base"
JUDGE_ENV = "base"


# ============== ==============

def run_cmd(cmd: list[str]) -> None:
    print(">>> Running:", " ".join(cmd))
    try:
        result = subprocess.run(cmd)
    except FileNotFoundError as e:
        print(f"!!! Failed to run command: {e}")
        sys.exit(1)

    if result.returncode != 0:
        print(f"!!! Command exited with code {result.returncode}")
        sys.exit(result.returncode)


# ============== ==============

def main():
    print(f">>> DATASET_PATH = {DATASET_PATH}")
    print(">>> JOBS:")
    for sub_path, results in JOBS:
        dir_path = os.path.join(DATASET_PATH, sub_path)
        print(f"    - SUB_PATH = {sub_path}")
        print(f"      DIR      = {dir_path}")
        print(f"      RESULTS  = {results}")
    print()

    if not JOBS:
        print("!!! JOBS is empty. Please configure at least one (SUB_PATH, RESULTS).")
        sys.exit(1)

    # ================== 1 ： base ================== 
    print(f">>> [ENV] conda env: {BASE_ENV}")

    # 1. （ DATASET_PATH ） 
    print(">>> [STEP 1] Running recalc_with_xlwings ...")
    run_cmd([
        "conda", "run", "-n", BASE_ENV,
        "python", "-m", "src.recalc_with_xlwings",
        "-f", DATASET_PATH,
    ])
    print(">>> [STEP 1] Finished recalc_with_xlwings")
    print()

    # (SUB_PATH, RESULTS) 
    for sub_path, results in JOBS:
        dir_path = os.path.join(DATASET_PATH, sub_path)
        print("=" * 80)
        print(f">>> Start job for SUB_PATH = {sub_path}")
        print(f">>> DIR     = {dir_path}")
        print(f">>> RESULTS = {results}")
        print("=" * 80)
        print()

        # 2. Excel & CSV & metadata 
        print(">>> [STEP 2] Running sheet_screenshot_generator ...")
        run_cmd([
            "conda", "run", "-n", BASE_ENV,
            "python", "-m", "src.sheet_screenshot_generator.main",
            "--dataset", dir_path,
        ])
        print(">>> [STEP 2] Finished sheet_screenshot_generator")
        print()

        # 3. （ sheet input.png / answer.png / output.png ） 
        print(">>> [STEP 3] Running image_merger ...")
        run_cmd([
            "conda", "run", "-n", BASE_ENV,
            "python", "-m", "src.image_merger",
            "--root-dir", dir_path,
        ])
        print(">>> [STEP 3] Finished image_merger")
        print()

        # ================== 2 ： spreadsheetllm ================== 
        print(f">>> [ENV] conda env: {JUDGE_ENV}")

        # 4. gpt 
        print(">>> [STEP 4] Running gpt_judger ...")
        run_cmd([
            "conda", "run", "-n", JUDGE_ENV,
            "python", "-m", "src.gpt_judger.gpt_judge_eval",
            "--root-dir", dir_path,
            "--out-excel", results,
        ])
        print(">>> [STEP 4] Finished gpt_judger")
        print()

        print(f">>> Job for SUB_PATH = {sub_path} finished.")
        print()

    print(">>> All jobs done.")


if __name__ == "__main__":
    main()
