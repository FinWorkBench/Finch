# pipeline.py
# Usage
# python pipeline.py
# python pipeline.py --models opus_4.5_output --target-dir data/eval_dataset_opus

import argparse
import os
import shlex
import subprocess
import sys
from pathlib import Path
from typing import List


def run_cmd(cmd: List[str], cwd: Path | None = None) -> None:
    """Run a command and stream output. Raise on failure."""
    printable = " ".join(shlex.quote(c) for c in cmd)
    print(f"\n=== Running ===\n{printable}\n===============")

    subprocess.run(cmd, cwd=str(cwd) if cwd else None, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run dataset pipeline commands in order.")
    parser.add_argument("--dataset-dir", default="Finch", help="organize_files --dataset-dir")
    parser.add_argument("--output-dir", default="Your model output", help="organize_files --output-dir")
    parser.add_argument("--target-dir", default="eval_set", help="organize_files --target-dir")

    parser.add_argument("--root-dir", default=None, help="preprocessor/content_builder root dir (default: --target-dir)")
    parser.add_argument("--models", default="", help="preprocessor --models (comma-separated allowed)")

    parser.add_argument("--project-root", default=".", help="Working directory to run commands from")
    parser.add_argument("--continue-on-error", action="store_true", help="Continue even if a step fails")

    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    if not project_root.exists():
        raise FileNotFoundError(f"project root not found: {project_root}")

    dataset_dir = args.dataset_dir
    output_dir = args.output_dir
    target_dir = args.target_dir

    root_dir = args.root_dir or target_dir
    models_arg = args.models

    py = sys.executable  # current python interpreter

    cmds = []

    # 1) organize_files
    cmds.append([
        py, "src/organize_files.py",
        "--dataset-dir", dataset_dir,
        "--output-dir", output_dir,
        "--target-dir", target_dir
    ])

    # 2) preprocessor (only pass --models if user provided)
    preproc_cmd = [py, "src/preprocessor/preprocessor_main.py", "--root-dir", root_dir]
    if models_arg and models_arg.strip():
        preproc_cmd.extend(["--models", models_arg])
    cmds.append(preproc_cmd)

    # 3) build prompts / content parts
    cmds.append([py, "-m", "src.build_prompt.content_builder.content_builder", root_dir])

    for i, cmd in enumerate(cmds, start=1):
        try:
            run_cmd(cmd, cwd=project_root)
        except subprocess.CalledProcessError as e:
            print(f"\n!!! Step {i} failed with exit code {e.returncode}")
            if args.continue_on_error:
                print("Continuing due to --continue-on-error ...")
                continue
            return e.returncode

    print("\n✅ Pipeline finished successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
