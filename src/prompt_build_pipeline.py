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


def run_cmd(cmd: List[str], cwd: Path | None = None, env=None) -> None:
    """Run a command and stream output. Raise on failure."""
    printable = " ".join(shlex.quote(c) for c in cmd)
    print(f"\n=== Running ===\n{printable}\n===============")

    subprocess.run(cmd, cwd=str(cwd) if cwd else None, env=env, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run dataset pipeline commands in order.")
    parser.add_argument("--dataset-dir", default="Finch", help="organize_files --dataset-dir")
    parser.add_argument("--output-dir", required=True, help="Directory containing model output subdirectories")
    parser.add_argument("--target-dir", default="eval_set", help="organize_files --target-dir")

    parser.add_argument("--root-dir", default=None, help="preprocessor/content_builder root dir (default: --target-dir)")
    parser.add_argument("--models", default="", help="Model subdirectories to process (comma or space-separated). Example: --models model1,model2")

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
    models_raw = args.models

    # Normalize models: accept both comma-separated and space-separated, then
    # pass as individual args (both preprocessor and content_builder use nargs).
    models_list: list[str] = []
    if models_raw and models_raw.strip():
        for part in models_raw.split(","):
            for token in part.split():
                token = token.strip()
                if token:
                    models_list.append(token)

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
    if models_list:
        preproc_cmd.extend(["--models"] + models_list)
    cmds.append(preproc_cmd)

    # 3) build prompts / content parts
    content_cmd = [py, "-m", "src.build_prompt.content_builder.content_builder", root_dir]
    if models_list:
        content_cmd.extend(["--models"] + models_list)
    cmds.append(content_cmd)

    # Ensure PYTHONPATH includes project root so that `python -m` can
    # resolve package paths (e.g., src.build_prompt.content_builder).
    env = os.environ.copy()
    existing = env.get("PYTHONPATH", "")
    project_root_str = str(project_root)
    if project_root_str not in existing.split(os.pathsep):
        env["PYTHONPATH"] = project_root_str + (os.pathsep + existing if existing else "")

    for i, cmd in enumerate(cmds, start=1):
        try:
            run_cmd(cmd, cwd=project_root, env=env)
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
