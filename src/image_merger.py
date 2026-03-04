# pip install pillow

import argparse
import logging
from pathlib import Path
from typing import List, Tuple, Optional, Dict

from PIL import Image

# Disable DecompressionBombError for very large images (use with care)
Image.MAX_IMAGE_PIXELS = None


# ========== Logging configuration ==========
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    filename="logs/image_merger.log",
    filemode="a",
)
logger = logging.getLogger(__name__)
# ==========================================


def merge_images_to_one(
    image_paths: List[Path],
    out_path: Path,
    direction: str = "vertical",
    padding: int = 10,
    bg_color: Tuple[int, int, int] = (255, 255, 255),
) -> Optional[Path]:
    """
    Merge multiple PNG images into a single large image.

    Args:
        image_paths:
            Ordered list of image paths to merge.
        out_path:
            Output PNG path, e.g. case_dir / "input.png".
        direction:
            "vertical" for top-to-bottom stacking,
            "horizontal" for left-to-right stacking.
        padding:
            Number of pixels between adjacent images.
        bg_color:
            Background color (R, G, B) used for the canvas.

    Returns:
        The output path if successful, None otherwise or if no valid images.
    """
    # Filter out non-existent paths just in case
    image_paths = [p for p in image_paths if p.is_file()]
    if not image_paths:
        msg = f"[INFO] No images to merge for {out_path.name}, skipped."
        print(msg)
        logger.info(msg)
        return None

    images = []
    try:
        # Open all images and convert to RGB to avoid mode issues
        for p in image_paths:
            img = Image.open(p).convert("RGB")
            images.append(img)

        widths, heights = zip(*(img.size for img in images))

        if direction == "vertical":
            # Vertical merge:
            #   - final width is the maximum width of all images
            #   - final height is the sum of all heights plus padding
            max_width = max(widths)
            total_height = sum(heights) + padding * (len(images) - 1)

            merged = Image.new("RGB", (max_width, total_height), bg_color)

            y = 0
            for img in images:
                w, h = img.size
                # Center horizontally within the max width
                x = (max_width - w) // 2
                merged.paste(img, (x, y))
                y += h + padding

        else:
            # Horizontal merge:
            #   - final height is the maximum height of all images
            #   - final width is the sum of all widths plus padding
            max_height = max(heights)
            total_width = sum(widths) + padding * (len(images) - 1)

            merged = Image.new("RGB", (total_width, max_height), bg_color)

            x = 0
            for img in images:
                w, h = img.size
                # Center vertically within the max height
                y = (max_height - h) // 2
                merged.paste(img, (x, y))
                x += w + padding

        # Ensure parent directory exists before saving
        out_path.parent.mkdir(parents=True, exist_ok=True)
        merged.save(out_path)

        msg = f"[OK] Merged {len(images)} images -> {out_path}"
        print(msg)
        logger.info(msg)
        return out_path

    except Exception as e:
        msg = f"[ERROR] Failed to merge images for {out_path}: {repr(e)}"
        print(f"[WARN] {msg}")
        logger.warning(msg)
        return None

    finally:
        # Always close all Image objects to avoid leaking file handles
        for img in images:
            try:
                img.close()
            except Exception:
                pass


def collect_case_images(case_dir: Path) -> Dict[str, List[Path]]:
    """
    Collect input/answer/output screenshot groups for a single case directory.

    Args:
        case_dir:
            Directory representing a single test case.

    Returns:
        A mapping of logical name -> sorted list of image paths.
    """
    return {
        "input": sorted(case_dir.glob("_input_*.png")),
        "answer": sorted(case_dir.glob("_answer_*.png")),
        "output": sorted(case_dir.glob("_output_*.png")),
    }


def process_case_dir(case_dir: Path, direction: str = "vertical") -> None:
    """
    Process a single case directory:
        - Merge _input_*.png into _input.png
        - Merge _answer_*.png into _answer.png
        - Merge _output_*.png into _output.png

    Args:
        case_dir:
            Directory representing a single test case under root_dir.
        direction:
            Merge direction for all image groups ("vertical" or "horizontal").
    """
    if not case_dir.is_dir():
        return

    print(f"\n[CASE] {case_dir}")
    logger.info("Processing case dir: %s", case_dir)

    image_groups = collect_case_images(case_dir)

    # Merge _input_*
    if image_groups["input"]:
        merge_images_to_one(
            image_groups["input"],
            case_dir / "_input.png",
            direction=direction,
        )

    # Merge _answer_*
    if image_groups["answer"]:
        merge_images_to_one(
            image_groups["answer"],
            case_dir / "_answer.png",
            direction=direction,
        )

    # Merge _output_*
    if image_groups["output"]:
        merge_images_to_one(
            image_groups["output"],
            case_dir / "_output.png",
            direction=direction,
        )


def parse_args() -> argparse.Namespace:
    """
    Parse command line arguments.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Merge multiple PNG screenshots (_input_*.png / _answer_*.png / _output_*.png) "
            "into single _input.png / _answer.png / _output.png for each case directory."
        )
    )
    parser.add_argument(
        "--root-dir",
        type=Path,
        required=True,
        help="Root directory containing case subdirectories.",
    )
    parser.add_argument(
        "--direction",
        type=str,
        choices=("vertical", "horizontal"),
        default="vertical",
        help="Merge direction for images (default: vertical).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    root_dir = args.root_dir.resolve()
    direction = args.direction

    if not root_dir.is_dir():
        raise NotADirectoryError(f"Root dir not found: {root_dir}")

    # Assume each direct subdirectory under root_dir is a case directory
    for case_dir in sorted(root_dir.iterdir()):
        if not case_dir.is_dir():
            continue
        try:
            process_case_dir(case_dir, direction=direction)
        except Exception as e:
            # Protect the outer loop so a single bad case does not stop the whole run
            msg = f"[CASE-ERROR] Failed to process case {case_dir}: {repr(e)}"
            print(f"[WARN] {msg}")
            logger.warning(msg)
            continue

    print("\n[DONE] All cases processed.")
    logger.info("All cases processed.")


if __name__ == "__main__":
    main()
