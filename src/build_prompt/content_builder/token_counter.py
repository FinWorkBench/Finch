"""
Token Counter Utility

Estimates token count for content parts and provides truncation functionality.
"""

import logging
from typing import List, Dict, Any, Tuple, Optional, Set



logger = logging.getLogger(__name__)

def can_add_text_within_char_limit(
    content_parts: List[Dict[str, Any]],
    new_text: str,
    max_text_chars: int
) -> bool:
    """
    Check if adding new_text would exceed total text char limit.
    Counts only text parts, images ignored.
    """
    current_chars = estimate_text_chars_from_content_parts(content_parts)
    return (current_chars + len(new_text)) <= max_text_chars


# def estimate_tokens_from_text(text: str) -> int:
#     """
#     Estimate token count from text using character count.
    
#     Args:
#         text: Text string
    
#     Returns:
#         Estimated token count
#     """
#     return len(text) // CHARS_PER_TOKEN


# def estimate_tokens_from_content_parts(content_parts: List[Dict[str, Any]]) -> int:
#     """
#     Estimate total token count from content parts.
    
#     Only counts text content; images are not counted toward tokens.
    
#     Args:
#         content_parts: List of content part dicts
    
#     Returns:
#         Estimated token count
#     """
#     total_tokens = 0
    
#     for part in content_parts:
#         if part.get("type") == "text":
#             text = part.get("text", "")
#             total_tokens += estimate_tokens_from_text(text)
    
#     return total_tokens


# def can_add_text_within_limit(
#     content_parts: List[Dict[str, Any]],
#     new_text: str,
#     max_tokens: int
# ) -> bool:
#     """
#     Check if adding new text would exceed token limit.
    
#     Args:
#         content_parts: Existing content parts
#         new_text: Text to potentially add
#         max_tokens: Maximum allowed tokens
    
#     Returns:
#         True if adding text stays within limit
#     """
#     current_tokens = estimate_tokens_from_content_parts(content_parts)
#     new_tokens = estimate_tokens_from_text(new_text)
    
#     return (current_tokens + new_tokens) <= max_tokens


# def get_token_usage_info(
#     content_parts: List[Dict[str, Any]],
#     max_tokens: int
# ) -> Dict[str, Any]:
#     """
#     Get detailed token usage information.
    
#     Args:
#         content_parts: Content parts list
#         max_tokens: Maximum token limit
    
#     Returns:
#         Dict with token usage statistics
#     """
#     text_parts = [p for p in content_parts if p.get("type") == "text"]
#     image_parts = [p for p in content_parts if p.get("type") == "image_url"]
    
#     total_tokens = estimate_tokens_from_content_parts(content_parts)
    
#     return {
#         "total_tokens": total_tokens,
#         "max_tokens": max_tokens,
#         "remaining_tokens": max(0, max_tokens - total_tokens),
#         "usage_percentage": (total_tokens / max_tokens * 100) if max_tokens > 0 else 0,
#         "text_parts_count": len(text_parts),
#         "image_parts_count": len(image_parts),
#         "total_parts_count": len(content_parts),
#         "exceeds_limit": total_tokens > max_tokens
#     }







def estimate_text_chars_from_content_parts(content_parts: List[Dict[str, Any]]) -> int:
    total_chars = 0
    for part in content_parts:
        if part.get("type") == "text":
            total_chars += len(part.get("text", ""))
    return total_chars


def count_images(content_parts: List[Dict[str, Any]]) -> int:
    return sum(1 for p in content_parts if p.get("type") == "image_url")


def _remove_images_from_end(
    parts: List[Dict[str, Any]],
    max_images: int,
    protected_indices: Set[int],
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Remove image_url parts from the end until image count <= max_images.
    Always respects protected_indices (e.g. first part, note part).
    """
    if max_images is None or max_images < 0:
        return parts, 0

    removed = 0
    while count_images(parts) > max_images:
        removed_any = False
        for i in range(len(parts) - 1, -1, -1):
            if i in protected_indices:
                continue
            if parts[i].get("type") == "image_url":
                parts.pop(i)
                removed += 1
                removed_any = True
                break
        if not removed_any:
            # Only protected images remain; cannot remove further
            break

    return parts, removed


def _truncate_text_chars_from_end(
    parts: List[Dict[str, Any]],
    max_text_chars: int,
    protected_indices: Set[int],
) -> Tuple[List[Dict[str, Any]], int, int]:
    """
    Truncate text from the end until total text chars <= max_text_chars.
    Do NOT delete text parts; only shorten their 'text' field.
    Always respects protected_indices.
    Returns: (parts, chars_removed, text_parts_touched)
    """
    if max_text_chars is None or max_text_chars < 0:
        return parts, 0, 0

    total_chars = estimate_text_chars_from_content_parts(parts)
    if total_chars <= max_text_chars:
        return parts, 0, 0

    need_remove = total_chars - max_text_chars
    chars_removed = 0
    parts_touched = 0

    for i in range(len(parts) - 1, -1, -1):
        if need_remove <= 0:
            break
        if i in protected_indices:
            continue

        part = parts[i]
        if part.get("type") != "text":
            continue

        text = part.get("text", "")
        if not text:
            continue

        parts_touched += 1

        if len(text) > need_remove:
            part["text"] = text[:-need_remove]
            chars_removed += need_remove
            need_remove = 0
            break
        else:
            # truncate this part to empty string
            part["text"] = ""
            chars_removed += len(text)
            need_remove -= len(text)

    return parts, chars_removed, parts_touched


def truncate_content_parts(
    content_parts: List[Dict[str, Any]],
    max_images: Optional[int] = None,
    max_text_chars: Optional[int] = None,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Only enforce:
      1) max_images: remove image_url parts from the end
      2) max_text_chars: truncate text chars from the end (do NOT delete text parts)
      3) If anything changed, append a note text part (protected from truncation)

    Always preserve first part (index 0).
    """
    if not content_parts:
        return content_parts, {
            "was_truncated": False,
            "images_removed": 0,
            "text_chars_removed": 0,
            "text_parts_truncated": 0,
            "note_added": False,
            "original_images": 0,
            "final_images": 0,
            "original_text_chars": 0,
            "final_text_chars": 0,
        }

    # Shallow copy each dict to avoid mutating original input
    parts: List[Dict[str, Any]] = [dict(p) for p in content_parts]

    original_images = count_images(parts)
    original_chars = estimate_text_chars_from_content_parts(parts)

    # Always protect the first part (base prompt)
    protected: Set[int] = {0}

    # 1) Remove extra images (from the end)
    images_removed = 0
    if max_images is not None:
        parts, images_removed = _remove_images_from_end(parts, max_images, protected)

    # 2) Truncate extra text chars (from the end)
    text_chars_removed = 0
    text_parts_truncated = 0
    if max_text_chars is not None:
        parts, text_chars_removed, text_parts_truncated = _truncate_text_chars_from_end(
            parts, max_text_chars, protected
        )

    changed = (images_removed > 0) or (text_chars_removed > 0)

    # 3) Append note if changed; and protect it from any further truncation
    note_added = False
    if changed:
        final_images_before_note = count_images(parts)
        final_chars_before_note = estimate_text_chars_from_content_parts(parts)

        note_lines = [
            "⚠ Content was reduced to satisfy input limits.",
            f"- Original: text_chars={original_chars}, images={original_images}",
            f"- Current : text_chars={final_chars_before_note}, images={final_images_before_note}",
        ]
        if max_images is not None and images_removed > 0:
            note_lines.append(
                f"- Removed {images_removed} image(s) from the end to satisfy max_images={max_images}."
            )
        if max_text_chars is not None and text_chars_removed > 0:
            note_lines.append(
                f"- Truncated {text_chars_removed} text char(s) from the end to satisfy max_text_chars={max_text_chars}."
            )

        parts.append({"type": "text", "text": "\n".join(note_lines)})
        note_added = True

        # Protect base prompt + note
        protected = {0, len(parts) - 1}

        # Re-apply limits after adding note (so note doesn't push us over),
        # but do NOT allow truncation/removal of the note itself.
        if max_images is not None:
            parts, removed_after_note = _remove_images_from_end(parts, max_images, protected)
            images_removed += removed_after_note

        if max_text_chars is not None:
            parts, removed_chars_after_note, touched_after_note = _truncate_text_chars_from_end(
                parts, max_text_chars, protected
            )
            text_chars_removed += removed_chars_after_note
            text_parts_truncated += touched_after_note

    final_images = count_images(parts)
    final_chars = estimate_text_chars_from_content_parts(parts)

    info = {
        "was_truncated": changed,
        "images_removed": images_removed,
        "text_chars_removed": text_chars_removed,
        "text_parts_truncated": text_parts_truncated,
        "note_added": note_added,
        "original_images": original_images,
        "final_images": final_images,
        "original_text_chars": original_chars,
        "final_text_chars": final_chars,
    }

    if changed:
        logger.warning(
            f"Content reduced: images {original_images}->{final_images} "
            f"(removed={images_removed}), text_chars {original_chars}->{final_chars} "
            f"(removed={text_chars_removed}). note_added={note_added}"
        )

    return parts, info

