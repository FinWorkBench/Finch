"""
Example: Demonstrating Content Parts Truncation

This script shows how the token truncation feature works.
"""

from content_builder.token_counter import (
    truncate_content_parts
)


def create_sample_content_parts():
    """Create sample content parts for demonstration."""
    content_parts = [
        {
            "type": "text",
            "text": "## Base Prompt\n\nYou are an expert evaluator. " * 50  # ~100 tokens
        },
        {
            "type": "text",
            "text": "## Instruction\n\nCalculate the sum of values in column A. " * 100  # ~200 tokens
        },
        {
            "type": "image_url",
            "image_url": {"url": "data:image/png;base64,iVBORw0KGgo..."}
        },
        {
            "type": "text",
            "text": "## Reference Output\n\nThe sum is 150. " * 200  # ~400 tokens
        },
        {
            "type": "image_url",
            "image_url": {"url": "data:image/png;base64,iVBORw0KGgo..."}
        },
        {
            "type": "text",
            "text": "## Model Output\n\nThe sum is 145. " * 200  # ~400 tokens
        },
        {
            "type": "image_url",
            "image_url": {"url": "data:image/png;base64,iVBORw0KGgo..."}
        },
        {
            "type": "text",
            "text": "## Additional Context\n\nHere is more information. " * 300  # ~600 tokens
        }
    ]
    return content_parts


def demonstrate_truncation():
    """Demonstrate the truncation feature."""
    
    print("="*60)
    print("Token Truncation Demonstration")
    print("="*60)
    
    # Create sample content
    content_parts = create_sample_content_parts()
    
    print(f"\n1. Original Content:")
    print(f"   Total parts: {len(content_parts)}")
    
    original_tokens = estimate_tokens_from_content_parts(content_parts)
    print(f"   Estimated tokens: {original_tokens}")
    
    for i, part in enumerate(content_parts):
        if part["type"] == "text":
            print(f"   Part {i}: Text ({len(part['text'])} chars)")
        else:
            print(f"   Part {i}: Image")
    
    # Set a token limit
    max_tokens = 500
    
    print(f"\n2. Token Limit: {max_tokens}")
    print(f"   Tokens over limit: {original_tokens - max_tokens}")
    
    # Get usage info before truncation
    usage_before = get_token_usage_info(content_parts, max_tokens)
    print(f"\n3. Usage Before Truncation:")
    print(f"   Total tokens: {usage_before['total_tokens']}")
    print(f"   Max tokens: {usage_before['max_tokens']}")
    print(f"   Usage: {usage_before['usage_percentage']:.1f}%")
    print(f"   Exceeds limit: {usage_before['exceeds_limit']}")
    
    # Truncate
    print(f"\n4. Applying Truncation (from the end)...")
    truncated_parts, truncation_info = truncate_content_parts(content_parts, max_tokens)
    
    print(f"\n5. Truncation Results:")
    print(f"   Was truncated: {truncation_info['was_truncated']}")
    print(f"   Original tokens: {truncation_info['original_tokens']}")
    print(f"   Final tokens: {truncation_info['final_tokens']}")
    print(f"   Parts removed: {truncation_info['parts_removed']}")
    print(f"   Text parts removed: {truncation_info['text_parts_removed']}")
    print(f"   Image parts removed: {truncation_info['image_parts_removed']}")
    
    print(f"\n6. Truncated Content:")
    print(f"   Total parts: {len(truncated_parts)}")
    
    for i, part in enumerate(truncated_parts):
        if part["type"] == "text":
            preview = part['text'][:50] + "..." if len(part['text']) > 50 else part['text']
            print(f"   Part {i}: Text - '{preview}'")
        else:
            print(f"   Part {i}: Image")
    
    # Get usage info after truncation
    usage_after = get_token_usage_info(truncated_parts, max_tokens)
    print(f"\n7. Usage After Truncation:")
    print(f"   Total tokens: {usage_after['total_tokens']}")
    print(f"   Max tokens: {usage_after['max_tokens']}")
    print(f"   Usage: {usage_after['usage_percentage']:.1f}%")
    print(f"   Exceeds limit: {usage_after['exceeds_limit']}")
    
    print(f"\n8. Summary:")
    print(f"   ✓ First part (base prompt) preserved")
    print(f"   ✓ Parts removed from the end: {truncation_info['parts_removed']}")
    print(f"   ✓ Token reduction: {truncation_info['original_tokens']} → {truncation_info['final_tokens']}")
    print(f"   ✓ Now within limit: {not usage_after['exceeds_limit']}")
    
    print("\n" + "="*60)


def demonstrate_no_truncation_needed():
    """Demonstrate when no truncation is needed."""
    
    print("\n" + "="*60)
    print("No Truncation Needed Demonstration")
    print("="*60)
    
    # Create smaller content
    content_parts = [
        {"type": "text", "text": "Base prompt here. " * 10},  # ~20 tokens
        {"type": "text", "text": "Instruction here. " * 10},  # ~20 tokens
    ]
    
    max_tokens = 1000
    
    print(f"\n1. Content tokens: {estimate_tokens_from_content_parts(content_parts)}")
    print(f"2. Token limit: {max_tokens}")
    
    truncated_parts, truncation_info = truncate_content_parts(content_parts, max_tokens)
    
    print(f"\n3. Truncation Info:")
    print(f"   Was truncated: {truncation_info['was_truncated']}")
    print(f"   Parts removed: {truncation_info['parts_removed']}")
    print(f"\n   ✓ No truncation needed - content within limit")
    
    print("="*60)


if __name__ == "__main__":
    # Run demonstrations
    demonstrate_truncation()
    demonstrate_no_truncation_needed()
    
    print("\n✓ Truncation feature demonstration complete!")
