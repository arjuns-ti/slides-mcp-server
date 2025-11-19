#!/usr/bin/env python3
"""
Slides MCP Server - A FastMCP server for managing Google Drive files
"""
import logging
from typing import Optional
from fastmcp import FastMCP
from dotenv import load_dotenv

# Handle both relative and absolute imports
try:
    from .drive_client import log_message
    from . import slides_tools
except ImportError:
    from drive_client import log_message
    import slides_tools

# Disable all console logging from google libraries
logging.getLogger('googleapiclient').setLevel(logging.CRITICAL)
logging.getLogger('google').setLevel(logging.CRITICAL)
logging.getLogger('google_auth_oauthlib').setLevel(logging.CRITICAL)

# Load environment variables
load_dotenv()

# Initialize FastMCP server
mcp = FastMCP("slides-mcp")




@mcp.tool()
def get_presentation_overview(presentation_id: str) -> dict:
    """
    Get presentation summary with title, slide count, and brief slide descriptions.
    
    Args:
        presentation_id: Google Slides ID from URL (after /d/)
    """
    return slides_tools.get_presentation_overview(presentation_id)


@mcp.tool()
def get_slide(presentation_id: str, slide_number: int) -> dict:
    """
    Get detailed content from a specific slide including text, images, tables with element IDs.
    
    Args:
        presentation_id: Google Slides ID
        slide_number: Slide position (1=first, 2=second, etc)
    """
    return slides_tools.get_slide(presentation_id, slide_number)


@mcp.tool()
def update_text(presentation_id: str, slide_number: int, element_id: str, text: str, 
                text_style: dict = None, paragraph_style: dict = None) -> dict:
    """
    Replace text in an element. Preserves existing formatting if text_style/paragraph_style not provided.
    
    Args:
        presentation_id: Google Slides ID
        slide_number: Slide position (1=first)
        element_id: Element ID from get_slide (e.g., "shape1")
        text: New text content (supports \\n for line breaks)
        text_style: Text formatting dict (bold, italic, fontSize, foregroundColor, etc.)
        paragraph_style: Paragraph formatting dict (alignment, lineSpacing, indents, etc.)
    """
    return slides_tools.update_text(presentation_id, slide_number, element_id, text, text_style, paragraph_style)


@mcp.tool()
def replace_slide_elements(presentation_id: str, slide_number: int, elements: list) -> dict:
    """
    Update multiple text elements at once. More efficient than multiple update_text calls.
    
    Args:
        presentation_id: Google Slides ID
        slide_number: Slide position
        elements: List of updates [{"id": "shape1", "text": "...", "text_style": {...}, "paragraph_style": {...}}]
                  text_style and paragraph_style are optional - if omitted, preserves existing formatting
    """
    return slides_tools.replace_slide_elements(presentation_id, slide_number, elements)


@mcp.tool()
def update_bullets(presentation_id: str, element_id: str, action: str = "create", 
                   bullet_preset: str = "BULLET_DISC_CIRCLE_SQUARE") -> dict:
    """
    Add or remove bullet formatting on a text element.
    
    Args:
        presentation_id: Google Slides ID
        element_id: Element ID from get_slide
        action: "create" or "delete"
        bullet_preset: Bullet style (BULLET_DISC_CIRCLE_SQUARE, NUMBERED_DECIMAL_ALPHA_ROMAN, etc.)
    """
    return slides_tools.update_bullets(presentation_id, element_id, action, bullet_preset)


@mcp.tool()
def add_element(presentation_id: str, slide_number: int, element_type: str, content: str = "", position: str = "center", chart_type: str = "", chart_data: list = []) -> dict:
    """
    Add image, table, or chart to a slide.
    
    Args:
        presentation_id: Google Slides ID
        slide_number: Slide position
        element_type: "image", "table", or "chart"
        content: Image URL (for images)
        position: "top", "center", or "bottom"
        chart_type: "bar", "line", "scatter", or "pie" (for charts)
        chart_data: [{"label": "A", "value": 10}, ...] (for charts)
    """
    element = {
        'type': element_type,
        'url': content if element_type == 'image' and content else None,
        'position': position,
        'chart_type': chart_type,
        'chart_data': chart_data
    }
    return slides_tools.add_element(presentation_id, slide_number, element)


@mcp.tool()
def duplicate_slide(presentation_id: str, source_slide: int, insert_at: Optional[int] = None) -> dict:
    """
    Duplicate a slide.
    
    Args:
        presentation_id: Google Slides ID
        source_slide: Slide to copy (1=first)
        insert_at: Position for duplicate (defaults to after source)
    """
    if insert_at is None:
        insert_at = source_slide + 1
    return slides_tools.duplicate_slide(presentation_id, source_slide, insert_at)


def main():
    """Run the FastMCP server with stdio transport"""
    log_message("Starting Slides MCP Server")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

