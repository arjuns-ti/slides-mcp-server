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
    Get presentation summary: title, slide count, brief descriptions. ~500 tokens for 10 slides.
    
    Args:
        presentation_id: Google Slides ID (required, from URL after /d/)
    """
    return slides_tools.get_presentation_overview(presentation_id)


@mcp.tool()
def get_slide(presentation_id: str, slide_number: int) -> dict:
    """
    Get all content from one slide: text, images, tables with element IDs. ~200-400 tokens.
    
    Args:
        presentation_id: Google Slides ID (required)
        slide_number: Slide position (required, 1=first, 2=second, etc)
    """
    return slides_tools.get_slide(presentation_id, slide_number)


@mcp.tool()
def update_text(presentation_id: str, slide_number: int, element_id: str, text: str) -> dict:
    """
    Replace text in one element. Get element_id from get_slide. Supports \\n for line breaks.
    
    Args:
        presentation_id: Google Slides ID (required)
        slide_number: Slide position (required, 1=first)
        element_id: Element ID like "shape1" from get_slide (required)
        text: New text content (required)
    """
    return slides_tools.update_text(presentation_id, slide_number, element_id, text)


@mcp.tool()
def replace_slide_elements(presentation_id: str, slide_number: int, elements: list) -> dict:
    """
    Update multiple text elements at once. More efficient than separate update_text calls.
    
    Args:
        presentation_id: Google Slides ID (required)
        slide_number: Slide position (required)
        elements: Array of updates (required) [{"id": "shape1", "text": "New"}, ...]
    """
    return slides_tools.replace_slide_elements(presentation_id, slide_number, elements)


@mcp.tool()
def add_element(presentation_id: str, slide_number: int, element_type: str, content: str = "", position: str = "center") -> dict:
    """
    Insert image or table on slide. Returns new element_id.
    
    Args:
        presentation_id: Google Slides ID (required)
        slide_number: Slide position (required)
        element_type: "image" or "table" (required)
        content: Image URL (optional, ignored for tables which create 2x2)
        position: "top", "center", or "bottom" (optional, default: "center")
    """
    element = {
        'type': element_type,
        'url': content if element_type == 'image' and content else None,
        'position': position
    }
    return slides_tools.add_element(presentation_id, slide_number, element)


@mcp.tool()
def duplicate_slide(presentation_id: str, source_slide: int, insert_at: Optional[int] = None) -> dict:
    """
    Clone a slide. Returns new slide number.
    
    Args:
        presentation_id: Google Slides ID (required)
        source_slide: Slide to copy (required, 1=first)
        insert_at: Position for clone (optional, defaults to after source)
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

