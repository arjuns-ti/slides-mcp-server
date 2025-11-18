"""Slides MCP Server Package"""
from .server import mcp, main
from .drive_client import get_drive_service, get_slides_service
from . import slides_tools

__all__ = ["mcp", "main", "get_drive_service", "get_slides_service", "slides_tools"]

