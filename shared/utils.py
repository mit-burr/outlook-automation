"""Shared utility functions"""
from typing import Any, Dict

def format_response(data: Any) -> Dict:
    """Template utility function"""
    return {
        "data": data,
        "success": True
    }
