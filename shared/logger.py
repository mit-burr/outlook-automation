# shared/logger.py
from datetime import datetime
from enum import Enum
from typing import Any, Dict, List, Optional, Union
import json

class Colors:
    INFO = "\033[36m"     # cyan
    WARN = "\033[33m"     # yellow
    ERROR = "\033[31m"    # red
    SUCCESS = "\033[32m"  # green
    DEBUG = "\033[35m"    # magenta
    RESET = "\033[0m"     # reset
    BRIGHT = "\033[1m"    # bright/bold

class Emojis:
    INFO = "ℹ️ "
    WARN = "⚠️ "
    ERROR = "❌"
    SUCCESS = "✅"
    DEBUG = "🔍"
    START = "🚀"
    END = "🏁"
    DATABASE = "🗃️ "
    TEST = "🧪"

class LogLevel(str, Enum):
    INFO = "info"
    WARN = "warn"
    ERROR = "error"
    SUCCESS = "success"
    DEBUG = "debug"

class Logger:
    def __init__(self):
        self.colors = {
            LogLevel.INFO: Colors.INFO,
            LogLevel.WARN: Colors.WARN,
            LogLevel.ERROR: Colors.ERROR,
            LogLevel.SUCCESS: Colors.SUCCESS,
            LogLevel.DEBUG: Colors.DEBUG
        }
        
        self.emojis = {
            "info": Emojis.INFO,
            "warn": Emojis.WARN,
            "error": Emojis.ERROR,
            "success": Emojis.SUCCESS,
            "debug": Emojis.DEBUG,
            "start": Emojis.START,
            "end": Emojis.END,
            "database": Emojis.DATABASE,
            "test": Emojis.TEST
        }

    def _timestamp(self) -> str:
        """Get current timestamp in ISO format."""
        return datetime.now().isoformat()

    def _format(self, message: str, level: LogLevel, prefix: Optional[str] = None) -> str:
        """Format message with color, emoji, and timestamp."""
        emoji = self.emojis.get(prefix or level)
        color = self.colors[level]
        return f"{color}{emoji} {message}{Colors.RESET}"

    def _format_value(self, value: Any, indent: int = 0) -> str:
        """Format value for pretty printing."""
        spaces = '  ' * indent
        
        if isinstance(value, (list, tuple)):
            if not value:
                return f"{spaces}[]"
            return '\n' + '\n'.join(f"{spaces}- {self._format_value(item, indent + 1)}" 
                                  for item in value)
        
        if value is None:
            return 'null'
        
        if isinstance(value, dict):
            if not value:
                return f"{spaces}{{}}"
            return '\n' + '\n'.join(
                f"{spaces}  {key}: {self._format_value(val, indent + 1)}"
                for key, val in value.items()
            )
            
        return str(value)

    def info(self, message: str, prefix: Optional[str] = None) -> None:
        """Log info message."""
        print(self._format(message, LogLevel.INFO, prefix))

    def warn(self, message: str, prefix: Optional[str] = None) -> None:
        """Log warning message."""
        print(self._format(message, LogLevel.WARN, prefix))

    def error(self, message: str, prefix: Optional[str] = None) -> None:
        """Log error message."""
        print(self._format(message, LogLevel.ERROR, prefix))

    def success(self, message: str, prefix: Optional[str] = None) -> None:
        """Log success message."""
        print(self._format(message, LogLevel.SUCCESS, prefix))

    def debug(self, message: str, prefix: Optional[str] = None) -> None:
        """Log debug message."""
        print(self._format(message, LogLevel.DEBUG, prefix))

    def start_section(self, title: str) -> None:
        """Start a new section with title."""
        print(f"\n{self._format(f'=== Starting: {title} ===', LogLevel.INFO, 'start')}")

    def end_section(self, title: str) -> None:
        """End a section with title."""
        print(f"{self._format(f'=== Completed: {title} ===', LogLevel.SUCCESS, 'end')}\n")

    def list(self, title: str, items: List[Any], level: LogLevel = LogLevel.INFO, 
            prefix: Optional[str] = None) -> None:
        """Log a list of items with a title."""
        formatted_title = f"{title}:" if title else ""
        formatted_items = self._format_value(items)
        message = f"{formatted_title}{formatted_items}"
        print(self._format(message, level, prefix))

    def db_query(self, query: str) -> None:
        """Log database query."""
        self.debug(f"Executing query: {query}", "database")

    def test_start(self, suite_name: str) -> None:
        """Log test suite start."""
        self.info(f"Starting test suite: {suite_name}", "test")

    def test_end(self, suite_name: str) -> None:
        """Log test suite completion."""
        self.success(f"Completed test suite: {suite_name}", "test")

# Create singleton instance
logger = Logger()

# Example usage:
if __name__ == "__main__":
    # Basic logging
    logger.info("This is an info message")
    logger.warn("This is a warning message")
    logger.error("This is an error message")
    logger.success("This is a success message")
    logger.debug("This is a debug message")
    
    # Sections
    logger.start_section("Test Section")
    logger.info("Doing some work...")
    logger.end_section("Test Section")
    
    # Lists
    test_data = {
        "name": "Test User",
        "items": ["item1", "item2"],
        "details": {
            "age": 30,
            "active": True
        }
    }
    logger.list("Test Data", [test_data])
    
    # Database and test logging
    logger.db_query("SELECT * FROM users")
    logger.test_start("User Tests")
    logger.test_end("User Tests")