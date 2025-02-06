# services/outlook_service/models.py
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Optional

@dataclass
class Meeting:
    """Represents a calendar meeting with all its properties."""
    subject: str
    start_time: datetime
    end_time: datetime
    duration: int  # in minutes
    organizer: str
    is_recurring: bool
    series_id: str
    location: Optional[str] = None
    categories: list[str] = None

    @property
    def display_dict(self) -> dict:
        """Returns a dictionary of properties to display in the UI."""
        return {
            "time": self.start_time.strftime("%H:%M"),
            "duration": f"{self.rounded_duration} min",
            "organizer": self.organizer.split(',')[0]  # Just last name
        }

    @property
    def rounded_duration(self) -> int:
        """Returns duration rounded up to nearest 30m interval."""
        return ((self.duration + 29) // 30) * 30  # Round up to nearest 30m

    @property
    def weekday(self) -> str:
        """Returns the weekday name of the meeting."""
        return self.start_time.strftime("%A")

    @classmethod
    def from_outlook_item(cls, item) -> 'Meeting':
        """Create a Meeting instance from an Outlook appointment item."""
        return cls(
            subject=item.Subject,
            start_time=item.Start,
            end_time=item.End,
            duration=item.Duration,
            organizer=item.Organizer,
            is_recurring=bool(getattr(item, 'RecurrenceState', 0)),
            series_id=str(getattr(item, 'ConversationID', 'N/A')),
            location=getattr(item, 'Location', None),
            categories=list(item.Categories.split(',')) if getattr(item, 'Categories', None) else []
        )