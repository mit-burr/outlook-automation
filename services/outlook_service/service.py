# services/outlook_service/outlook_service.py
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Any
import win32com.client
from dataclasses import dataclass

@dataclass
class CalendarEvent:
    """Data class representing a calendar event."""
    subject: str
    start_time: datetime
    end_time: datetime
    duration: timedelta
    categories: List[str]
    body: str
    organizer: str

class OutlookService:
    """Service for interacting with Outlook calendar."""
    
    def __init__(self):
        """Initialize the Outlook service."""
        self.outlook = None
        self.namespace = None
        self._connect_to_outlook()

    def _connect_to_outlook(self) -> None:
        """Establish connection to Outlook application."""
        try:
            self.outlook = win32com.client.Dispatch('Outlook.Application')
            self.namespace = self.outlook.GetNamespace('MAPI')
        except Exception as e:
            raise ConnectionError(f"Failed to connect to Outlook: {str(e)}")

    def get_calendar(self) -> Any:
        """Get the default calendar folder."""
        try:
            return self.namespace.GetDefaultFolder(9)  # 9 represents calendar folder
        except Exception as e:
            raise ValueError(f"Failed to access calendar: {str(e)}")

    def get_calendar_events(self, start_date: datetime, end_date: datetime) -> List[CalendarEvent]:
        """
        Retrieve calendar events for the specified date range.
        
        Args:
            start_date: Start date for the range
            end_date: End date for the range
            
        Returns:
            List of CalendarEvent objects
        """
        calendar = self.get_calendar()
        
        # Create restriction for date range
        restriction = (
            f"[Start] >= '{start_date.strftime('%m/%d/%Y')}' AND "
            f"[End] <= '{end_date.strftime('%m/%d/%Y')}'"
        )
        
        appointments = calendar.Items.Restrict(restriction)
        appointments.Sort("[Start]")
        
        events = []
        for appt in appointments:
            event = CalendarEvent(
                subject=appt.Subject,
                start_time=appt.Start,
                end_time=appt.End,
                duration=appt.Duration,
                categories=list(appt.Categories.split(',')) if appt.Categories else [],
                body=appt.Body,
                organizer=appt.Organizer
            )
            events.append(event)
            
        return events

    def get_previous_week_events(self) -> List[CalendarEvent]:
        """Get all calendar events from the previous week."""
        today = datetime.now()
        start_of_previous_week = today - timedelta(days=today.weekday() + 7)
        start_date = datetime(start_of_previous_week.year, 
                            start_of_previous_week.month,
                            start_of_previous_week.day)
        end_date = start_date + timedelta(days=7)
        
        return self.get_calendar_events(start_date, end_date)

    def get_current_week_events(self) -> List[CalendarEvent]:
        """Get all calendar events for the current week, including future meetings."""
        today = datetime.now()
        start_of_week = today - timedelta(days=today.weekday())
        start_date = datetime(start_of_week.year, 
                            start_of_week.month,
                            start_of_week.day)
        end_date = start_date + timedelta(days=7)
        
        return self.get_calendar_events(start_date, end_date)

    def calculate_total_meeting_hours(self, events: List[CalendarEvent]) -> float:
        """
        Calculate total hours spent in meetings.
        
        Args:
            events: List of CalendarEvent objects
            
        Returns:
            Total hours as float
        """
        total_minutes = sum(event.duration for event in events)
        return total_minutes / 60.0  # Convert to hours