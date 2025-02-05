# services/cli_service/service.py
from datetime import datetime, timedelta
from prompt_toolkit import prompt
from prompt_toolkit.completion import WordCompleter
from services.outlook_service.service import OutlookService
from services.outlook_service.models import Meeting
from shared.logger import logger
import win32com.client
import pythoncom
from collections import defaultdict
import pytz

class CLIService:
    def __init__(self):
        self.outlook = None
        self.choices = {
            '1': ('Check this week\'s meetings', self.check_current_week),
            '2': ('Check next week\'s meetings', self.check_next_week),
            '3': ('Check last week\'s meetings', self.check_last_week),
            '4': ('Generate meeting report', self.generate_report),
            'q': ('Quit', self.quit_program)
        }

    def display_menu(self):
        """Display the main menu and handle user input."""
        while True:
            logger.info("\nOutlook Calendar Automation", "start")
            logger.info("What would you like to do?")
            
            # Display menu options
            for key, (description, _) in self.choices.items():
                logger.info(f"{key}. {description}")
            
            # Get user choice
            choice = prompt("\nEnter your choice: ").lower().strip()
            
            if choice in self.choices:
                func = self.choices[choice][1]
                func()
                if choice == 'q':
                    break
            else:
                logger.error("Invalid choice. Please try again.")

    def check_meetings(self, week_offset: int, week_name: str):
        """Check meetings for a specific week offset."""
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Connect to Outlook
            logger.info("Connecting to Outlook...", "start")
            outlook = win32com.client.Dispatch('Outlook.Application')
            namespace = outlook.GetNamespace('MAPI')
            
            # Log Outlook version
            logger.info(f"Connected to Outlook version: {outlook.Version}")
            
            # Get calendar folder
            calendar = namespace.GetDefaultFolder(9)
            items = calendar.Items
            
            if not items or items.Count == 0:
                logger.warn("No calendar items found.")
                return
            
            # Get target week's date range with timezone awareness
            local_tz = pytz.timezone('America/Chicago')  # Adjust timezone if needed
            today = datetime.now(local_tz)
            start_of_week = (today - timedelta(days=today.weekday()) + timedelta(weeks=week_offset))
            end_of_week = start_of_week + timedelta(days=6)
            
            # Format dates for logging
            start_date = start_of_week.strftime('%m/%d/%Y')
            end_date = end_of_week.strftime('%m/%d/%Y')
            
            logger.info(f"Looking for meetings between {start_date} and {end_date}")
            
            # Include recurrences and sort
            items.IncludeRecurrences = True
            items.Sort("[Start]")
            weekly_meetings = []
            
            # Create naive datetime objects for comparison
            start_naive = start_of_week.replace(tzinfo=None, hour=0, minute=0, second=0, microsecond=0)
            end_naive = end_of_week.replace(tzinfo=None, hour=23, minute=59, second=59, microsecond=999999)
            
            for item in items:
                try:
                    # Get meeting start time as naive datetime for comparison
                    meeting_start = item.Start
                    if isinstance(meeting_start, str):
                        meeting_start = datetime.strptime(meeting_start, '%Y-%m-%d %H:%M')
                    elif hasattr(meeting_start, 'tzinfo') and meeting_start.tzinfo:
                        meeting_start = meeting_start.replace(tzinfo=None)
                    
                    # Compare naive datetimes (inclusive)
                    if start_naive <= meeting_start <= end_naive:
                        weekly_meetings.append(item)
                except Exception as e:
                    logger.error(f"Error processing meeting: {str(e)}")
                    continue
            
            if not weekly_meetings:
                logger.info(f"No meetings found for {week_name} ({start_date} to {end_date})")
                return
            
            # Group meetings by series/unique meeting
            meeting_groups = defaultdict(list)
            series_groups = defaultdict(list)
            meeting_totals = defaultdict(int)
            
            for item in weekly_meetings:
                try:
                    meeting = Meeting.from_outlook_item(item)
                    # Group by subject for summary
                    series_groups[meeting.subject].append(meeting)
                    
                except Exception as e:
                    logger.error(f"Error accessing meeting details: {str(e)}")
            
            # Log meetings grouped by series
            logger.success(f"Found {len(weekly_meetings)} total meetings for {week_name}:")
            
            for subject, meetings in series_groups.items():
                if len(meetings) > 1:
                    recurring_note = " (Recurring)" if meetings[0].is_recurring else " (Multiple instances)"
                    logger.list(f"📅 {subject}{recurring_note}", 
                              [meeting.display_dict for meeting in sorted(meetings, key=lambda x: x.start_time)])
                    total_hours = sum(m.duration for m in meetings) / 60
                    logger.info(f"Total time for '{subject}': {total_hours:.2f} hours")
                else:
                    logger.list(f"📅 {subject}", [meetings[0].display_dict])
            
            # Calculate totals
            total_minutes = sum(meeting.duration for meetings in series_groups.values() for meeting in meetings)
            total_hours = total_minutes / 60
            unique_series = len(set(meeting.series_id for meetings in series_groups.values() for meeting in meetings))
            unique_subjects = len(series_groups)
            
            # Log summary
            logger.success(f"\nSummary:")
            logger.list("Totals", [{
                "unique_meeting_subjects": unique_subjects,
                "unique_meeting_series": unique_series,
                "total_meeting_instances": len(weekly_meetings),
                "total_hours": f"{total_hours:.2f}"
            }])

        except Exception as e:
            logger.error(f"Failed to access Outlook calendar: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
            input("\nPress Enter to continue...")

    def check_current_week(self):
        """Check current week's meetings."""
        self.check_meetings(0, "this week")

    def check_next_week(self):
        """Check next week's meetings."""
        self.check_meetings(1, "next week")

    def check_last_week(self):
        """Check last week's meetings."""
        self.check_meetings(-1, "last week")

    def generate_report(self):
        """Generate a meeting report."""
        logger.info("Report generation coming soon!")
        input("\nPress Enter to continue...")

    def quit_program(self):
        """Exit the program."""
        logger.info("Thank you for using Outlook Calendar Automation!", "end")