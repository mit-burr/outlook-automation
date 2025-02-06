# services/cli_service/service.py
from datetime import datetime, timedelta
from prompt_toolkit import prompt
from prompt_toolkit.completion import WordCompleter
from services.outlook_service.service import OutlookService
from services.outlook_service.models import Meeting
from services.categorization_service.services import CategorizationService, MeetingCategory
from shared.logger import logger
import win32com.client
import pythoncom
from collections import defaultdict
import pytz
from typing import List, Dict

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

    def format_duration(self, minutes: int) -> str:
        """Convert minutes to a readable duration string."""
        hours = minutes // 60
        remaining_minutes = minutes % 60
        if hours and remaining_minutes:
            return f"{hours}h{remaining_minutes}m"
        elif hours:
            return f"{hours}h"
        else:
            return f"{remaining_minutes}m"

    def get_daily_summary(self, meetings: List[Meeting]) -> Dict[str, List[Meeting]]:
        """Group meetings by day and calculate daily totals."""
        daily_meetings = defaultdict(list)
        for meeting in meetings:
            daily_meetings[meeting.weekday].append(meeting)
        return daily_meetings

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
                    meeting_start = item.Start
                    if isinstance(meeting_start, str):
                        meeting_start = datetime.strptime(meeting_start, '%Y-%m-%d %H:%M')
                    elif hasattr(meeting_start, 'tzinfo') and meeting_start.tzinfo:
                        meeting_start = meeting_start.replace(tzinfo=None)
                    
                    if start_naive <= meeting_start <= end_naive:
                        weekly_meetings.append(item)
                except Exception as e:
                    logger.error(f"Error processing meeting: {str(e)}")
                    continue

            if not weekly_meetings:
                logger.info(f"No meetings found for {week_name} ({start_date} to {end_date})")
                return

            # Convert to Meeting objects
            meetings = [Meeting.from_outlook_item(m) for m in weekly_meetings]
            
            # Display daily summary
            self.display_daily_summary(meetings)
            
            # Offer to adjust meetings
            # Future feature implementation
            # logger.info("\nWould you like to adjust any meeting times? (y/n)")
            # if prompt().lower().strip() == 'y':
            #     self.adjust_meetings(meetings)

        except Exception as e:
            logger.error(f"Failed to access Outlook calendar: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

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

    def adjust_meetings(self, meetings: List[Meeting]):
        """Interactive menu for adjusting meeting times."""
        while True:
            logger.info("\nMeeting Adjustment Options:")
            logger.info("1. Adjust specific meeting duration")
            logger.info("2. Scale all meeting times")
            logger.info("3. Adjust recurring meeting series")
            logger.info("q. Return to main menu")

            choice = prompt("\nEnter your choice: ").lower().strip()

            made_changes = False
            if choice == '1':
                made_changes = self.adjust_specific_meeting(meetings)
            elif choice == '2':
                made_changes = self.scale_all_meetings(meetings)
            elif choice == '3':
                made_changes = self.adjust_recurring_series(meetings)
            elif choice == 'q':
                break
            else:
                logger.error("Invalid choice")
                continue

            # Show updated summary if changes were made
            if made_changes:
                logger.success("\nUpdated Daily Summary:")
                self.display_daily_summary(meetings)

    def adjust_specific_meeting(self, meetings: List[Meeting]) -> bool:
        """Adjust duration of a specific meeting instance."""
        # List all meetings
        for i, meeting in enumerate(meetings, 1):
            logger.info(f"{i}. {meeting.subject} ({meeting.start_time.strftime('%A %H:%M')})")

        try:
            idx = int(prompt("\nSelect meeting number: ")) - 1
            if 0 <= idx < len(meetings):
                adjustment = int(prompt("Enter time adjustment in minutes (multiple of 30, can be negative): "))
                if adjustment % 30 == 0:
                    meetings[idx].duration = max(0, meetings[idx].duration + adjustment)
                    logger.success(f"Adjusted meeting duration to {self.format_duration(meetings[idx].rounded_duration)}")
                    return True
                else:
                    logger.error("Adjustment must be in 30-minute intervals")
            return False
        except ValueError:
            logger.error("Invalid input")
            return False

    def scale_all_meetings(self, meetings: List[Meeting]) -> bool:
        """Scale all meeting durations by a factor."""
        try:
            factor = float(prompt("Enter scaling factor (e.g., 1.25 for 25% increase): "))
            for meeting in meetings:
                meeting.duration = int(meeting.duration * factor)
            logger.success(f"Scaled all meeting durations by {factor}x")
            return True
        except ValueError:
            logger.error("Invalid input")
            return False

    def adjust_recurring_series(self, meetings: List[Meeting]) -> bool:
        """Adjust all instances of a recurring meeting series."""
        # Get unique meeting subjects
        series = set(meeting.subject for meeting in meetings if meeting.is_recurring)
        
        if not series:
            logger.warn("No recurring meeting series found")
            return False

        for i, subject in enumerate(series, 1):
            logger.info(f"{i}. {subject}")

        try:
            idx = int(prompt("\nSelect series number: ")) - 1
            if 0 <= idx < len(series):
                series_name = list(series)[idx]
                factor = float(prompt("Enter scaling factor (e.g., 1.25 for 25% increase): "))
                
                for meeting in meetings:
                    if meeting.subject == series_name:
                        meeting.duration = int(meeting.duration * factor)
                
                logger.success(f"Scaled all instances of '{series_name}' by {factor}x")
                return True
            return False
        except ValueError:
            logger.error("Invalid input")
            return False

    def display_daily_summary(self, meetings: List[Meeting]):
        """Display summary of meetings grouped by day and category."""
        categorization = CategorizationService()
        daily_meetings = self.get_daily_summary(meetings)
        
        logger.success("\nDaily Summary:")
        
        for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
            if day in daily_meetings:
                day_meetings = daily_meetings[day]
                
                # Get categories for this day's meetings
                day_categorized = categorization.categorize_meetings(day_meetings)
                
                # Calculate totals by category
                category_totals = {
                    str(cat.value): sum(m.rounded_duration for m in cat_meetings)  # Use string value of category
                    for cat, cat_meetings in day_categorized.items()
                    if cat_meetings  # Only include categories with meetings
                }
                
                # Display day totals by category
                logger.info(f"\n{day}:")
                for category, total_minutes in category_totals.items():
                    if total_minutes > 0:
                        logger.info(f"  {category}: {self.format_duration(total_minutes)}")
                
                # Display day total
                total_minutes = sum(category_totals.values())
                logger.info(f"  Total: {self.format_duration(total_minutes)}")
                
        # Calculate and display week totals by category
        categorized_meetings = categorization.categorize_meetings(meetings)
        logger.success("\nWeek Totals by Category:")
        for category, cat_meetings in categorized_meetings.items():
            if cat_meetings:  # Only show categories with meetings
                total_minutes = sum(m.rounded_duration for m in cat_meetings)
                logger.info(f"  {category.value}: {self.format_duration(total_minutes)}")  # Use category.value
        
        # Week total
        week_total = sum(m.rounded_duration for m in meetings)
        logger.info(f"  Total: {self.format_duration(week_total)}")

    def generate_report(self):
        """Generate a detailed report showing how meetings were categorized."""
        logger.start_section("Detailed Meeting Report")
        
        try:
            pythoncom.CoInitialize()
            
            # Get current week's meetings
            outlook = win32com.client.Dispatch('Outlook.Application')
            namespace = outlook.GetNamespace('MAPI')
            calendar = namespace.GetDefaultFolder(9)
            items = calendar.Items
            
            if not items:
                logger.warn("No meetings found")
                return
                
            items.IncludeRecurrences = True
            items.Sort("[Start]")
            
            today = datetime.now()
            start_of_week = today - timedelta(days=today.weekday())
            end_of_week = start_of_week + timedelta(days=6)
            
            start_naive = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
            end_naive = end_of_week.replace(hour=23, minute=59, second=59, microsecond=999999)
            
            meetings = []
            for item in items:
                meeting_start = item.Start
                if isinstance(meeting_start, str):
                    meeting_start = datetime.strptime(meeting_start, '%Y-%m-%d %H:%M')
                if hasattr(meeting_start, 'tzinfo'):
                    meeting_start = meeting_start.replace(tzinfo=None)
                
                if start_naive <= meeting_start <= end_naive:
                    meetings.append(Meeting.from_outlook_item(item))
            
            # Show detailed categorization
            categorization = CategorizationService()
            categorized_meetings = categorization.categorize_meetings(meetings)
            
            logger.info("\nDetailed Meeting Categorization:")
            for category, cat_meetings in categorized_meetings.items():
                if cat_meetings:
                    logger.info(f"\n{category}:")
                    for meeting in sorted(cat_meetings, key=lambda m: m.start_time):
                        logger.list(f"{meeting.start_time.strftime('%A %H:%M')}", [{
                            "subject": meeting.subject,
                            "duration": self.format_duration(meeting.rounded_duration),
                            "organizer": meeting.organizer.split(',')[0]
                        }])
            
        except Exception as e:
            logger.error(f"Error generating report: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
            input("\nPress Enter to continue...")