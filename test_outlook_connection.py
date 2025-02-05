# test_outlook_connection.py
from services.outlook_service.service import OutlookService
from shared.logger import logger
from datetime import datetime, timedelta
import win32com.client
import pythoncom
from collections import defaultdict
import pytz

def get_meeting_key(item):
    """
    Generate a key to identify unique meetings and their recurring instances.
    """
    try:
        # For recurring meetings, use ConversationID to group instances
        if hasattr(item, 'ConversationID') and item.ConversationID:
            return f"{item.Subject}_{item.ConversationID}"
        # For non-recurring meetings, use GlobalObjectID if available
        elif hasattr(item, 'GlobalObjectID') and item.GlobalObjectID:
            return f"{item.Subject}_{item.GlobalObjectID}"
        # Fallback to subject only
        else:
            return item.Subject
    except Exception as e:
        logger.error(f"Error getting meeting key: {str(e)}")
        return item.Subject

def check_outlook_meetings():
    logger.start_section("Outlook Calendar Check")
    
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
        
        # Get next week's date range with timezone awareness
        local_tz = pytz.timezone('America/Chicago')  # Adjust to your timezone
        today = datetime.now(local_tz)
        start_of_next_week = (today - timedelta(days=today.weekday()) + timedelta(days=7))
        end_of_week = start_of_next_week + timedelta(days=6)
        
        # Format dates for logging
        start_date = start_of_next_week.strftime('%m/%d/%Y')
        end_date = end_of_week.strftime('%m/%d/%Y')
        
        logger.info(f"Looking for meetings between {start_date} and {end_date}")
        
        # Include recurrences and sort
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        weekly_meetings = []
        
        # Create naive datetime objects for comparison, using start and end of day
        start_naive = start_of_next_week.replace(tzinfo=None, hour=0, minute=0, second=0, microsecond=0)
        end_naive = end_of_week.replace(tzinfo=None, hour=23, minute=59, second=59, microsecond=999999)
        
        for item in items:
            try:
                # Get meeting start time as naive datetime for comparison
                meeting_start = item.Start
                if isinstance(meeting_start, str):
                    meeting_start = datetime.strptime(meeting_start, '%Y-%m-%d %H:%M')
                elif hasattr(meeting_start, 'tzinfo') and meeting_start.tzinfo:
                    meeting_start = meeting_start.replace(tzinfo=None)
                
                # Debug log the comparison
                # logger.debug(f"Comparing meeting: {meeting_start} with range: {start_naive} to {end_naive}")
                
                # Compare naive datetimes (inclusive)
                if start_naive <= meeting_start <= end_naive:
                    weekly_meetings.append(item)
            except Exception as e:
                logger.error(f"Error processing meeting: {str(e)}")
                continue
        
        if not weekly_meetings:
            logger.info(f"No meetings found for the week of {start_date} to {end_date}")
            return
        
        # Group meetings by series/unique meeting
        meeting_groups = defaultdict(list)
        series_groups = defaultdict(list)
        meeting_totals = defaultdict(int)
        
        for item in weekly_meetings:
            try:
                meeting_info = {
                    "start": item.Start.strftime("%Y-%m-%d %H:%M") if hasattr(item.Start, 'strftime') else str(item.Start),
                    "end": item.End.strftime("%Y-%m-%d %H:%M") if hasattr(item.End, 'strftime') else str(item.End),
                    "duration": item.Duration,
                    "organizer": item.Organizer,
                    "is_recurring": bool(getattr(item, 'RecurrenceState', 0)),
                    "series_id": getattr(item, 'ConversationID', 'N/A')
                }
                
                # Group by meeting key (which includes series information)
                meeting_key = get_meeting_key(item)
                meeting_groups[meeting_key].append(meeting_info)
                meeting_totals[meeting_key] += item.Duration
                
                # Also group by subject for summary
                series_groups[item.Subject].append(meeting_info)
                
            except Exception as e:
                logger.error(f"Error accessing meeting details: {str(e)}")
        
        # Log meetings grouped by series
        logger.success(f"Found {len(weekly_meetings)} total meetings next week:")
        
        for subject, meetings in series_groups.items():
            if len(meetings) > 1:
                recurring_note = " (Recurring)" if meetings[0]["is_recurring"] else " (Multiple instances)"
                logger.list(f"📅 {subject}{recurring_note}", sorted(meetings, key=lambda x: x['start']))
                total_hours = sum(m["duration"] for m in meetings) / 60
                logger.info(f"Total time for '{subject}': {total_hours:.2f} hours")
            else:
                logger.list(f"📅 {subject}", meetings)
        
        # Calculate totals
        total_minutes = sum(item.Duration for item in weekly_meetings)
        total_hours = total_minutes / 60
        unique_series = len(set(get_meeting_key(item) for item in weekly_meetings))
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
        logger.warn("""
If you're seeing connection errors, please check:
1. Make sure you're using Outlook Classic/Desktop version
2. Outlook is running and you're logged in
3. You have necessary permissions to access the calendar""")
    finally:
        pythoncom.CoUninitialize()
    
    logger.end_section("Outlook Calendar Check")

if __name__ == "__main__":
    check_outlook_meetings()