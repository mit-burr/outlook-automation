# test_outlook_connection.py
from services.outlook_service.service import OutlookService
from shared.logger import logger
from datetime import datetime, timedelta
import win32com.client
import pythoncom

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
        calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        items = calendar.Items
        
        if not items or items.Count == 0:
            logger.warn("""
No calendar items found. This might be because:
1. You're using the new version of Outlook (Microsoft 365) which doesn't support automation
2. You don't have any meetings scheduled
3. There might be permission issues

Solution: Try using Outlook Classic/Desktop version instead of the new Outlook.""")
            return
        
        # Get this week's date range
        today = datetime.now()
        start_of_week = today - timedelta(days=today.weekday())
        end_of_week = start_of_week + timedelta(days=6)
        
        # Format dates for restriction
        start_date = start_of_week.strftime('%m/%d/%Y')
        end_date = end_of_week.strftime('%m/%d/%Y')
        
        # Create restriction for this week
        restriction = f"[Start] >= '{start_date}' AND [End] <= '{end_date}'"
        weekly_items = items.Restrict(restriction)
        
        if weekly_items.Count == 0:
            logger.info(f"No meetings found for the week of {start_date} to {end_date}")
            return
        
        # Sort items by start time
        weekly_items.Sort("[Start]")
        
        # Log meetings found
        logger.success(f"Found {weekly_items.Count} meetings this week:")
        
        for item in weekly_items:
            try:
                meeting_info = {
                    "subject": item.Subject,
                    "start": item.Start.strftime("%Y-%m-%d %H:%M"),
                    "end": item.End.strftime("%Y-%m-%d %H:%M"),
                    "duration": f"{item.Duration} minutes",
                    "organizer": item.Organizer
                }
                logger.list("Meeting", [meeting_info])
            except Exception as e:
                logger.error(f"Error accessing meeting details: {str(e)}")
        
        # Calculate total hours
        total_minutes = sum(item.Duration for item in weekly_items)
        total_hours = total_minutes / 60
        logger.success(f"\nTotal meeting hours this week: {total_hours:.2f}")

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