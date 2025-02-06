# Outlook Calendar Automation

A Python application for analyzing Outlook calendar data, calculating meeting hours, and categorizing meetings by type. The application provides daily and weekly summaries of time spent in different types of meetings.

## Requirements

- Python 3.11 or higher
- Microsoft Outlook (Classic/Desktop version)
- Poetry for dependency management

## Important Note About Outlook Compatibility

This application works with Outlook Classic/Desktop version. It does not support the "new" Outlook (Microsoft 365 web-based version). If you're using the new Outlook, you'll need to switch to the Classic version for this application to work.

## Setup

1. Install Poetry if you haven't already:
```bash
curl -sSL https://install.python-poetry.org | python3 -
```

2. Clone this repository and navigate to it:
```bash
git clone [repository-url]
cd outlook-automation
```

3. Install dependencies:
```bash
poetry install
```

4. Activate the virtual environment:
```bash
poetry shell
```

## Usage

1. Ensure Outlook (Classic/Desktop version) is running and you're logged in.

2. Run the application:
```bash
python cli.py
```

The application provides an interactive menu with the following options:
1. Check this week's meetings
2. Check next week's meetings
3. Check last week's meetings
4. Generate detailed report

For each time period, the application will show:
- Daily breakdown of meetings by category
- Total time spent in each category per day
- Weekly summary of time across all categories

### Meeting Categories

Meetings are automatically categorized into:
- Team/Staff: Small group meetings, standups, 1:1s
- Department: Larger organizational meetings, planning sessions
- Company-Wide: All-hands meetings, town halls
- Onboarding: New hire and training related meetings

### Time Calculations

- Meeting durations are rounded up to the nearest 30-minute interval
- Daily totals show time spent by category
- Weekly summaries provide an overview of total time in each category

## Project Structure

```
+-- services/
│   +-- outlook_service/          # Core Outlook interaction service
│   │   +-- service.py           # Main Outlook service
│   │   +-- models.py            # Meeting data models
│   │   +-- tests/               # Service tests
│   +-- categorization_service/   # Meeting categorization
│   │   +-- service.py           # Categorization logic
+-- shared/
│   +-- logger.py                # Logging utility
+-- cli.py                       # Main application entry point
```

## Development

- Run tests: `python -m pytest`
- Add new categories: Update `categorization_service/service.py`
- Modify time calculations: Update `outlook_service/models.py`

## Troubleshooting

1. No meetings found:
   - Verify you're using Outlook Classic/Desktop version
   - Ensure Outlook is running and you're logged in
   - Check if you have calendar items visible in Outlook itself

2. Connection issues:
   - Try running the script as administrator
   - Check Outlook security settings
   - Verify you have necessary permissions

3. Incorrect categorization:
   - Check category keywords in `categorization_service/service.py`
   - Meeting titles and organizers are used for categorization

## Future Enhancements

Planned features:
- Export capabilities for timecard systems
- Custom category definitions
- Historical reporting and trends
- Meeting pattern analysis

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request