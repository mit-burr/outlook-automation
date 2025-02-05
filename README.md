# Outlook Calendar Automation

A Python application for automating Outlook calendar tasks, specifically focused on tracking meeting hours and categorizing time spent.

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

2. Run the calendar check script:
```bash
python test_outlook_connection.py
```

This will:
- Connect to your Outlook calendar
- Display meetings for the current week
- Calculate total meeting hours
- Show detailed meeting information including duration and organizer

If you encounter any issues, the application will provide guidance about:
- Outlook version compatibility
- Connection requirements
- Necessary permissions

## Project Structure

Services and files which exist out of this structure are relics of a template, or planned future functionality.

```
+-- services/
│   +-- outlook_service/    # Core Outlook interaction service
│       +-- service.py      # Main service implementation
│       +-- tests/          # Service tests
+-- shared/
│   +-- logger.py          # Logging utility
+-- test_outlook_connection.py  # Main test script
```

## Development

- Run tests: `python -m pytest`
- Main service file: `services/outlook_service/service.py`
- Logger utility: `shared/logger.py`

## Troubleshooting

1. No meetings found:
   - Verify you're using Outlook Classic/Desktop version
   - Ensure Outlook is running and you're logged in
   - Check if you have calendar items visible in Outlook itself

2. Connection issues:
   - Try running the script as administrator
   - Check Outlook security settings
   - Verify you have necessary permissions

## Future Enhancements

Planned features:
- Time categorization for meetings
- Reporting capabilities
- Export to timecard systems

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request