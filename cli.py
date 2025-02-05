# cli.py
from services.cli_service.service import CLIService
from shared.logger import logger

def main():
    try:
        cli_service = CLIService()
        cli_service.display_menu()
    except KeyboardInterrupt:
        logger.info("\nThank you for using Outlook Calendar Automation!", "end")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {str(e)}")
        input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()