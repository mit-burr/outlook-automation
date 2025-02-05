# services/outlook_service/tests/test_outlook_service.py
import pytest
from datetime import datetime, timedelta
from unittest.mock import Mock, patch
from ..service import OutlookService, CalendarEvent

@pytest.fixture
def mock_outlook():
    with patch('win32com.client.Dispatch') as mock_dispatch:
        outlook_mock = Mock()
        namespace_mock = Mock()
        calendar_mock = Mock()
        
        mock_dispatch.return_value = outlook_mock
        outlook_mock.GetNamespace.return_value = namespace_mock
        namespace_mock.GetDefaultFolder.return_value = calendar_mock
        
        yield {
            'outlook': outlook_mock,
            'namespace': namespace_mock,
            'calendar': calendar_mock
        }

@pytest.fixture
def outlook_service(mock_outlook):
    return OutlookService()

def test_connect_to_outlook(mock_outlook):
    service = OutlookService()
    assert service.outlook is not None
    assert service.namespace is not None
    mock_outlook['outlook'].GetNamespace.assert_called_once_with('MAPI')

def test_get_calendar(outlook_service, mock_outlook):
    calendar = outlook_service.get_calendar()
    assert calendar == mock_outlook['calendar']
    mock_outlook['namespace'].GetDefaultFolder.assert_called_once_with(9)

def test_get_calendar_events(outlook_service, mock_outlook):
    # Setup mock calendar items
    items_mock = Mock()
    restricted_items = Mock()
    
    mock_outlook['calendar'].Items = items_mock
    items_mock.Restrict.return_value = restricted_items
    
    # Create mock appointment
    appointment = Mock()
    appointment.Subject = "Test Meeting"
    appointment.Start = datetime.now()
    appointment.End = datetime.now() + timedelta(hours=1)
    appointment.Duration = 60
    appointment.Categories = "Category1,Category2"
    appointment.Body = "Test Body"
    appointment.Organizer = "Test Organizer"
    
    restricted_items.__iter__ = lambda self: iter([appointment])
    
    # Test
    start_date = datetime.now()
    end_date = start_date + timedelta(days=1)
    events = outlook_service.get_calendar_events(start_date, end_date)
    
    assert len(events) == 1
    assert events[0].subject == "Test Meeting"
    assert events[0].categories == ["Category1", "Category2"]
    assert events[0].duration == 60

def test_calculate_total_meeting_hours():
    service = OutlookService()
    
    # Create test events
    events = [
        CalendarEvent(
            subject="Meeting 1",
            start_time=datetime.now(),
            end_time=datetime.now() + timedelta(hours=1),
            duration=60,
            categories=[],
            body="",
            organizer=""
        ),
        CalendarEvent(
            subject="Meeting 2",
            start_time=datetime.now(),
            end_time=datetime.now() + timedelta(hours=2),
            duration=120,
            categories=[],
            body="",
            organizer=""
        )
    ]
    
    total_hours = service.calculate_total_meeting_hours(events)
    assert total_hours == 3.0