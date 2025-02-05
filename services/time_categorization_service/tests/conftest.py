from datetime import datetime
import pytest
from typing import Dict

@pytest.fixture
def sample_request_data() -> Dict:
    """Fixture for sample request data"""
    return {
        "id": "test-id",
        "timestamp": datetime.now(),
        "data": {"test": "data"}
    }

@pytest.fixture
def mock_service():
    """Fixture for mocked service dependencies"""
    return {}
