import pytest
from datetime import datetime
from fastapi.testclient import TestClient

from ..main import app
from ..models import RequestModel, ResponseModel
from ..services import ServiceTemplate

client = TestClient(app)

def test_process_request(sample_request_data):
    """Test the process endpoint"""
    response = client.post("/api/v1/process", json=sample_request_data)
    assert response.status_code == 200
    data = response.json()
    assert "id" in data
    assert "processed_at" in data
    assert "result" in data
    assert "status" in data

@pytest.mark.asyncio
async def test_service_template():
    """Test the service class"""
    service = ServiceTemplate()
    request = RequestModel(
        id="test-id",
        timestamp=datetime.now(),
        data={"test": "data"}
    )
    
    response = await service.process_request(request)
    assert isinstance(response, ResponseModel)
    assert response.status == "success"
