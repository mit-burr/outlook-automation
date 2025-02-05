from shared.utils import format_response

def test_format_response():
    """Test the format_response utility"""
    data = {"test": "data"}
    response = format_response(data)
    assert response["success"] is True
    assert response["data"] == data
