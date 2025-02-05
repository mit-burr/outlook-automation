from datetime import datetime
from typing import Dict

from .models import RequestModel, ResponseModel

class ServiceTemplate:
    """Template service class"""
    
    async def process_request(self, request: RequestModel) -> ResponseModel:
        """Template processing method"""
        return ResponseModel(
            id=request.id,
            processed_at=datetime.now(),
            result={"processed": True},
            status="success"
        )
