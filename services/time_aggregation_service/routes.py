from fastapi import APIRouter, HTTPException
from datetime import datetime

from .models import RequestModel, ResponseModel
from .services import ServiceTemplate

router = APIRouter()
service = ServiceTemplate()

@router.post("/process", response_model=ResponseModel)
async def process_request(request: RequestModel):
    """Template endpoint"""
    try:
        return await service.process_request(request)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
