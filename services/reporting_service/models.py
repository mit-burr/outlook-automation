from datetime import datetime
from typing import Optional
from pydantic import BaseModel

class RequestModel(BaseModel):
    """Template request model"""
    id: str
    timestamp: datetime
    data: dict

class ResponseModel(BaseModel):
    """Template response model"""
    id: str
    processed_at: datetime
    result: dict
    status: str
