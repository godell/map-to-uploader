from fastapi import FastAPI, APIRouter
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List
import uuid
from datetime import datetime, timezone


ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# MongoDB connection
mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

# Create the main app without a prefix
app = FastAPI()

# Create a router with the /api prefix
api_router = APIRouter(prefix="/api")


# Define Models
class StatusCheck(BaseModel):
    model_config = ConfigDict(extra="ignore")  # Ignore MongoDB's _id field
    
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    client_name: str
    timestamp: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class StatusCheckCreate(BaseModel):
    client_name: str


# ── Pick Ticket / Batch Print History Models ──
class PickTicketRecord(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    batch_code: str  # e.g. PTF-Batch-1-260807
    batch_number: int
    batch_size: int
    picker_count: int
    to_numbers: List[str]
    printed_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class PickTicketCreate(BaseModel):
    batch_code: str
    batch_number: int
    batch_size: int
    picker_count: int
    to_numbers: List[str]

class PickTicketCheckRequest(BaseModel):
    to_numbers: List[str]

class PickTicketPreviousPrint(BaseModel):
    to_number: str
    batch_code: str
    printed_at: str

# Add your routes to the router instead of directly to app
@api_router.get("/")
async def root():
    return {"message": "Hello World"}

@api_router.post("/status", response_model=StatusCheck)
async def create_status_check(input: StatusCheckCreate):
    status_dict = input.model_dump()
    status_obj = StatusCheck(**status_dict)
    
    # Convert to dict and serialize datetime to ISO string for MongoDB
    doc = status_obj.model_dump()
    doc['timestamp'] = doc['timestamp'].isoformat()
    
    _ = await db.status_checks.insert_one(doc)
    return status_obj

@api_router.get("/status", response_model=List[StatusCheck])
async def get_status_checks():
    # Exclude MongoDB's _id field from the query results
    status_checks = await db.status_checks.find({}, {"_id": 0}).to_list(1000)
    
    # Convert ISO string timestamps back to datetime objects
    for check in status_checks:
        if isinstance(check['timestamp'], str):
            check['timestamp'] = datetime.fromisoformat(check['timestamp'])
    
    return status_checks


# ── Pick Ticket / Batch Print History Endpoints ──
@api_router.post("/pick-tickets/check")
async def check_pick_tickets(payload: PickTicketCheckRequest):
    """
    Given a list of TO numbers, return which ones have been printed before
    (each entry: to_number, batch_code, printed_at ISO string).
    """
    if not payload.to_numbers:
        return {"previous_prints": []}

    cursor = db.pick_tickets.find(
        {"to_numbers": {"$in": payload.to_numbers}},
        {"_id": 0, "batch_code": 1, "printed_at": 1, "to_numbers": 1},
    ).sort("printed_at", -1)

    docs = await cursor.to_list(2000)
    seen: dict = {}
    for d in docs:
        printed_at = d.get("printed_at")
        for to in d.get("to_numbers", []):
            if to in payload.to_numbers and to not in seen:
                seen[to] = {
                    "to_number": to,
                    "batch_code": d.get("batch_code", ""),
                    "printed_at": printed_at if isinstance(printed_at, str) else (printed_at.isoformat() if printed_at else ""),
                }
    return {"previous_prints": list(seen.values())}


@api_router.post("/pick-tickets", response_model=PickTicketRecord)
async def create_pick_ticket(payload: PickTicketCreate):
    record = PickTicketRecord(**payload.model_dump())
    doc = record.model_dump()
    doc["printed_at"] = doc["printed_at"].isoformat()
    await db.pick_tickets.insert_one(doc)
    return record


@api_router.get("/pick-tickets", response_model=List[PickTicketRecord])
async def list_pick_tickets():
    docs = await db.pick_tickets.find({}, {"_id": 0}).sort("printed_at", -1).to_list(500)
    for d in docs:
        if isinstance(d.get("printed_at"), str):
            d["printed_at"] = datetime.fromisoformat(d["printed_at"])
    return docs

# Include the router in the main app
app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()