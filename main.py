import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from router import router

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("main")


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("SharePoint Integration API starting up")
    yield
    logger.info("SharePoint Integration API shutting down")


app = FastAPI(
    title="SharePoint Integration API",
    description="Standalone Microsoft Graph SharePoint integration (read + upload).",
    version="1.0.0",
    lifespan=lifespan,
)

app.include_router(router)


@app.get("/")
async def root():
    return {
        "service": "sharepoint-integration",
        "docs": "/docs",
        "router": "/sharepoint",
    }
