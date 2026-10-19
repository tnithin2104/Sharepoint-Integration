import json
import logging
import os
from contextlib import asynccontextmanager
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, Request, HTTPException, status
from fastapi.responses import PlainTextResponse, JSONResponse
from pydantic import BaseModel, ValidationError

from router import router
from sharepoint_integration.config import has_credentials
from sharepoint_integration.delta_sync import shutdown_delta_sync

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("main")


class ResourceData(BaseModel):
    """Resource-specific data in the notification."""
    # This is intentionally left flexible to accommodate different resource types
    pass


class NotificationItem(BaseModel):
    """Single notification item in the value array."""
    subscriptionId: str
    changeType: str
    resource: str
    resourceData: Dict[str, Any]
    clientState: Optional[str] = None
    tenantId: Optional[str] = None
    subscriptionExpirationDateTime: Optional[str] = None
    lifecycleEvent: Optional[str] = None


class NotificationPayload(BaseModel):
    """Microsoft Graph notification payload."""
    value: List[NotificationItem]


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("SharePoint Integration API starting up")
    yield
    logger.info("SharePoint Integration API shutting down")
    shutdown_delta_sync()


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


@app.post("/webhooks/sharepoint")
async def sharepoint_webhook(request: Request):
    """
    Microsoft Graph SharePoint webhook endpoint.
    
    Handles:
    1. Validation token handshake (GET-like behavior via validationToken query param)
    2. Change notification processing
    """
    # Handle validation token handshake
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        logger.info("Received validation token request")
        return PlainTextResponse(
            content=validation_token,
            status_code=status.HTTP_200_OK,
            media_type="text/plain"
        )
    
    # Handle actual notification
    try:
        body = await request.json()
        payload = NotificationPayload(**body)
        # Validate clientState if configured
        expected_client_state = os.getenv("SHAREPOINT_WEBHOOK_CLIENT_STATE")
        if expected_client_state is not None:
            for i, notification in enumerate(payload.value):
                client_state = notification.clientState
                if client_state is None:
                    logger.warning(
                        f"Notification {i} missing clientState (subscriptionId: {notification.subscriptionId})"
                    )
                    raise HTTPException(
                        status_code=status.HTTP_400_BAD_REQUEST,
                        detail="Missing clientState in notification"
                    )
                
                # Constant-time comparison to prevent timing attacks
                if not _constant_time_compare(client_state, expected_client_state):
                    logger.warning(
                        f"Invalid clientState for subscriptionId: {notification.subscriptionId}"
                    )
                    raise HTTPException(
                        status_code=status.HTTP_401_UNAUTHORIZED,
                        detail="Invalid clientState"
                    )
        
        # Process notifications
        await _process_notifications(payload.value)
        
        # Return success response
        return JSONResponse(
            status_code=status.HTTP_202_ACCEPTED,
            content={"status": "accepted"}
        )
        
    except Exception as e:
        logger.error(f"Error processing webhook: {str(e)}")
        # Return 400 for client errors, 500 for server errors
        if isinstance(e, HTTPException):
            raise e
        elif (
            isinstance(e, json.JSONDecodeError)
            or isinstance(e, ValidationError)
            or "Validation error" in str(e)
            or "json" in str(e).lower()
        ):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Invalid notification payload"
            )
        else:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Internal server error"
            )


def _constant_time_compare(a: str, b: str) -> bool:
    """Constant-time string comparison to prevent timing attacks."""
    if len(a) != len(b):
        return False
    result = 0
    for x, y in zip(a, b):
        result |= ord(x) ^ ord(y)
    return result == 0


async def _process_notifications(notifications: List[NotificationItem]):
    """Handle incoming Graph change notifications.
    Notifications are only *signals* that something changed in the
    library — they are not the list of changes. Each notification
    schedules a background Microsoft Graph Delta API poll (coalesced
    across duplicate/concurrent notifications) which detects the
    actual new/modified/deleted items.
    """
    change_types = {}
    lifecycle_events = 0

    for notification in notifications:
        # Count by change type
        change_type = notification.changeType
        change_types[change_type] = change_types.get(change_type, 0) + 1

        # Count lifecycle events
        if notification.lifecycleEvent:
            lifecycle_events += 1

        # Log notification details (excluding sensitive data)
        logger.info(
            f"Received notification: subscriptionId={notification.subscriptionId}, "
            f"changeType={notification.changeType}, resource={notification.resource}"
        )

        if notification.lifecycleEvent:
            logger.info(
                f"Lifecycle event: {notification.lifecycleEvent} for subscriptionId={notification.subscriptionId}"
            )

    # Log summary
    logger.info(
        f"Processed {len(notifications)} notifications: "
        f"change types={change_types}, lifecycle events={lifecycle_events}"
    )

    # Notifications trigger a Delta API poll in the background; the
    # webhook response is acknowledged immediately (no lengthy work here).
    if not has_credentials():
        logger.info(
            "SharePoint credentials are not configured; "
            "skipping delta sync trigger"
        )
        return

    delta_sync = _get_delta_sync()
    delta_sync.notify()


def _get_delta_sync():
    """Lazy accessor for the app-wide delta sync service.

    Isolated in a function so tests can patch ``main._get_delta_sync``.
    """
    from sharepoint_integration.delta_sync import get_delta_sync

    return get_delta_sync()