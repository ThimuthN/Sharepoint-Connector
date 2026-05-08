"""FastAPI application for Microsoft SharePoint connector."""
import logging
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from routes import router

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(
    title="Microsoft SharePoint Connector",
    description="MVP connector for SharePoint file browsing and download",
    version="1.0.0",
)


# Lazy config loading for middleware
@app.on_event("startup")
async def startup():
    from config import get_config
    try:
        config = get_config()
        logger.info("Configuration loaded successfully")
        
        # Add CORS middleware
        app.add_middleware(
            CORSMiddleware,
            allow_origins=[config.APP_BASE_URL],
            allow_credentials=True,
            allow_methods=["*"],
            allow_headers=["*"],
        )
    except ValueError as e:
        logger.error(f"Configuration error: {e}")
        raise


# Include routers
app.include_router(router, prefix="/api")


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
