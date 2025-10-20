#!/usr/bin/env python3
"""
Minimal test app to verify Render deployment works
"""

from fastapi import FastAPI
import os

app = FastAPI()

@app.get("/")
def root():
    return {"message": "Test app is running", "status": "ok"}

@app.get("/health")
def health():
    return {"ok": True}

@app.get("/debug/gmail_status")
def gmail_status():
    return {
        "status": "test",
        "gmail_connected": False,
        "message": "This is a test endpoint"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)



