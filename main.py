from fastapi import FastAPI

app = FastAPI()

from routes.proposal_router import proposal_router

app.include_router(proposal_router)

