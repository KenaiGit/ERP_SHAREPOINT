from fastapi import FastAPI
from pydantic import BaseModel
from lanchain_helper import get_similar_answer_from_documents

app = FastAPI()

class QueryRequest(BaseModel):
    question: str

@app.post("/ask")
def ask_question(request: QueryRequest):
    question = request.question
    print(f"🔍 Received question: {question}")
    response, _ = get_similar_answer_from_documents(question)
    return {"question": question, "response": response}

