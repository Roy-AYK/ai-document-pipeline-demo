# ai-document-pipeline-demo

Mock AI document extract → review → confirm → SQLite. Optional OCR → Ollama (`llama3.2`) path.

## Run

```bash
cd backend
python -m venv .venv
# Windows: .venv\Scripts\activate
pip install -r ../requirements.txt
python ../scripts/generate_sample_data.py
uvicorn main:app --reload --port 8000
```

Optional OCR+Ollama: `ollama pull llama3.2`

Open http://localhost:8000/

Schema builder: http://localhost:8000/tools/schema-builder/  
Clean export: http://localhost:8000/tools/clean-export/
