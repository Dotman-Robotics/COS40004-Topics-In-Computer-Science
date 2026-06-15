# COS40004-Topics-In-Computer-Science
Complete collection for extra projects of Content for COS40004


Documented here are the Testing cases for the initial versions of Project Lateralus. I would have liked for the final version of this project to remain in its own file, but due to time constraints I couldnt move all the files for from this repo to another one

Multi_Agent_System contains the final version of this project: Other files are the earilier versions of some of the modules that have been incorporated into the system

Requirements for the system:
- make sure you have outlook and an outlook email address (Please use Outlook Classic, it functions the best for multiple emails from one system)
- Ollama with the following two models installed
    - Llama 3.2 (ollama pull llama3.2)
    - mxbai-embed-large (ollama pull mxbai-embed-large)
- Setup a virtual environment (python -m venv) and activate it (venv/Scripts/activate.ps1)
- install the following (pip install flask flask-cors pywebview requests beautifulsoup4 lxml ddgs pandas scikit-learn langchain-chroma langchain-ollama langchain-core pywin32 chromadb)

Using the system:
Main : python main.py --account youraddress@outlook.com --port 5000

-- account : Your Email address
-- port : for when you have more than one instance on the system
-- cli : activates the command line interface

File Structure:
├── main.py                  # Entry point
├── app.py                   # Flask REST API
├── planner_agent.py         # NLP request routing with RAG context
├── email_agent.py           # Email composition and sending
├── email_monitor.py         # Background inbox polling
├── email_summarizer.py      # Email and session summarisation
├── calendar_agent.py        # Availability checking and booking
├── negotiation_agent.py     # Multi-round negotiation state machine
├── search_agent.py          # DDG search, scraping, outreach drafting
├── rag_store.py             # ChromaDB RAG memory store
├── vector.py                # Contact lookup with TF-IDF re-ranking
├── ollama_client.py         # Ollama HTTP API wrappers
├── utils.py                 # Shared helpers
├── mock_provider_server.py  # Test website server (localhost:8080)
├── sample-data.csv          # Contact seed data
├── gui/
│   └── index.html           # Desktop GUI (7 panels)
└── mock_provider/           # Static HTML for test website
    ├── index.html
    ├── contact.html
    └── services.html

Generated per run or on first run
  chroma_langchain_db/         # Contact vector store
  chroma_rag_db/               # RAG memory store
  negotiations_<account>_<thread>.json   # Per-negotiation state
  negotiations_<account>_index.json      # Negotiation index
  blacklist_<account>.json               # Per-account blacklist

