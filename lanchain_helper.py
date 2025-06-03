import os
import requests
from msal import ConfidentialClientApplication
from langchain_community.vectorstores import FAISS
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.schema.document import Document

# 🔐 Microsoft App Credentials (App Registration)
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# 🌐 SharePoint Info
SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST")
SITE_NAME = os.getenv("SITE_NAME")
DOC_LIB_PATH = os.getenv("DOC_LIB_PATH")

# 🔎 Embeddings
EMBEDDINGS_MODEL = "sentence-transformers/all-mpnet-base-v2"
embeddings = HuggingFaceEmbeddings(model_name=EMBEDDINGS_MODEL)


def authenticate():
    """App-only authentication using client credentials (non-interactive)."""
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
    )

    result = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" not in result:
        raise Exception(f"❌ Failed to acquire token: {result.get('error_description')}")

    print("✅ Successfully authenticated via App-Only flow.")
    return result["access_token"]


def fetch_txt_files_from_sharepoint():
    """Download .txt files from SharePoint using Microsoft Graph API."""
    token = authenticate()
    headers = {"Authorization": f"Bearer {token}"}

    try:
        # ➤ Retrieve Site ID
        site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
        site_resp = requests.get(site_url, headers=headers)
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]
        print(f"🔍 Site ID found: {site_id}")

        # ➤ Retrieve Drive ID
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        drive_resp = requests.get(drives_url, headers=headers)
        drive_resp.raise_for_status()
        drive_id = next((d["id"] for d in drive_resp.json()["value"] if d["name"] == "Documents"), None)
        if not drive_id:
            raise Exception("❌ 'Documents' drive not found.")
        print(f"📦 Drive ID found: {drive_id}")

        # ➤ List and fetch .txt files
        encoded_path = DOC_LIB_PATH.replace(" ", "%20")
        files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
        files_resp = requests.get(files_url, headers=headers)
        files_resp.raise_for_status()

        docs = []
        for item in files_resp.json().get("value", []):
            if item["name"].endswith(".txt"):
                content_resp = requests.get(item["@microsoft.graph.downloadUrl"])
                content_resp.raise_for_status()
                docs.append(Document(page_content=content_resp.text, metadata={
                    "source": item["name"],
                    "full_content": content_resp.text
                }))

        print(f"📄 Retrieved {len(docs)} .txt documents from SharePoint:")
        for doc in docs:
            print(f" - {doc.metadata['source']}")

        return docs

    except requests.HTTPError as http_err:
        print(f"❌ HTTP error occurred: {http_err}")
        return []
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return []


def index_documents():
    """Index and store documents locally."""
    print("📥 Beginning indexing of SharePoint documents...")
    documents = fetch_txt_files_from_sharepoint()
    if not documents:
        raise Exception("❌ No .txt files found to index.")

    text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
    
    # Build a map of source name to full content
    source_to_full = {doc.metadata["source"]: doc.metadata["full_content"] for doc in documents}

    # Split the documents into chunks
    chunks = text_splitter.split_documents(documents)

    # Embed full content into each chunk’s metadata
    for chunk in chunks:
        source = chunk.metadata.get("source")
        chunk.metadata["full_content"] = source_to_full.get(source, "")

    vectorstore = FAISS.from_documents(chunks, embeddings)
    vectorstore.save_local("./vector_index")
    print("✅ Indexing complete and stored locally.")


def get_similar_answer_from_documents(query: str):
    """Query local FAISS index for relevant answers."""
    print(f"🧐 Querying vector index for: '{query}'")

    if not os.path.exists("./vector_index"):
        print("⚠️ Vector index missing. Initiating indexing...")
        index_documents()

    try:
        vectorstore = FAISS.load_local("./vector_index", embeddings, allow_dangerous_deserialization=True)
    except Exception as e:
        print(f"⚠️ Error loading vector index: {e}. Rebuilding index...")
        index_documents()
        vectorstore = FAISS.load_local("./vector_index", embeddings, allow_dangerous_deserialization=True)

    docs_with_scores = vectorstore.similarity_search_with_score(query, k=3)

    if not docs_with_scores:
        return "❓ Apologies, I couldn't find relevant information.", None

    for doc, score in docs_with_scores:
        if score < 0.6:
            return f"❌ Sorry, we do not offer information on '{query.lower()}' at this time.", None
        full_content = doc.metadata.get("full_content", doc.page_content)
        return f"🔍 **Answer:** {doc.page_content}", full_content

    return f"❌ Sorry, we do not offer information on '{query.lower()}' at this time.", None
