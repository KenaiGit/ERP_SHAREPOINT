import streamlit as st
import re
from lanchain_helper import get_similar_answer_from_documents, fetch_txt_files_from_sharepoint, index_documents
import os

# Detect if running in Streamlit Cloud
IS_CLOUD = st.secrets.get("RUN_ENV", "local") == "cloud"

# Optional imports for local voice features
if not IS_CLOUD:
    import pyttsx3
    import speech_recognition as sr
    import threading

# 🎨 UI Setup
col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image("kenai.png", width=100)
with col2:
    st.markdown("<h1 style='display: flex; align-items: center;'>Oracle ConvoPilot</h1>", unsafe_allow_html=True)

# Initialize session state
if "messages" not in st.session_state:
    st.session_state.messages = []

if "indexed" not in st.session_state:
    st.session_state.indexed = False

# Text-to-speech setup (local only)
if not IS_CLOUD:
    engine = pyttsx3.init()
    engine.setProperty('rate', 150)
    engine.setProperty('volume', 1)
    tts_lock = threading.Lock()

    def speak_text(text):
        def run_speech():
            with tts_lock:
                try:
                    engine.say(text)
                    engine.runAndWait()
                except RuntimeError as e:
                    print(f"⚠️ TTS RuntimeError ignored: {e}")
        threading.Thread(target=run_speech, daemon=True).start()

    def get_voice_input():
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            try:
                st.info("🎤 Listening...")
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
                return recognizer.recognize_google(audio)
            except sr.WaitTimeoutError:
                return "You didn't say anything. Please try again."
            except sr.UnknownValueError:
                return "Sorry, I didn't catch that. Please try again."
            except sr.RequestError:
                return "Could not request results. Check your internet connection."
else:
    def speak_text(text): pass
    def get_voice_input(): return None

# Auto-index docs on first load
if not st.session_state.indexed:
    if not os.path.exists("./vector_index"):
        with st.spinner("🗓 Indexing documents from SharePoint for first use..."):
            try:
                index_documents()
                st.session_state.indexed = True
                st.success("✅ Document index ready!")
            except Exception as e:
                st.error(f"❌ Failed to index documents: {e}")
    else:
        st.session_state.indexed = True

# Input section (always at the top)
input_container = st.container() 
with input_container:
    input_col, mic_col = st.columns([0.9, 0.1])
    question = None

    with input_col:
        question = st.chat_input("Ask me anything...")

    with mic_col:
        if not IS_CLOUD and st.button("🎤", help="Click to speak"):
            voice_input = get_voice_input()
            if voice_input:
                st.session_state.messages.append({"role": "user", "content": voice_input})
                question = voice_input

# Process question
if question:
    # Prevent duplicate user entry
    if not (st.session_state.messages and st.session_state.messages[-1]["role"] == "user" and st.session_state.messages[-1]["content"] == question):
        st.session_state.messages.append({"role": "user", "content": question})

    if not re.match(r'^[\s\S]{3,}$', question):
        response = "I couldn't understand that. Please ask a clear question."
        full_doc = None
    else:
        with st.spinner("🔍 Fetching answer..."):
            try:
                response, full_doc = get_similar_answer_from_documents(question, score_threshold=1.0)
                if not response or response.strip() == "":
                    response = "I'm not sure how to help with that. Please ask something related to Oracle documents."
                    full_doc = None
            except Exception as e:
                response = "I'm not sure how to help with that. Please ask something related to Oracle documents."
                full_doc = None

    st.session_state.messages.append({"role": "assistant", "content": response, "full_doc": full_doc})
    speak_text(response)

# Display chat history (always below input)
chat_container = st.container()
with chat_container:
    reversed_messages = list(reversed(st.session_state.messages))
    pairs = []
    temp_pair = []
    for msg in reversed_messages:
        temp_pair.append(msg)
        if msg["role"] == "user":
            pairs.append(temp_pair)
            temp_pair = []

    doc_counter = 0
    for pair in pairs:
        for msg in reversed(pair):
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if msg["role"] == "assistant" and "full_doc" in msg and msg["full_doc"]:
                    with st.expander("📄 View Full Document"):
                        st.text_area("Document Content", msg["full_doc"], height=400, key=f"doc_text_{doc_counter}")
                        st.download_button(
                            label="📂 Download .txt",
                            data=msg["full_doc"],
                            file_name="matched_document.txt",
                            mime="text/plain",
                            key=f"download_{doc_counter}"
                        )
                        doc_counter += 1
