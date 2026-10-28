# LearnMate AI — redesigned audio tutor

## What this version does
- **One complete audio:** creates one continuous downloadable MP3 for the whole uploaded file.
- **Slide-by-slide:** automatically processes every slide/page in order and creates a separate audio for each; no manual slide selection.
- **Language-focused teaching:** selected language is explicitly required throughout; Tamil mode uses Tamil sentences and preserves technical terms in English. Tanglish is a separate mode.
- **Doubt chat:** ask questions about the uploaded material; the answer is generated in the selected language and read aloud.
- **Duration target:** choose 5–60 minutes. The app allocates a target word count across the source sections. Real duration is approximate, not guaranteed.

## Start on Windows
1. Extract the ZIP.
2. Double-click `run_learnmate.bat`.
3. Wait for installation and open the local URL shown in the terminal (usually http://localhost:8501).
4. In the left sidebar, select **Gemini (cloud)** and paste your Google AI Studio API key. Default model is `gemini-3.5-flash-lite`.
5. Upload your PPTX/PDF/DOCX/TXT, choose language and mode, then click **Generate lesson and audio**.

## Free use / privacy
- Gemini needs internet and an API key. Free quotas and model availability can change; don't enable billing if you want to avoid charges.
- The key is entered in the UI and isn't deliberately saved by this app.
- Extracted material/questions go to Gemini when cloud mode is selected. Narration text is sent to the online Edge TTS service to create audio.
- Local Ollama mode is available, but may be slower/less capable on an 8 GB laptop.

## Important limitations
- Requested duration is approximate, not a guarantee. Detailed explanations of a large deck can take several minutes.
- Text extraction may miss diagrams, screenshots, scanned PDFs, or images containing text.
- The model can still make mistakes or fail to follow the selected language perfectly. Check important exam facts.
