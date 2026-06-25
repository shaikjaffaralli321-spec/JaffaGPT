# JaffaGPT

An attractive single-chatbot AI utility app. One chat window can handle Chat Q&A, translation, image-prompt generation, image scanning, voice input/output, and document-based answers.

## Features

- Real-time streaming chat with Server-Sent Events
- One chatbot composer with Chat Q&A, Translate, Image Generate, and Voice Assist tools
- Optional OpenAI Responses API integration through `openai_config.py`
- Local fallback answers when no OpenAI API key is configured
- Translation and image-prompt help through the same chat flow
- Image scan uploads that summarize visible information with OpenAI vision
- Image to PDF conversion from the upload document panel
- Generated SVG image preview for image requests
- Voice input and spoken replies using browser Web Speech APIs
- Document upload keeps original files locally and indexes readable text when possible
- Built-in text extraction for text/code files, Markdown, CSV/TSV, JSON, XML, HTML, RTF, email files, DOCX, PPTX, XLSX, and optional PDF
- Unknown file types are still saved with a knowledge-base note when text cannot be extracted
- Uploads stay out of the chat conversation; ask a prompt later to use the saved knowledge base
- Local knowledge base stored in `data/knowledge_base.json`
- Source-matched answers with practical next steps
- Answer modes for simple, action, official, and student-friendly responses
- Responsive polished chatbot UI

## OpenAI API Key

Edit `openai_config.py`:

```python
OPENAI_API_KEY = "your-api-key-here"
OPENAI_MODEL = "gpt-4o-mini"
```

You can also use environment variables instead:

```powershell
$env:OPENAI_API_KEY="your-api-key-here"
$env:OPENAI_MODEL="gpt-4o-mini"
```

## Google Gemini API Key

The app can also connect to Google Gemini. Edit `openai_config.py`:

```python
GEMINI_API_KEY = "your-google-gemini-api-key-here"
GEMINI_MODEL = "gemini-1.5-flash"
```

Or use environment variables:

```powershell
$env:GEMINI_API_KEY="your-google-gemini-api-key-here"
$env:GEMINI_MODEL="gemini-1.5-flash"
```

Restart the app. The server will use Gemini automatically when it is the first configured AI service available.

## Run

```powershell
python app.py
```

Open `http://127.0.0.1:8000`.

## Optional PDF Support

PDF text extraction uses PyMuPDF:

```powershell
pip install pymupdf
```

Without it, PDFs are still kept locally, but detailed PDF Q&A needs selectable text extraction.

## Optional Image Conversion Support

Image to PDF conversion uses Pillow:

```powershell
pip install pillow
```

Image scanning needs an OpenAI API key because it uses a vision model.

## Next Upgrades

- Add user accounts and admin review for trusted sources
- Store embeddings in PostgreSQL with pgvector
- Add OCR for scanned documents
- Add server-side speech-to-text and text-to-speech for accessibility
- Add audit logs and safety disclaimers by category
