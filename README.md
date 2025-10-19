# HTML to PowerPoint Converter

This project provides a powerful Python tool that leverages the Gemini 2.5 Pro multimodal AI to convert complex HTML files into editable PowerPoint (`.pptx`) presentations. It captures a visual screenshot of the HTML, sends it to the AI along with the source code and a detailed prompt, and orchestrates a self-healing loop to generate and debug a Python script that builds the presentation.

## Features

-   **🤖 AI-Powered Conversion:** Uses Gemini 2.5 Pro to understand the HTML structure and visual layout.
-   **📸 Visual Ground Truth:** Captures a 1920x1080 screenshot of the HTML using a headless browser (Playwright) and sends it to the AI, ensuring high-fidelity visual replication.
-   **🔄 Self-Healing Loop:** If the AI-generated code fails, the script automatically sends the code and the error message back to the AI to fix it, retrying up to 5 times.
-   **📦 Library & CLI:** Can be run as a standalone command-line tool or imported as a function (`convert_html_to_pptx`) into other Python projects.
-   **📝 Extensive Logging:** All conversion attempts, errors, and fixes are logged to a file (`conversion_log.txt`) for easy debugging.
-   **⚙️ Debugging Tools:** Automatically saves generated builder scripts and screenshots for each attempt, allowing you to trace the AI's logic.

## How It Works

The process is an orchestration of several steps designed for robustness:

1.  **Screenshot Capture:** A headless browser renders the input HTML and captures a pixel-perfect PNG image.
2.  **Multimodal Prompting:** The screenshot, the full HTML source, and a highly-detailed instructional prompt are sent to the Gemini 2.5 Pro API.
3.  **Code Generation:** Gemini generates a standalone Python script (`builder.py`) that uses the `python-pptx` library to construct the presentation.
4.  **Execution & Validation:** The main script executes this `builder.py`.
5.  **Self-Healing:** If the builder script fails:
    *   The error is captured.
    *   The faulty code and the error message are sent back to Gemini.
    *   Gemini provides a corrected script.
    *   The process repeats up to 5 times.
6.  **Output:** On success, the final `.pptx` file is saved locally or uploaded to a specified S3 bucket.

## Prerequisites

-   Python 3.9+
-   An active Google Gemini API Key.

## Quick Start

1.  **Clone the repository:**
    ```bash
    git clone <your-repo-url>
    cd PDFToPPTCode
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    playwright install chromium
    ```

3.  **Set your Gemini API Key:**
    ```bash
    export GEMINI_API_KEY="your_google_gemini_api_key_here"
    ```

4.  **Run the CLI:**
    To convert an HTML file, run the main script from your terminal.
    ```bash
    python htmlToPPT.py --html sampleHTML/slide_1.html --save-screenshot
    ```
    This will create `slide_1.pptx` and `slide_1_screenshot.png` inside the `sampleHTML` directory.

## Using as a Library

You can easily integrate the conversion logic into your own projects. See `example_usage.py` for a complete working example.

```python
# your_script.py
import os
from htmlToPPT import convert_html_to_pptx, ConversionError

# Ensure you have an absolute path
html_file = os.path.abspath("sampleHTML/slide_2.html")

try:
    print(f"Starting conversion for {html_file}...")

    # Call the function with desired options
    pptx_path = convert_html_to_pptx(
        html_path=html_file,
        save_builder_scripts=True,  # Recommended for debugging
        save_screenshot=True
    )
    
    print(f"\n✅ Success! PowerPoint saved at: {pptx_path}")

except FileNotFoundError:
    print(f"❌ Error: The file {html_file} was not found.")
except ConversionError as e:
    print(f"❌ Conversion failed after all attempts: {e}")
```

## Project Structure

```
.
├── htmlToPPT.py                # The main script with the core logic and CLI.
├── example_usage.py            # Example script showing how to use as a library.
├── requirements.txt            # Python dependencies.
├── sampleHTML/                 # Directory containing example HTML files.
├── README.md                   # This file.
├── QUICK_START.md              # Quick start guide.
├── PROMPT_UPDATES.md           # Documentation on prompt changes.
└── SCREENSHOT_FEATURE.md       # Documentation on the screenshot feature.
```

## The Prompting Strategy

The core of this project is a highly-engineered prompt that guides the AI. It includes:
-   Strict rules about what `python-pptx` methods *not* to use.
-   Explicit instructions on API usage patterns (e.g., no command chaining).
-   Emphasis on using the screenshot as the "visual ground truth".

This prompt is continuously refined based on the errors logged during failed conversions, making the system progressively more intelligent and reliable.
