"""
Orchestrator: Gemini 2.5 Pro -> (codegen) -> run generated script -> PPTX -> S3 link

Prereqs:
  pip install google-generativeai python-pptx beautifulsoup4 lxml boto3 requests playwright
  playwright install chromium

Env:
  export GEMINI_API_KEY=your_key
  export AWS_ACCESS_KEY_ID=...
  export AWS_SECRET_ACCESS_KEY=...
  export AWS_DEFAULT_REGION=ap-south-1   # or your region

Usage:
  python main.py --html slide.html --s3-bucket my-bucket --key-prefix kelp/pptx/
  # -> prints a HTTPS download URL when done
"""

import argparse, os, re, sys, tempfile, subprocess, uuid, json, pathlib, textwrap
import logging
import requests
import boto3
from botocore.config import Config
from playwright.sync_api import sync_playwright
import base64
from PIL import Image

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class ConversionError(Exception):
    """Custom exception for failures in the HTML to PPT conversion process."""
    pass

# ---------- Cost Calculation Constants ----------
# NOTE: Using pricing for Gemini 2.0 Pro as of late 2025. 
# Update these values if the pricing for Gemini 2.5 Pro changes.
# Pricing is based on characters, not tokens, for this model.
COST_PER_1K_CHARS_INPUT = 0.00125
COST_PER_1K_CHARS_OUTPUT = 0.010

# ---------- 1) Capture HTML Screenshot ----------
def capture_html_screenshot(html_path: str, output_path: str, width: int = 1920, height: int = 1080) -> str:
    """
    Capture a screenshot of the HTML file using Playwright.
    Returns the path to the saved screenshot.
    """
    logging.info(f"Capturing screenshot of HTML file: {html_path}")
    logging.info(f"Viewport size: {width}x{height}")
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={'width': width, 'height': height})
            
            # Load the HTML file
            html_file_url = f"file://{html_path}"
            logging.info(f"Loading: {html_file_url}")
            page.goto(html_file_url, wait_until="networkidle")
            
            # Wait a bit for any animations or dynamic content
            page.wait_for_timeout(1000)
            
            # Capture screenshot
            page.screenshot(path=output_path, full_page=False)
            browser.close()
            
        logging.info(f"Screenshot saved to: {output_path}")
        
        # Verify screenshot was created
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            logging.info(f"Screenshot size: {size / 1024:.2f} KB")
            return output_path
        else:
            raise RuntimeError("Screenshot file was not created")
            
    except Exception as e:
        logging.error(f"Failed to capture screenshot: {e}")
        raise

# ---------- 2) Gemini client ----------
def get_gemini_client():
    import google.generativeai as genai
    logging.info("Initializing Gemini client...")
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.error("GEMINI_API_KEY environment variable is not set")
        raise RuntimeError("GEMINI_API_KEY is not set")
    logging.info(f"API key found: {api_key[:20]}...")
    genai.configure(api_key=api_key)
    # Model: use the latest 2.5 Pro name you have access to
    model = genai.GenerativeModel("gemini-2.5-pro")
    logging.info("Gemini client initialized successfully")
    return model

def request_code_fix(original_code: str, error_message: str, model) -> str:
    """Sends the faulty code and error message to Gemini and asks for a fix."""
    logging.info("="*30 + " Requesting Code Fix " + "="*30)
    fix_prompt = textwrap.dedent(f"""
    The following Python script, which uses the `python-pptx` library, failed to execute.
    Please analyze the code and the error message to identify and fix the issue.

    **CRITICAL:** 
    - Provide only the complete, corrected, and runnable Python script.
    - Do not include any explanations, apologies, or markdown formatting outside of the code block.
    - Ensure the corrected script is a single, self-contained file.
    - Pay close attention to common `python-pptx` API errors like chained `.solid().fore_color` calls or using non-existent classes like `Px`.

    --- FAULTY CODE ---
    ```python
    {original_code}
    ```

    --- ERROR MESSAGE ---
    ```
    {error_message}
    ```

    --- CORRECTED PYTHON SCRIPT ---
    """)
    
    logging.info("Sending faulty code and error message to Gemini for correction...")
    try:
        response = model.generate_content(fix_prompt)
        
        # --- Cost Calculation ---
        input_char_count = len(fix_prompt)
        output_char_count = len(response.text or "")
        # --- End Cost Calculation ---

        fixed_code = extract_code_block(response.text or "")
        if not fixed_code:
            logging.warning("Gemini did not return a code block for the fix.")
            return original_code # Return original code if fix is empty
        logging.info("Received corrected code from Gemini.")
        return fixed_code, input_char_count, output_char_count
    except Exception as e:
        logging.error(f"An exception occurred while requesting a code fix: {e}")
        return original_code, 0, 0 # Return original on failure, with zero cost

# ---------- 2) Presign S3 upload ----------
def presign_s3_pair(bucket: str, key_prefix: str = "pptx/", filename: str = None, expires=3600):
    logging.info("Generating S3 presigned URLs...")
    if filename is None:
        filename = f"{uuid.uuid4()}.pptx"
    key = f"{key_prefix.rstrip('/')}/{filename}".lstrip("/")
    logging.info(f"S3 Key: {key}")
    s3 = boto3.client("s3", config=Config(signature_version="s3v4"))
    put_url = s3.generate_presigned_url(
        "put_object",
        Params={
            "Bucket": bucket,
            "Key": key,
            "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        },
        ExpiresIn=expires,
    )
    # Public-style GET URL (works if bucket or path is publicly readable or behind CloudFront).
    # In most setups, you'll return a short-lived CloudFront URL or a signed GET URL.
    get_url = f"https://{bucket}.s3.amazonaws.com/{key}"
    logging.info("S3 presigned URLs generated successfully")
    return put_url, get_url, key

# ---------- 3) Prompt for code generation ----------
def build_codegen_prompt(html_str: str) -> str:
    # Comprehensive prompt with all visual replication requirements
    return textwrap.dedent(f"""
    You are an expert presentation engineer specializing in pixel-perfect HTML-to-PowerPoint conversion.
    
    🎯 OBJECTIVE: Write a STANDALONE Python script (builder.py) that converts the given HTML into an EXACT VISUAL REPLICA PowerPoint (.pptx).
    
    ⚠️ CRITICAL: I have provided you with:
    1. A SCREENSHOT showing the exact visual output of the HTML file as it renders in a browser
    2. The complete HTML source code
    
    YOU MUST STUDY THE SCREENSHOT VERY CAREFULLY. The screenshot shows the EXACT layout, colors, fonts, positions, 
    spacing, and visual appearance that your PowerPoint output MUST match. Use the HTML code to understand the 
    structure and extract data, but use the SCREENSHOT as the ground truth for all visual aspects.
    
    ═══════════════════════════════════════════════════════════════════════════════
    📋 CORE DIRECTIVES & CONSTRAINT HIERARCHY (MOST IMPORTANT - FOLLOW THIS PRIORITY):
    ═══════════════════════════════════════════════════════════════════════════════
    
    When constraints conflict, adhere to this strict priority order:
    
    1. **DISPLAY ALL DATA:** Every single element, text node, image, and component from the HTML MUST be visibly rendered in the PowerPoint. Do NOT omit, truncate, hide, or reorder ANY data or visual elements.
    
    2. **EXACT VISUAL REPLICATION:** Before writing any code, CAREFULLY and DILIGENTLY analyze BOTH the screenshot and HTML:
       - LOOK AT THE SCREENSHOT FIRST: Study the actual rendered visual output
       - Note the exact positioning of every element in the screenshot
       - Observe the actual colors, fonts, sizes, and styling as they appear visually
       - Identify charts, tables, lists, and their exact data presentation
       - Measure spacing, padding, and gaps between elements visually
       - Note the visual hierarchy and balance of the composition
       - Then use the HTML to extract text content, data values, and structural information
       - Your PowerPoint MUST look identical to the screenshot - this is the PRIMARY reference
       
    3. **FIXED SLIDE DIMENSIONS:** The PowerPoint slide MUST be exactly 1920px × 1080px (16:9 aspect ratio).
       - All content MUST fit within these dimensions
       - NO content should overflow or be hidden
       - NO scrolling should be required
       
    4. **MAINTAIN READABILITY:** Text must remain readable:
       - Body text font size should NOT fall below 12px
       - If content overflows: shrink padding first, then gaps, then cautiously reduce font sizes
       - Preserve visual hierarchy and importance
    
    ═══════════════════════════════════════════════════════════════════════════════
    🔧 TECHNICAL REQUIREMENTS:
    ═══════════════════════════════════════════════════════════════════════════════
    
    **Libraries:**
    - Use ONLY: python-pptx, beautifulsoup4, lxml, Pillow, requests
    - The script MUST provide a CLI: python builder.py --html input.html --out output.pptx
    - Do NOT write base64 to stdout. Do NOT print PPTX content. Only create file on disk.
    
    **CRITICAL - Methods/Classes That DO NOT EXIST in python-pptx:**
    - Px() - DOES NOT EXIST! Use Pt(), Emu, or direct multiplication by 9525.
    - shapes.add_freeform_builder() - DOES NOT EXIST!
    - MSO_SHAPE.LINE - DOES NOT EXIST! Use MSO_CONNECTOR.STRAIGHT for lines.
    - **CRITICAL API USAGE - DO NOT CHAIN FILL COMMANDS:**
        - The `.solid()` method on a fill returns `None`. Chaining it will **ALWAYS** cause an `AttributeError`.
        - **NEVER WRITE THIS:** `shape.fill.solid().fore_color.rgb = ...`
        - **ALWAYS USE TWO STEPS:**
            - `shape.fill.solid()`
            - `shape.fill.fore_color.rgb = ...`
    - Available imports from pptx.util: Inches, Pt, Cm, Mm, Emu (NO Px!).
    - For pixel values, use direct EMU conversion: `px_value * 9525`.
    - For line widths: use `Pt()` or direct EMU values.
    - For lines/connectors: Use `shapes.add_connector(MSO_CONNECTOR.STRAIGHT, ...)`.
    - DO NOT WRITE: `Px()`, `add_freeform_builder()`, `MSO_SHAPE.LINE` - these will cause errors!
    
    **API USAGE - BULLET POINTS:**
    - Bullet properties belong to the `paragraph` object, NOT the `font` object.
    - Incorrect: `run.font.bullet.char = "•"` -> This will raise an AttributeError.
    - Correct: `paragraph.font.bold = True` and then `paragraph.level = 0` to apply default bullet style.

    **API USAGE - PARAGRAPH FORMAT:**
    - Paragraph formatting is accessed directly on the paragraph object, not a sub-property.
    - Incorrect: `p.paragraph_format.alignment = ...` -> This will raise an AttributeError.
    - Correct: `p.alignment = ...`
    
    **JSON EXTRACTION:**
    - If data is inside a <script> tag (e.g., `const INPUT_JSON = {{...}};`), you MUST use a regular expression to reliably extract the JSON object.
    - DO NOT use simple string splitting like `.split('=')[1]`, as it is brittle and will fail.
    - Correct regex pattern: `re.search(r'const INPUT-JSON = ({{.*?}});', script_content, re.DOTALL)`
    - This captures everything from the opening `{{` to the closing `}};`.
    
    **Coordinate System:**
    - Slide dimensions: 1920 × 1080 pixels
    - Convert to EMU: 1 px = 9525 EMU (exactly)
    - Slide width = 1920 × 9525 = 18288000 EMU
    - Slide height = 1080 × 9525 = 10287000 EMU
    - To convert pixels to EMU in code: emu_value = px_value * 9525
    - Example: left=100px → left=952500 EMU, width=200px → width=1905000 EMU
    
    **Positioning & Layout:**
    - Parse ALL CSS positioning (absolute, relative, flex, grid)
    - For flexbox/grid layouts: compute final absolute positions for each element
    - Map CSS box model (margin, border, padding, content) to PowerPoint shapes
    - Respect z-index layering (render elements in correct order)
    - Handle CSS transforms (translate, scale) if present
    
    **Typography:**
    - Default font family: "Segoe UI", fallback to "Calibri" or "Arial"
    - Font sizes must match HTML exactly (convert px to pt: pt = px * 0.75)
    - Preserve: bold, italic, underline, strikethrough
    - Preserve: text-align (left, center, right, justify)
    - Preserve: line-height, letter-spacing if specified
    - Preserve: color (convert hex/rgb to RGB tuple)
    - Handle inline formatting: <b>, <i>, <strong>, <em>, <span style="color:...">
    
    **Colors & Backgrounds:**
    - Parse and apply all background colors, gradients if possible
    - Parse and apply text colors
    - Parse and apply border colors, widths, and styles
    - Support border-radius for rounded corners (use rounded shapes)
    
    **SHAPE STYLE:**
    - All container boxes and shapes MUST have sharp, 90-degree corners.
    - DO NOT use rounded rectangles or shapes with any `border-radius`. All corners must be edgy and rectangular.
    
    **Images:**
    - If <img> has remote src, download it using requests
    - If <img> has data: URI, decode and save
    - Place images at exact positions with exact dimensions
    - Support object-fit: cover/contain/fill
    - Maintain aspect ratios when appropriate
    
    **Charts & Visualizations:**
    - If HTML contains <canvas> elements or Chart.js references, analyze the data
    - Recreate charts using PowerPoint shapes, text boxes, and smart positioning
    - For bar charts: use rectangles with exact dimensions
    - For line charts: use line shapes connecting data points
    - For pie charts: use pie chart shapes or approximations
    - Display ALL data labels, values, and legends exactly as in HTML
    - Match colors, fonts, and styling of chart elements
    
    **Tables:**
    - Use PowerPoint table objects for <table> elements
    - Match cell borders, backgrounds, padding
    - Match text alignment and formatting within cells
    
    **Lists:**
    - Use PowerPoint text boxes with bullet formatting for <ul>/<ol>
    - Match bullet styles, indentation, and spacing
    
    **Special Components:**
    - For cards/containers: use grouped shapes with borders and backgrounds
    - For badges/pills: use rounded rectangle shapes with text
    - For icons (Font Awesome, etc.): use Unicode characters or small images
    
    ═══════════════════════════════════════════════════════════════════════════════
    🎨 VISUAL & BRANDING GUIDELINES:
    ═══════════════════════════════════════════════════════════════════════════════
    
    **Color Scheme:**
    - Primary color: #0078D4 (use for headings, primary text)
    - Respect all colors from HTML exactly
    
    **Typography Hierarchy:**
    - h1: 36px (27pt)
    - h2: 32px (24pt)
    - h3: 28px (21pt)
    - h4: 24px (18pt)
    - p, span, labels: 16px (12pt)
    - Minimum body text: 12px (9pt)
    
    **Spacing System:**
    - Use consistent spacing based on CSS variables if present
    - Default gap: 12px between flex items
    - Default padding: 16px inside containers
    
    ═══════════════════════════════════════════════════════════════════════════════
    🏢 KPMG FOOTER REQUIREMENT (IF PRESENT IN HTML):
    ═══════════════════════════════════════════════════════════════════════════════
    
    If the HTML contains a footer with KPMG branding:
    
    **Position:** Full-width bar at the very bottom (y = ~1020-1080px range)
    
    **Content Structure (left to right):**
    1. **KPMG Logo (left side):**
       - Image URL: https://nextgeneration.vc/wp-content/uploads/2018/07/kpmg-logo.png
       - Height: 65px
       - Download and embed in PowerPoint
    
    2. **Copyright Disclaimer (immediately to the right of logo, same line):**
       - Text: "© 2025 KPMG India Services LLP, an Indian limited liability company and a member firm of the KPMG network of independent member firms affiliated with KPMG International Cooperative ("KPMG International"), a Swiss entity. All rights reserved."
       - Font size: ~10-11px
       - Keep on one line if possible, wrap intelligently if needed
    
    3. **Classification (centered below, or right side):**
       - Text: "Document Classification: KPMG Confidential"
       - **Bold** formatting
       - Centered or right-aligned
    
    ═══════════════════════════════════════════════════════════════════════════════
    🧪 DYNAMIC SIZING & SPACE OPTIMIZATION:
    ═══════════════════════════════════════════════════════════════════════════════
    
    - Analyze content density: components with more data get more space
    - Components with little data (few key-values) should be sized smaller
    - Do NOT leave large empty areas
    - Distribute space proportionally based on content volume
    - If content threatens to overflow:
      1. Reduce padding/margins first
      2. Reduce gaps between elements
      3. Reduce font sizes (but not below minimums)
      4. Adjust component heights dynamically
    
    ═══════════════════════════════════════════════════════════════════════════════
    📦 DELIVERABLES:
    ═══════════════════════════════════════════════════════════════════════════════
    
    - Output ONLY a single, complete Python file between triple backticks (```python ... ```)
    - The code must be self-contained and executable after: pip install python-pptx beautifulsoup4 lxml Pillow requests
    - Must work with: python builder.py --html input.html --out output.pptx
    - Include comprehensive comments explaining your visual replication strategy
    - Handle errors gracefully (missing images, parsing issues)
    
    ═══════════════════════════════════════════════════════════════════════════════
    ⚡ ULTIMATE RULE:
    ═══════════════════════════════════════════════════════════════════════════════
    
    LOOK AT THE PROVIDED SCREENSHOT VERY VERY CAREFULLY before generating any code. 
    This screenshot shows you EXACTLY what the output should look like.
    
    Your PowerPoint slide MUST be a pixel-perfect replica of this screenshot.
    - Every element position must match what you see in the screenshot
    - Every color must match what you see in the screenshot  
    - Every font size and style must match what you see in the screenshot
    - Every spacing and gap must match what you see in the screenshot
    - Every piece of text and data must match what you see in the screenshot
    
    Use the HTML code to extract data and understand structure, but the SCREENSHOT is your visual ground truth.
    This is ABSOLUTELY CRITICAL. Your output will be compared side-by-side with this screenshot.
    
    ═══════════════════════════════════════════════════════════════════════════════

    Here is the HTML to convert (verbatim between <HTML> tags):
    <HTML>
    {html_str}
    </HTML>
    """)

# ---------- 4) Extract code from model output ----------
def extract_code_block(text: str) -> str:
    logging.info("Extracting code block from model response...")
    # Look for the first fenced code block ```python ... ```
    m = re.search(r"```(?:python)?\s*(.*?)```", text, flags=re.DOTALL | re.IGNORECASE)
    if m:
        logging.info("Code block extracted successfully")
        return m.group(1).strip()
    # Fallback: return whole text (not ideal)
    logging.warning("No code block found, using entire response")
    return text.strip()

# ---------- 5) Run generated script in a subprocess ----------
def run_generated_builder(builder_path: str, html_path: str, out_path: str, timeout_sec=300):
    cmd = [sys.executable, builder_path, "--html", html_path, "--out", out_path]
    logging.info(f"Running generated builder: {' '.join(cmd)}")
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout_sec, check=False)
        return proc
    except subprocess.TimeoutExpired as e:
        logging.error(f"Builder script timed out after {timeout_sec} seconds")
        # Create a mock process object for timeout to be handled in the main loop
        mock_proc = subprocess.CompletedProcess(cmd, timeout=True, returncode=1, stdout='', stderr=str(e))
        return mock_proc

# ---------- 6) Upload to S3 PUT URL ----------
def upload_via_presigned_put(put_url: str, file_path: str):
    logging.info(f"Uploading file to S3: {file_path}")
    file_size = os.path.getsize(file_path)
    logging.info(f"File size: {file_size / 1024:.2f} KB")
    with open(file_path, "rb") as f:
        r = requests.put(put_url, data=f, headers={
            "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        })
    if r.status_code not in (200, 201):
        logging.error(f"Upload failed with status {r.status_code}: {r.text}")
        raise RuntimeError(f"Upload failed: {r.status_code} {r.text}")
    logging.info("File uploaded successfully to S3")

# ---------- 7) Main Conversion Logic (as a callable function) ----------
def convert_html_to_pptx(
    html_path: str,
    output_pptx_path: str = None,
    s3_bucket: str = None,
    s3_key_prefix: str = "pptx/",
    save_builder_scripts: bool = False,
    save_screenshot: bool = False,
    log_file: str = "conversion_log.txt",
    max_retries: int = 5
):
    """
    Converts an HTML file to a PowerPoint presentation with a self-healing loop.

    This is the primary function to be used when integrating this script as a library.

    Args:
        html_path (str): The absolute path to the input HTML file.
        output_pptx_path (str, optional): The desired local path for the output PPTX. 
            If not provided, it will be saved next to the HTML file. Defaults to None.
        s3_bucket (str, optional): The S3 bucket to upload the final PPTX to. 
            If provided, the function returns an S3 URL instead of a local path. Defaults to None.
        s3_key_prefix (str, optional): The prefix (folder) to use in the S3 bucket. Defaults to "pptx/".
        save_builder_scripts (bool, optional): If True, saves the generated Python builder script for each attempt. Defaults to False.
        save_screenshot (bool, optional): If True, saves the captured screenshot for debugging. Defaults to False.
        log_file (str, optional): Path to the log file to append all logs. Defaults to "conversion_log.txt".
        max_retries (int, optional): The maximum number of times to try fixing a failed script. Defaults to 5.

    Returns:
        str: The local path to the generated PPTX file or the S3 download URL if `s3_bucket` is provided.

    Raises:
        ConversionError: If the conversion fails after all retry attempts.
        FileNotFoundError: If the input `html_path` does not exist.
    """
    # Setup file-based logging for this specific conversion run
    file_handler = logging.FileHandler(log_file, mode='a')
    # Add a unique header for this run to the log file
    logging.getLogger().addHandler(file_handler)
    
    logging.info("="*60)
    logging.info(f"Starting New Conversion for: {html_path}")
    logging.info("="*60)

    html_path_obj = pathlib.Path(html_path).resolve()
    if not html_path_obj.exists():
        raise FileNotFoundError(f"Input HTML file not found: {html_path}")

    logging.info(f"Reading HTML file: {html_path_obj}")
    html_str = html_path_obj.read_text(encoding="utf-8", errors="ignore")
    logging.info(f"HTML file size: {len(html_str)} characters")

    model = get_gemini_client()
    code = ""
    
    # --- Cost Tracking Initialization ---
    total_input_chars = 0
    total_output_chars = 0
    total_api_calls = 0
    # --- End Cost Tracking Initialization ---

    # 1) Capture screenshot and perform initial code generation
    with tempfile.TemporaryDirectory() as screenshot_dir:
        screenshot_path = os.path.join(screenshot_dir, "screenshot.png")
        capture_html_screenshot(str(html_path_obj), screenshot_path, width=1920, height=1080)
        
        if save_screenshot:
            screenshot_save_path = html_path_obj.with_name(f"{html_path_obj.stem}_screenshot.png")
            logging.info(f"Saving screenshot for reference: {screenshot_save_path}")
            import shutil
            shutil.copy(screenshot_path, screenshot_save_path)
            print(f"📸 Saved screenshot: {screenshot_save_path}")
        
        logging.info("Loading screenshot image for Gemini...")
        screenshot_image = Image.open(screenshot_path)
        
        prompt = build_codegen_prompt(html_str)
        
        logging.info("Requesting initial code generation from Gemini (with screenshot)...")
        resp = model.generate_content([screenshot_image, prompt])
        total_api_calls += 1
        
        # --- Cost Calculation ---
        # For multimodal input, character count is calculated on the text part.
        # The cost of the image is separate and typically priced per-image, but we'll note it.
        # For simplicity in this calculation, we focus on character count as per the model's pricing structure.
        input_chars = len(prompt)
        output_chars = len(resp.text or "")
        total_input_chars += input_chars
        total_output_chars += output_chars
        logging.info(f"API Call {total_api_calls}: Input Chars = {input_chars}, Output Chars = {output_chars}")
        # --- End Cost Calculation ---

        text = resp.text or ""
        logging.info(f"Received initial response from Gemini ({len(text)} characters)")
        code = extract_code_block(text)
        
        if not code or "def main(" not in code and "__name__" not in code:
            logging.warning("Initial response from Gemini did not appear to be a valid script.")

    # 2) Self-healing loop
    for attempt in range(1, max_retries + 1):
        logging.info("="*30 + f" Attempt {attempt}/{max_retries} " + "="*30)

        builder_save_path = None
        if save_builder_scripts:
            builder_save_path = html_path_obj.with_name(f"{html_path_obj.stem}_builder_attempt_{attempt}.py")
            logging.info(f"Saving builder script for attempt {attempt}: {builder_save_path}")
            with open(builder_save_path, "w", encoding="utf-8") as f:
                f.write(code)
            print(f"🧩 Saved builder script for attempt {attempt}: {builder_save_path}")

        with tempfile.TemporaryDirectory() as td:
            builder_path = os.path.join(td, "builder.py")
            with open(builder_path, "w", encoding="utf-8") as f:
                f.write(code)

            out_pptx = os.path.join(td, "out.pptx")
            proc = run_generated_builder(builder_path, str(html_path_obj), out_pptx)

            if proc.returncode == 0 and os.path.exists(out_pptx) and os.path.getsize(out_pptx) > 0:
                logging.info(f"SUCCESS: Builder script executed successfully on attempt {attempt}.")
                
                final_output_path = ""
                if s3_bucket:
                    _, get_url, _ = presign_s3_pair(s3_bucket, s3_key_prefix)
                    upload_via_presigned_put(get_url, out_pptx)
                    final_output_path = get_url
                    print(f"\n✅ PPTX uploaded successfully: {final_output_path}\n")
                else:
                    local_out = output_pptx_path or html_path_obj.with_suffix(".pptx")
                    pathlib.Path(out_pptx).replace(local_out)
                    final_output_path = str(local_out)
                    print(f"\n✅ PPTX saved locally: {final_output_path}\n")
                
                logging.info("="*30 + " Process Completed Successfully " + "="*30)
                # Clean up the handler for this run to avoid duplicate logging
                logging.getLogger().removeHandler(file_handler)
                
                # --- Final Cost Calculation on Success ---
                input_cost = (total_input_chars / 1000) * COST_PER_1K_CHARS_INPUT
                output_cost = (total_output_chars / 1000) * COST_PER_1K_CHARS_OUTPUT
                total_cost = input_cost + output_cost
                
                cost_summary = (
                    f"\n--- COST SUMMARY (SUCCESS) ---\n"
                    f"Total API Calls: {total_api_calls}\n"
                    f"Total Input Characters: {total_input_chars:,}\n"
                    f"Total Output Characters: {total_output_chars:,}\n"
                    f"Input Cost: ${input_cost:.6f}\n"
                    f"Output Cost: ${output_cost:.6f}\n"
                    f"Total Estimated Cost: ${total_cost:.6f}\n"
                    f"Note: Image input cost is not included in this calculation.\n"
                    f"-----------------------------"
                )
                print(cost_summary)
                logging.info(cost_summary)
                # --- End Final Cost Calculation ---
                
                return final_output_path

            # --- Handle Failure ---
            logging.error(f"FAIL: Builder script failed on attempt {attempt}.")
            error_output = f"STDOUT:\n{proc.stdout}\n\nSTDERR:\n{proc.stderr}"
            logging.error(f"Execution failed with return code {proc.returncode}.\n{error_output}")
            
            if attempt < max_retries:
                # Request a fix from Gemini and update costs
                fixed_code, input_chars, output_chars = request_code_fix(code, error_output, model)
                code = fixed_code
                total_api_calls += 1
                total_input_chars += input_chars
                total_output_chars += output_chars
                logging.info(f"API Call {total_api_calls}: Input Chars = {input_chars}, Output Chars = {output_chars} (for fix)")
            else:
                logging.error(f"Maximum number of retries ({max_retries}) reached. Giving up.")
                final_error_message = f"Failed to generate PPTX for {html_path} after {max_retries} attempts."
                if builder_save_path:
                    final_error_message += f" Check the final faulty builder script at: {builder_save_path}"
                
                # --- Final Cost Calculation on Failure ---
                input_cost = (total_input_chars / 1000) * COST_PER_1K_CHARS_INPUT
                output_cost = (total_output_chars / 1000) * COST_PER_1K_CHARS_OUTPUT
                total_cost = input_cost + output_cost
                
                cost_summary = (
                    f"\n--- COST SUMMARY (FAILED) ---\n"
                    f"Total API Calls: {total_api_calls}\n"
                    f"Total Input Characters: {total_input_chars:,}\n"
                    f"Total Output Characters: {total_output_chars:,}\n"
                    f"Input Cost: ${input_cost:.6f}\n"
                    f"Output Cost: ${output_cost:.6f}\n"
                    f"Total Estimated Cost: ${total_cost:.6f}\n"
                    f"Note: Image input cost is not included in this calculation.\n"
                    f"-----------------------------"
                )
                print(cost_summary)
                logging.info(cost_summary)
                # --- End Final Cost Calculation ---
                
                # Clean up handler before raising
                logging.getLogger().removeHandler(file_handler)
                raise ConversionError(final_error_message)

    # Fallback in case loop finishes unexpectedly
    logging.getLogger().removeHandler(file_handler)
    raise ConversionError("Process failed unexpectedly after retry loop.")


# ---------- 8) Main CLI Entrypoint ----------
def main():
    """Parses command line arguments and calls the main conversion function."""
    ap = argparse.ArgumentParser(
        description="Converts an HTML file to a PowerPoint presentation using a self-healing AI loop.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    ap.add_argument("--html", required=True, help="Path to the input HTML file.")
    ap.add_argument("--output", help="Optional: Path to save the output PPTX file.")
    ap.add_argument("--s3-bucket", help="Optional: S3 bucket to upload the result to.")
    ap.add_argument("--key-prefix", default="pptx/", help="S3 key prefix (folder).")
    ap.add_argument("--save-builder-scripts", action="store_true", help="Save the generated Python builder script for each attempt.")
    ap.add_argument("--save-screenshot", action="store_true", help="Save the captured screenshot for debugging.")
    ap.add_argument("--log-file", default="conversion_log.txt", help="Path to the log file for all attempts.")
    args = ap.parse_args()

    # Setup a stream handler for console output for the CLI mode
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().addHandler(stream_handler)
    
    try:
        result_path = convert_html_to_pptx(
            html_path=args.html,
            output_pptx_path=args.output,
            s3_bucket=args.s3_bucket,
            s3_key_prefix=args.key_prefix,
            save_builder_scripts=args.save_builder_scripts,
            save_screenshot=args.save_screenshot,
            log_file=args.log_file
        )
        print(f"Conversion successful. Result available at: {result_path}")
    except (ConversionError, FileNotFoundError) as e:
        logging.error(f"CONVERSION FAILED: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred during the conversion process: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
