# Quick Start Guide - HTML to PPT Converter

## Prerequisites

### 1. Install Dependencies
```bash
pip install google-generativeai python-pptx beautifulsoup4 lxml boto3 requests Pillow playwright
playwright install chromium
```

**Or use requirements.txt:**
```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Set Gemini API Key
```bash
export GEMINI_API_KEY="your_gemini_api_key_here"
```

Get your API key from: https://makersuite.google.com/app/apikey

## Basic Usage

### Convert HTML to Local PPTX
```bash
python htmlToPPT.py --html path/to/your/file.html
```
**Output:** Creates `file.pptx` in the same directory as the HTML file.

### Save the Generated Builder Script
```bash
python htmlToPPT.py --html path/to/your/file.html --save-builder
```
**Output:** Creates both `file.pptx` and `file_builder.py`

### Save Screenshot for Reference (Recommended)
```bash
python htmlToPPT.py --html path/to/your/file.html --save-builder --save-screenshot
```
**Output:** Creates `file.pptx`, `file_builder.py`, and `file_screenshot.png`
**Note:** The screenshot is automatically captured and sent to Gemini for better accuracy!

### Upload to S3 (Optional)
```bash
# Set AWS credentials first
export AWS_ACCESS_KEY_ID="your_aws_key"
export AWS_SECRET_ACCESS_KEY="your_aws_secret"
export AWS_DEFAULT_REGION="ap-south-1"

# Convert and upload
python htmlToPPT.py --html file.html --s3-bucket my-bucket --key-prefix kelp/pptx/
```

## Features

### ✅ What the Script Does

1. **Captures screenshot** of HTML using headless browser (Playwright)
2. **Reads HTML file** with all styles and content
3. **Sends to Gemini AI** with screenshot + HTML code + comprehensive rules
4. **Generates Python builder script** that creates the PowerPoint
5. **Runs the builder** to create the PPTX file
6. **Saves locally** or uploads to S3

### 🎯 What Gemini Will Do (Based on Updated Prompt)

- **🖼️ Study the screenshot first** - Visual ground truth for exact replication
- **Analyze HTML carefully** before generating code
- **Create exact visual replica** matching the screenshot pixel-by-pixel
- **Preserve all data** - no omissions or truncations
- **Match all colors, fonts, and sizes** as seen in the screenshot
- **Recreate charts and visualizations** with data labels
- **Handle KPMG branding** (logo, copyright, classification)
- **Fit everything** in 1920×1080 slide with no overflow
- **Maintain readability** (minimum 12px text)

## What Gets Created

### Example File Structure After Running:
```
your-project/
├── sampleHTML/
│   ├── slide_1.html               # Your input HTML
│   ├── slide_1.pptx               # ✨ Generated PowerPoint
│   ├── slide_1_builder.py         # 🔧 Generated builder script (with --save-builder)
│   └── slide_1_screenshot.png     # 📸 Screenshot reference (with --save-screenshot)
├── htmlToPPT.py                    # Main converter script
├── requirements.txt                # Python dependencies
└── QUICK_START.md                  # This file
```

## Logging Output

The script now includes detailed logging:

```
2025-10-19 16:18:30 - INFO - ============================================================
2025-10-19 16:18:30 - INFO - HTML to PPT Converter - Starting
2025-10-19 16:18:30 - INFO - ============================================================
2025-10-19 16:18:30 - INFO - Input HTML file: sampleHTML/slide_1.html
2025-10-19 16:18:30 - INFO - Reading HTML file: /path/to/slide_1.html
2025-10-19 16:18:30 - INFO - HTML file size: 23088 characters
2025-10-19 16:18:31 - INFO - Initializing Gemini client...
2025-10-19 16:18:31 - INFO - Gemini client initialized successfully
2025-10-19 16:18:31 - INFO - Requesting code generation from Gemini...
2025-10-19 16:21:02 - INFO - Received response from Gemini (21194 characters)
2025-10-19 16:21:02 - INFO - Code block extracted successfully
2025-10-19 16:21:02 - INFO - Running generated builder...
2025-10-19 16:21:04 - INFO - PPTX file created successfully: 30.75 KB
2025-10-19 16:21:04 - INFO - Process completed successfully
```

## Troubleshooting

### Error: "GEMINI_API_KEY is not set"
**Solution:** Set the environment variable:
```bash
export GEMINI_API_KEY="your_key"
```

### Error: "ModuleNotFoundError"
**Solution:** Install missing dependencies:
```bash
pip install google-generativeai python-pptx beautifulsoup4 lxml boto3 requests Pillow
```

### Error: "API key not valid"
**Solution:** Get a valid Gemini API key from https://makersuite.google.com/app/apikey
(Keys start with "AIza...")

### Builder Script Fails
**Solution:** Check the saved builder script for errors:
```bash
python htmlToPPT.py --html file.html --save-builder
# Then manually inspect file_builder.py
```

## Advanced Options

### Custom S3 Bucket and Prefix
```bash
python htmlToPPT.py \
  --html myfile.html \
  --s3-bucket production-pptx \
  --key-prefix reports/2025/
```

### Process Multiple Files
```bash
for file in sampleHTML/*.html; do
  echo "Converting $file..."
  python htmlToPPT.py --html "$file"
done
```

## Performance Notes

- **Processing Time:** 2-3 minutes per HTML file (Gemini generation takes most time)
- **File Size:** Typical output is 30-50 KB for simple slides
- **Chart Handling:** Charts may take longer to process due to complexity
- **Image Downloads:** Remote images are downloaded and embedded

## Key Features from Updated Prompt

### 1. Exact Visual Replication
- Gemini carefully analyzes HTML structure
- Recreates exact layout, colors, fonts
- Preserves visual hierarchy

### 2. Data Preservation
- All content is displayed
- No truncation or omission
- Charts show all data points

### 3. KPMG Branding Support
- Logo placement (if present)
- Copyright disclaimer
- Classification label

### 4. Dynamic Sizing
- Content-based space allocation
- No large empty areas
- Intelligent overflow handling

### 5. Chart Support
- Bar charts, line charts, pie charts
- Data labels on charts
- Color and style matching

## Next Steps

1. **Test the converter:**
   ```bash
   python htmlToPPT.py --html sampleHTML/slide_1.html --save-builder
   ```

2. **Review the output:**
   - Open `slide_1.pptx` in PowerPoint/Keynote
   - Compare with original HTML in browser
   - Review `slide_1_builder.py` to see Gemini's approach

3. **Iterate if needed:**
   - Adjust HTML structure if output isn't perfect
   - Tweak CSS for better PowerPoint translation
   - Report issues for prompt improvements

## Support

For issues or improvements:
1. Check the logs for detailed error messages
2. Review the generated builder script
3. Adjust the prompt in `htmlToPPT.py` if needed
4. Test with simpler HTML first, then increase complexity

---

**Ready to convert?** Run:
```bash
export GEMINI_API_KEY="your_key"
python htmlToPPT.py --html your_file.html --save-builder
```

