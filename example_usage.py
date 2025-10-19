import os
import logging
from htmlToPPT import convert_html_to_pptx, ConversionError

# --- Configuration ---
# Configure basic logging to see the output from the converter
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Main Example ---
if __name__ == "__main__":
    # --- Example 1: Successful Conversion ---
    print("\n" + "="*50)
    print("🚀 Example 1: Running a successful conversion for slide_1.html")
    print("="*50)
    try:
        # Use an absolute path for reliability
        slide1_html_path = os.path.abspath("sampleHTML/slide_1.html")
        
        if not os.path.exists(slide1_html_path):
            print(f"❌ ERROR: Cannot find the HTML file at {slide1_html_path}")
            print("Please ensure the file exists before running the example.")
        else:
            # Call the main conversion function
            pptx_path = convert_html_to_pptx(
                html_path=slide1_html_path,
                save_builder_scripts=True,  # Save scripts for debugging
                save_screenshot=True       # Save the screenshot for comparison
            )
            print(f"\n✅ SUCCESS! PowerPoint for slide 1 created at: {pptx_path}")

    except FileNotFoundError as e:
        print(f"❌ FILE NOT FOUND ERROR: {e}")
    except ConversionError as e:
        print(f"❌ CONVERSION FAILED for slide 1: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    # --- Example 2: Handling a Known-Bad File (if it exists) ---
    print("\n" + "="*50)
    print("🔥 Example 2: Simulating a failure with a non-existent file")
    print("="*50)
    try:
        non_existent_html_path = os.path.abspath("sampleHTML/non_existent_slide.html")
        
        # This call is expected to fail with a FileNotFoundError
        convert_html_to_pptx(html_path=non_existent_html_path)

    except FileNotFoundError as e:
        print(f"✅ SUCCESS! Correctly caught expected error: {e}")
    except ConversionError as e:
        # This won't be hit in this specific case, but it's good practice
        print(f"❌ CONVERSION FAILED for non-existent file: {e}")

    print("\n" + "="*50)
    print("✨ Example script finished. ✨")
    print("="*50)
