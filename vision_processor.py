"""
Vision Processor
Extracts structured data from images using Llama Vision.
"""

import os
import json
import base64
from groq import Groq
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

OCR_DATA_DIR = "vision_data"

def encode_image(image_path):
    """Encodes an image file to a base64 string."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def process_image(image_path, output_dir=OCR_DATA_DIR):
    """
    Sends an image to the Vision Model to extract structured data.
    Saves the JSON response to ocr_data/ folder with _vision.json suffix.
    """
    # API Key from environment variable
    api_key = os.environ.get("GROQ_API_KEY")

    if not api_key:
        print("  [Vision] Error: GROQ_API_KEY not found in environment variables.")
        print("  Please set it in your terminal: export GROQ_API_KEY='your_key_here'")
        return None

    try:
        client = Groq(api_key=api_key)
        base64_image = encode_image(image_path)
        image_data_url = f"data:image/jpeg;base64,{base64_image}"

        # User's specific prompt for high accuracy
        prompt_text = """### INPUT ANALYSIS PHASE
1. **Identify Script Style**: First, detect if the text is printed, handwritten, or a mix. Adjust your internal recognition weights for handwritten characters.
2. **Spatial Anchoring**: Use visual lines and margins to define table boundaries.

### EXTRACTION REQUIREMENTS
1. **Hierarchical Extraction**:
   - **Header Data**: Extract titles, dates, and names.
   - **Tabular Structures**: Reconstruct grid-like data.
   - **Currency Detection**: Capture currency symbols exactly.
2. **Handwritten Edge Cases**:
   - Ignore crossed-out text.
   - Use [illegible] for unreadable words.
3. **The "Logical Audit"**:
   - Manually recalculate sums.
   - Report discrepancies as `handwritten_total` and `computed_total`.

### OUTPUT STRUCTURE (JSON)
Constraints:
- Return ONLY valid JSON.
- Do not include comments or markdown formatting (```json).
- "normalized_value" must be a single number or string, NOT a mathematical expression (e.g., use 100.50, not 50+50.50).
- If a calculation is needed, perform it internally and output the result.

{
  "document_summary": { "style": "handwritten/printed", "domain": "auto-detect" },
  "entities": { "label": "value" },
  "tables": [
    {
      "table_description": "string",
      "headers": [],
      "rows": [
        { "column_name": "raw_text", "normalized_value": numeric_or_string, "currency": "ISO_CODE" }
      ],
      "validation": {
        "math_check": "passed/failed",
        "notes": "string"
      }
    }
  ]
}
"""

        print(f"  [Vision] Improving OCR for {os.path.basename(image_path)}...")

        completion = client.chat.completions.create(
            model="meta-llama/llama-4-maverick-17b-128e-instruct", 
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt_text},
                        {"type": "image_url", "image_url": {"url": image_data_url}}
                    ]
                }
            ],
            temperature=0.1,
            response_format={"type": "json_object"}, 
            stream=False
        )
        
        content = completion.choices[0].message.content
        
        try:
            data = json.loads(content)
        except json.JSONDecodeError:
            print(f"  [Vision] Error: Response was not valid JSON.")
            return None

        # Save to file
        os.makedirs(output_dir, exist_ok=True)
        img_name = os.path.basename(image_path)
        stem = img_name.rsplit('.', 1)[0]
        json_path = os.path.join(output_dir, f"{stem}_vision.json")
        
        with open(json_path, "w") as f:
            json.dump(data, f, indent=2)
            
        print(f"    → Saved improved data to {json_path}")
        return data

    except Exception as e:
        print(f"  [Vision] Error: {e}")
        return None

def process_all_images(image_list, base_dir="input"):
    """Run Vision extraction on all images."""
    results = []
    output_dir = OCR_DATA_DIR
    os.makedirs(output_dir, exist_ok=True)
    
    for img_name in image_list:
        full_path = os.path.join(base_dir, img_name)
        if os.path.exists(full_path):
            data = process_image(full_path, output_dir)
            if data:
                results.append(data)
        else:
            print(f"  [Vision] Image not found: {img_name}")
            
    return results

if __name__ == "__main__":
    input_dir = "input"
    if os.path.isdir(input_dir):
        images = sorted(f for f in os.listdir(input_dir) if f.lower().endswith(('.jpg', '.png', '.jpeg')))
        process_all_images(images, base_dir=input_dir)
