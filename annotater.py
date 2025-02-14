import os
import fitz  # PyMuPDF
import pandas as pd
from google import genai
import time

# Initialize Gemini API client
client = genai.Client(api_key="AIzaSyCeZ8FVMarfFc2NQLT8CeoMYWK2I_0wmIA")  # Replace with your actual API key

# Folder containing PDFs
pdf_folder = "paper"
output_file = "pdf_metadata.xlsx"

# Allowed labels
allowed_labels = {"Deep Learning", "Computer Vision", "Reinforcement Learning", "NLP", "Optimization"}

# Function to extract metadata using Gemini API
def extract_metadata(text):
    prompt = f"""
    You are an expert in analyzing research papers. Extract the following details:
    - Title
    - Authors
    - University
    - 3 best-suited topic labels (must be exactly 3 from: Deep Learning, Computer Vision, Reinforcement Learning, NLP, Optimization)

    If any information is missing, return 'Unknown'. Ensure exactly 3 labels are returned.

    Extract from this text:
    {text[:1000]}  # Limiting input to avoid long processing times
    """

    try:
        response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
        if response and hasattr(response, "text"):
            return response.text
    except Exception as e:
        print(f"API error: {e}")
    
    return "Unknown"

# Extract PDFs and get metadata
data = []
for root, dirs, files in os.walk(pdf_folder):
    for file in files:
        if file.endswith(".pdf"):
            pdf_path = os.path.join(root, file)

            # Extract year from folder name
            year = "Unknown"
            parts = root.split(os.sep)
            for part in parts:
                if part.isdigit() and len(part) == 4:  # Checking if the folder name is a year
                    year = part
                    break

            try:
                # Extract text from PDF
                doc = fitz.open(pdf_path)
                text = "\n".join([page.get_text("text") for page in doc])

                # Get metadata from Gemini
                metadata = extract_metadata(text)
                time.sleep(25)  # Delay to prevent hitting rate limits

                # Parse metadata response (assuming structured output)
                title, authors, university, labels = "Unknown", "Unknown", "Unknown", "Unknown"

                lines = metadata.split("\n")
                for line in lines:
                    if "Title:" in line:
                        title = line.replace("Title:", "").strip()
                    elif "Authors:" in line:
                        authors = line.replace("Authors:", "").strip()
                    elif "University:" in line:
                        university = line.replace("University:", "").strip()
                    elif "Labels:" in line:
                        raw_labels = line.replace("Labels:", "").strip()
                        labels_list = [lbl.strip() for lbl in raw_labels.split(",") if lbl.strip() in allowed_labels]

                        # Ensure exactly 3 labels are returned
                        if len(labels_list) == 3:
                            labels = ", ".join(labels_list)
                        elif len(labels_list) > 3:
                            labels = ", ".join(labels_list[:3])
                        else:
                            labels = "Unknown"

                # Store data (saving full text instead of PDF path)
                data.append([text, title, authors, year, university, labels])

            except Exception as e:
                print(f"Error processing {file}: {e}")

# Save data to XLSX
df = pd.DataFrame(data, columns=["PDF Text", "Title", "Authors", "Year", "University", "Labels"])
df.to_excel(output_file, index=False)

print(f"Metadata saved to {output_file}")