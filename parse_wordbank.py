import docx
import os

def clean_text(text):
    return [p.strip() for p in text.split("\n")]

def extract_phrases(path):
    # Load the DOCX file
    doc = docx.Document(path)

    # Define dictionaries to store the phrases
    causal_phrases = {
        "Strong Causal": [],
        "Soft Causal": [],
        "Hedged Causal": [],
        "Direct Causal": []
    }

    mediation_phrases = {
        "Causal": [],
        "[Correlational] Descriptive": []
    }

    # Iterate over tables in the document
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            cells = row.cells
            
            # Skip the catergory titles
            if i == 0:
                continue
            # Check for causal language table
            if len(cells) == 4:
                #causal_phrases["Strong Causal"] = causal_phrases.get("Strong Causal", [])
                causal_phrases["Strong Causal"].extend(clean_text(cells[0].text.strip()))

                #causal_phrases["Soft Causal"] = causal_phrases.get("Soft Causal", [])
                causal_phrases["Soft Causal"].extend(clean_text(cells[1].text.strip()))
                
                #causal_phrases["Hedged Causal"] = causal_phrases.get("Hedged Causal", [])
                causal_phrases["Hedged Causal"].extend(clean_text(cells[2].text.strip()))
                
                #causal_phrases["Direct Causal"] = causal_phrases.get("Direct Causal", [])
                causal_phrases["Direct Causal"].extend(clean_text(cells[3].text.strip()))
            # Check for mediation language table
            elif len(cells) == 2:
                mediation_phrases["Causal"].extend(clean_text(cells[0].text.strip()))
                mediation_phrases["[Correlational] Descriptive"].extend(clean_text(cells[1].text.strip()))

    # Remove empty strings from the lists
    causal_phrases = {k: [phrase for phrase in v if phrase] for k, v in causal_phrases.items()}
    mediation_phrases = {k: [phrase for phrase in v if phrase] for k, v in mediation_phrases.items()}

    return causal_phrases, mediation_phrases

if __name__ == "__main__":
    # Load the DOCX file
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
    file_name = '/wordbanktest.docx'
    file_path = desktop + file_name

    causal_phrases, mediation_phrases = extract_phrases(file_path)

    # Print the extracted phrases
    print("Causal Phrases:")
    for category, phrases in causal_phrases.items():
        print(f"{category}: {phrases}")

    print("\nMediation Phrases:")
    for category, phrases in mediation_phrases.items():
        print(f"{category}: {phrases}")
