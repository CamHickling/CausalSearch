import fitz  # PyMuPDF
import spacy
from spacy.matcher import Matcher
import pandas as pd
import textwrap
import parse_wordbank
import save_to_file
import re


# Load the spaCy model
nlp = spacy.load("en_core_web_sm")

# Function to extract text from a PDF file
def generate_patterns(phrases):
    patterns = []
    for phrase in phrases:
        doc = nlp(phrase)
        pattern = [{"LEMMA": token.lemma_} for token in doc if not token.is_punct and not token.is_space]
        patterns.append(pattern)
    return patterns
    
# Examples phrase patterns for testing
'''
# Define causal and correlational patterns
causal_patterns = [
    [{"LEMMA": "cause"}],
    [{"LEMMA": "lead"}, {"LEMMA": "to"}],
    [{"LEMMA": "result"}, {"LEMMA": "in"}],
    [{"LEMMA": "because"}],
    [{"LEMMA": "due"}, {"LEMMA": "to"}],
    [{"LEMMA": "affect"}],
    [{"LEMMA": "influence"}]
]

mediation_patterns = [
    [{"LEMMA": "correlate"}, {"LEMMA": "with"}],
    [{"LEMMA": "associate"}, {"LEMMA": "with"}],
    [{"LEMMA": "link"}, {"LEMMA": "to"}],
    [{"LEMMA": "relate"}, {"LEMMA": "to"}],
    [{"LEMMA": "connect"}, {"LEMMA": "with"}]
]
'''

def clean_illegal_characters(df):
    def remove_illegal_chars(value):
        # Convert value to string and then remove illegal characters
        value_str = str(value)
        return re.sub(r'[^\x20-\x7E]', '', value_str)  # Keep only printable ASCII characters
    
    # Apply the function to each element in the DataFrame
    return df.apply(lambda col: col.apply(remove_illegal_chars))

def get_phrase(sentence):
    phrase = []
    for word in sentence.split():
        if word.isupper():
            phrase.append(word)
    return " ".join(phrase)

# Function to extract sentences with phrases based on patterns
def extract_sentences_with_phrases(doc, patterns):
    matcher = Matcher(nlp.vocab)
    for pattern in patterns:
        matcher.add("PATTERN", [pattern])
    matches = matcher(doc)
    sentences = []
    for match_id, start, end in matches:
        phrase_span = doc[start:end]
        sentence = phrase_span.sent
        key_phrase = phrase_span.text
        # Highlight the key phrase in the sentence
        highlighted_sentence = sentence.text.replace(key_phrase, key_phrase.upper())
        sentences.append(highlighted_sentence)
    return sentences

# Extract text from the PDF file
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()
    return text

def search_for_keyphrases(pdf_path, wordbank_path):
    # Extract text from the PDF
    scientific_paper_text = extract_text_from_pdf(pdf_path)

    # Process the text with spaCy
    doc = nlp(scientific_paper_text)

    # Extract causal and mediation phrases
    causal_phrases, mediation_phrases = parse_wordbank.extract_phrases(wordbank_path)

    #print("\nGenerating causal patterns")
    # Generate causal patterns
    causal_patterns = {}
    for key in causal_phrases:
        #print("causal phrases:", causal_phrases[key])
        patterns = generate_patterns(causal_phrases[key])
        #print("patterns:")
        #for p in patterns:
        #    print(p)
        causal_patterns[key] = causal_patterns.get(key, []) + patterns
        #print("causal_patterns: ", causal_patterns)
        
    #print("\nGenerating mediation patterns:")
    # Generate correlational patterns
    mediation_patterns = {}
    for key in mediation_phrases:
        #print("mediation phrases:", mediation_phrases[key])
        patterns = generate_patterns(mediation_phrases[key])
        #print("patterns:")
        #for p in patterns:
        #    print(p)
        mediation_patterns[key] = mediation_patterns.get(key, []) + patterns
        #print("mediation_patterns:", mediation_patterns)
        

    # Extract sentences with causal and mediation phrases
    causal_sentences = {key : extract_sentences_with_phrases(doc, causal_patterns[key]) for key in causal_patterns}
    mediation_sentences = {key : extract_sentences_with_phrases(doc, mediation_patterns[key]) for key in mediation_patterns}

    return causal_sentences, mediation_sentences

#!!! THIS FUNCTION IS IN PROGRESS
def to_dataframe(causal_sentences, mediation_sentences, criteria):
     # Initialize an empty DataFrame with the required columns
    df = pd.DataFrame(columns=criteria)

    for key in causal_sentences:
        for sentence in causal_sentences[key]:
            # Create a dictionary for the new row
            row_dict = {
                'paper_no': 1,
                'year': 2016,
                'coder': 'CH',
                'statement': get_phrase(sentence),
                'no_occurrences': 1,
                'section': 'section',
                'causal': 'causal',
                'causal_type': key,
                'valence': 'who knows',
                'full_sentence': sentence
            }
            # Add any criteria columns that are not in the row_dict
            row_dict.update({k: None for k in criteria if k not in row_dict})
            # Append the new row to the DataFrame
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)

    for key in mediation_sentences:
        for sentence in mediation_sentences[key]:
            # Create a dictionary for the new row
            row_dict = {
                'paper_no': 1,
                'year': 2016,
                'coder': 'CH',
                'statement': get_phrase(sentence),
                'no_occurrences': 1,
                'section': 'section',
                'causal': 'mediation',
                'causal_type': key,
                'valence': 'who knows',
                'full_sentence': sentence
            }
            # Add any criteria columns that are not in the row_dict
            row_dict.update({k: None for k in criteria if k not in row_dict})
            # Append the new row to the DataFrame
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)

    df_cleaned = clean_illegal_characters(df)
    return df_cleaned

if __name__ == "__main__":
    # Path to the PDF file
    pdf_path = r"C:\Users\camhi\Desktop\Bhargave, Montgomery, Redden.pdf"
    wordbank_path = r"C:\Users\camhi\Desktop\wordbanktest.docx"
    
    causal_criteria = ['paper_no', 'year', 'coder', 'statement', 'no_occurrences', 
                    'section', 'causal', 'causal_type', 'valence', 'full_sentence']

    pieters_criteria = ['power', 'c_power', 'reliability', 'c_reliability', 'confounds', 
                        'd_methods', 'a_confounds', 'sens_analysis', 't_basis', 
                        'l_analysis', 'sem', 'sem_reliability', 'limitations', 
                        'limitations_reliability', 'limitations_confounds', 
                        'limitations_methods', 'limitations_analysis', 
                        'limitations_sem', 'limitations_sem_reliability']

    causal_sentences, mediation_sentences = search_for_keyphrases(pdf_path, wordbank_path)

    '''
    # Print the extracted sentences
    print("\nCausal Sentences:")
    for causal_type in causal_sentences:
        causal_type = causal_type.replace("\n", " ")  
        print(textwrap.fill(f"- {causal_type}", width=75, initial_indent="\n", subsequent_indent="\t"))

        for sentence in causal_sentences[causal_type]:
            sentence = sentence.replace("\n", " ")
            print(textwrap.fill(f"- {sentence}", width=75, initial_indent="\n\t", subsequent_indent="\t"))

    print("\nMediation Sentences:")
    for mediation_type in mediation_sentences:
        meditation_type = mediation_type.replace("\n", " ")
        print(textwrap.fill(f"- {mediation_type}", width=75, initial_indent="\n", subsequent_indent="\t"))

        for sentence in mediation_sentences[mediation_type]:
            sentence = sentence.replace("\n", " ")
            print(textwrap.fill(f"- {sentence}", width=75, initial_indent="\n\t", subsequent_indent="\t"))
    '''

    # Create a dataframe
    df = to_dataframe(causal_sentences, mediation_sentences, causal_criteria + pieters_criteria)
    save_to_file.save(r"C:\Users\camhi\Desktop\formatted_data.xlsx", df)
    