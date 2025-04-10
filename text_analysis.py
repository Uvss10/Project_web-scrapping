import os
import re
import pandas as pd
import string
from nltk.corpus import stopwords
from nltk.tokenize import TreebankWordTokenizer
import nltk

# Setup paths
INPUT_FILE = 'Input.xlsx'
TEXT_FOLDER = 'text_files'
STOPWORDS_DIR = 'StopWords'
DICTIONARY_DIR = 'MasterDictionary'
OUTPUT_FILE = 'Output Data Structure.xlsx'  # <-- updated

# Load stop words
def load_stopwords(folder):
    stop_words = set()
    for fname in os.listdir(folder):
        with open(os.path.join(folder, fname), 'r', encoding='latin-1') as f:
            stop_words.update([line.strip().lower() for line in f if line.strip()])
    return stop_words

# Load positive/negative words
def load_dict(path, fname):
    with open(os.path.join(path, fname), 'r', encoding='latin-1') as f:
        return set([line.strip().lower() for line in f if line.strip() and not line.startswith(';')])

# Syllable count helper
def count_syllables(word):
    word = word.lower()
    word = re.sub(r'(es|ed)$', '', word)
    return len(re.findall(r'[aeiouy]+', word))

# Complex word check
def is_complex(word):
    return count_syllables(word) > 2

# Pronoun counter
def count_pronouns(text):
    pronoun_pattern = r'\b(I|we|my|ours|us)\b'
    return len(re.findall(pronoun_pattern, text, re.I))

# Main text analysis
def analyze_text(text, pos_dict, neg_dict, stop_words):
    text_lower = text.lower()
    tokenizer = TreebankWordTokenizer()
    tokens = tokenizer.tokenize(text_lower)
    clean_tokens = [t for t in tokens if t not in stop_words and t not in string.punctuation]

    # Sentiment
    pos_score = sum(1 for t in clean_tokens if t in pos_dict)
    neg_score = sum(1 for t in clean_tokens if t in neg_dict)
    polarity_score = (pos_score - neg_score) / ((pos_score + neg_score) + 0.000001)
    subjectivity_score = (pos_score + neg_score) / (len(clean_tokens) + 0.000001)

    # Readability
    sentences = re.split(r'[.!?]+', text)
    sentences = [s.strip() for s in sentences if s.strip()]
    words = clean_tokens
    word_count = len(words)
    sent_count = len(sentences)
    avg_sent_len = word_count / (sent_count + 0.000001)
    complex_words = [w for w in words if is_complex(w)]
    perc_complex = len(complex_words) / (word_count + 0.000001)
    fog_index = 0.4 * (avg_sent_len + perc_complex)
    avg_words_per_sentence = avg_sent_len
    complex_word_count = len(complex_words)

    # Syllables and word length
    syllable_count = sum(count_syllables(w) for w in words)
    syllable_per_word = syllable_count / (word_count + 0.000001)
    avg_word_len = sum(len(w) for w in words) / (word_count + 0.000001)

    # Personal Pronouns
    pronoun_count = count_pronouns(text)

    return [
        pos_score,
        neg_score,
        polarity_score,
        subjectivity_score,
        avg_sent_len,
        perc_complex,
        fog_index,
        avg_words_per_sentence,
        complex_word_count,
        word_count,
        syllable_per_word,
        pronoun_count,
        avg_word_len
    ]

# Load everything
stop_words = load_stopwords(STOPWORDS_DIR)
positive_words = load_dict(DICTIONARY_DIR, 'positive-words.txt')
negative_words = load_dict(DICTIONARY_DIR, 'negative-words.txt')

# Load input
input_df = pd.read_excel(INPUT_FILE)
output_data = []

for _, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    filepath = os.path.join(TEXT_FOLDER, f"{url_id}.txt")

    if not os.path.exists(filepath):
        print(f"[!] Missing: {filepath}")
        continue

    with open(filepath, 'r', encoding='utf-8') as f:
        full_text = f.read()
        text_only = '\n'.join(full_text.split('\n')[1:])

    analysis = analyze_text(text_only, positive_words, negative_words, stop_words)
    output_data.append([url_id, url] + analysis)

# Output to Excel (updated from CSV)
columns = [
    'URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE',
    'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS',
    'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT',
    'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
]

output_df = pd.DataFrame(output_data, columns=columns)
output_df.to_excel(OUTPUT_FILE, index=False)
print(f"[✓] Analysis complete! Output saved to: {OUTPUT_FILE}")
