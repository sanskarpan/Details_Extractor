import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import nltk
import re

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('wordnet')

input_file = 'input.xlsx'
output_structure_file = 'Output Data Structure.xlsx'
df_in = pd.read_excel(input_file)
df_out_structure = pd.read_excel(output_structure_file)

positive_words = set(open('positive-words.txt').read().splitlines())
negative_words = set(open('negative-words.txt').read().splitlines())

stopwords_files = [
    'StopWords_Auditor.txt',
    'StopWords_Currencies.txt',
    'StopWords_DatesandNumbers.txt',
    'StopWords_Generic.txt',
    'StopWords_GenericLong.txt',
    'StopWords_Geographic.txt',
    'StopWords_Names.txt'
]
stopwords = set()
for f in stopwords_files:
    stopwords.update(set(open(f).read().splitlines()))

def clean(txt):
    txt = re.sub(r'[^\w\s]', '', txt)
    return txt

def syllable_count(word):
    vowels = "aeiou"
    word = word.lower()
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith('e'):
        count -= 1
    if count == 0:
        count += 1
    return count

def calculate(article_text):
    tokens = nltk.word_tokenize(article_text)
    words = [w.lower() for w in tokens if w.isalpha() and w.lower() not in stopwords]
    total_words = len(words)
    total_sentences = len(nltk.sent_tokenize(article_text))
    syllables = sum(syllable_count(w) for w in words)
    complex_words = [w for w in words if syllable_count(w) > 2]
    total_complex_words = len(complex_words)
    total_personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', article_text, flags=re.IGNORECASE))
    total_characters = sum(len(w) for w in words)
    positive_score = sum(1 for w in words if w in positive_words)
    negative_score = sum(1 for w in words if w in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    average_sentence_length = total_words / total_sentences
    percentage_complex_words = total_complex_words / total_words
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
    average_words_per_sentence = total_words / total_sentences
    average_word_length = total_characters / total_words
    return positive_score, negative_score, polarity_score, subjectivity_score, average_sentence_length, \
           percentage_complex_words, fog_index, average_words_per_sentence, total_complex_words, total_words, \
           syllables, total_personal_pronouns, average_word_length

output_dir = 'extracted_articles'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

output_data = []
for i, r in df_in.iterrows():
    url_id = r['URL_ID']
    url = r['URL']
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    article_text = soup.get_text()
    cleaned_article_text = clean(article_text)
    variables = calculate(cleaned_article_text)
    output_row = [url_id, url] + list(variables)
    output_data.append(output_row)

output_columns = ['URL ID', 'URL Link', 'Positive Score', 'Negative Score', 'Polarity Score', 'Subjectivity Score',
                  'Average Sentence Length', 'Percentage of Complex Words', 'Fog Index', 'Average Number of Words Per Sentence',
                  'Complex Word Count', 'Word Count', 'Syllable Per Word', 'Personal Pronouns', 'Average Word Length']
df_out = pd.DataFrame(output_data, columns=output_columns)

output_file = 'output_assignment.xlsx'
df_out.to_excel(output_file, index=False)