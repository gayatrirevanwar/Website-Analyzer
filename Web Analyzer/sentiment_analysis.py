#!/usr/bin/env python
# coding: utf-8

# In[5]:


import nltk
import os
import pandas as pd
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
from textblob import TextBlob

nltk.download('punkt')
nltk.download('stopwords')
nltk.download('vader_lexicon')

positive_words_file = 'C:\\Users\\gayat\\OneDrive\\Desktop\\BlackCoffer\\MasterDictionary\\positive-words.txt'
negative_words_file = 'C:\\Users\\gayat\\OneDrive\\Desktop\\BlackCoffer\\MasterDictionary\\negative-words.txt'
stop_words_folder = 'C:\\Users\\gayat\\OneDrive\\Desktop\\BlackCoffer\\StopWords'

with open(positive_words_file, 'r', encoding='latin1') as f:
    positive_words = set(f.read().splitlines())

with open(negative_words_file, 'r', encoding='latin1') as f:
    negative_words = set(f.read().splitlines())

for filename in os.listdir(stop_words_folder):
    if filename.endswith('.txt'):
        with open(stop_words_file, 'r', encoding='latin1') as f:
            stop_words = set(f.read().splitlines())

def calculate_metrics(text):
    blob = TextBlob(text)
    polarity = blob.sentiment.polarity
    subjectivity = blob.sentiment.subjectivity

    words = word_tokenize(text)
    sentences = sent_tokenize(text)

    word_count = len(words)
    sentence_count = len(sentences)
    avg_sentence_len = sum(len(sent.split()) for sent in sentences) / sentence_count if sentence_count > 0 else 0
    percentage_complex_words = sum(1 for word in words if word.lower() not in stop_words and len(word) > 2) / word_count * 100
    fog_index = 0.4 * (avg_sentence_len + percentage_complex_words)
    avg_word_per_sentence = word_count / sentence_count if sentence_count > 0 else 0
    complex_word_count = sum(1 for word in words if word.lower() not in stop_words and len(word) > 2)
    syllables = sum(count_syllables(word) for word in words if isinstance(word, str))

    avg_word_length = sum(len(word) for word in words) / word_count if word_count > 0 else 0
    personal_pronouns = sum(1 for word in words if word.lower() in ['i', 'me', 'my', 'mine', 'myself', 'we', 'us', 'our', 'ours', 'ourselves'])

    positive_count = sum(1 for word in words if word.lower() in positive_words)
    negative_count = sum(1 for word in words if word.lower() in negative_words)

    return {
        'POSITIVE SCORE': positive_count,
        'NEGATIVE SCORE': negative_count,
        'POLARITY SCORE': polarity,
        'SUBJECTIVITY SCORE': subjectivity,
        'AVG SENTENCE LENGTH': avg_sentence_len,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_word_per_sentence,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllables / word_count if word_count > 0 else 0,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': avg_word_length
    }

def count_syllables(word):
    vowels = 'aeiouy'
    word = word.lower().strip(".:;?!")
    if len(word) == 0:
        return 0

    if word[-2:] == 'es' or word[-2:] == 'ed':
        if word[-3:] == 'ted' or word[-3:] == 'tes':
            return 2
        return 1

    count = 0
    prev_char_was_vowel = False
    for char in word:
        if char in vowels:
            if not prev_char_was_vowel:
                count += 1
            prev_char_was_vowel = True
        else:
            prev_char_was_vowel = False

    if word[-1] == 'e' and word[-2:] != 'le' and word[-2:] != 'be':
        count -= 1

    if count == 0:
        count = 1

    return count

folder_path = 'C:\\Users\\gayat\\OneDrive\\Desktop\\BlackCoffer\\web_scraping'

results = []

for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        with open(os.path.join(folder_path, filename), 'r', encoding='utf-8') as file:  # Specify encoding here
            text = file.read()
            metrics = calculate_metrics(text)
            results.append(metrics)

df = pd.DataFrame(results)

existing_excel_file = "C:\\Users\\gayat\\OneDrive\\Desktop\\BlackCoffer\\Output Data Structure.xlsx"
existing_df = pd.read_excel(existing_excel_file)

start_row = 0 
for key, values in df.items():
    existing_df[key] = pd.Series(values)

existing_df.to_excel(existing_excel_file, index=False, startrow=start_row, header=True)

print("Data appended successfully.")


# In[ ]:




