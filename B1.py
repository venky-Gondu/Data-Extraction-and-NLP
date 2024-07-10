#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# web Scrapping from website url present in xlsl file uisng beatifulsoup
get_ipython().system('pip install pandas requests beautifulsoup4 openpyxl')


# In[ ]:


import pandas as pd
import requests
from bs4 import BeautifulSoup
import openpyxl
from urllib.parse import urlparse

def extract_main_content(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.decompose()

        # Try to find main content
        main_content = None
        potential_content_tags = ['article', 'main', '[role="main"]', '.main-content', '#main-content']

        for tag in potential_content_tags:
            main_content = soup.select_one(tag)
            if main_content:
                break

        if not main_content:
            # If no main content container found, use the body
            main_content = soup.body

        # Remove potential non-article elements
        for elem in main_content.select('header, footer, nav, aside, .sidebar, .comments'):
            elem.decompose()

        # Extract text from remaining paragraphs
        paragraphs = main_content.find_all('p')
        text = ' '.join([p.get_text(strip=True) for p in paragraphs])

        return text.strip()
    except Exception as e:
        return f"Error: {str(e)}"

def process_urls_from_excel(input_file, output_file):
    # Read URLs from Excel file
    df = pd.read_excel(input_file)

    # Ensure there's a 'URL' column and a 'URL_ID' column
    if 'URL' not in df.columns or 'URL_ID' not in df.columns:
        raise ValueError("Excel file must contain both 'URL' and 'URL_ID' columns")

    results = []

    # Process each URL
    for _, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']

        text = extract_main_content(url)
        results.append({'URL_ID': url_id, 'URL': url, 'Text': text})

    # Create a new DataFrame with results
    output_df = pd.DataFrame(results)

    # Write results to a new Excel file
    output_df.to_excel(output_file, index=False)
    print(f"Results written to {output_file}")

# Usage
input_file = 'Input.xlsx'
output_file = 'output.xlsx'
process_urls_from_excel(input_file, output_file)


# In[ ]:


# display Xlsx file in pandas frame
df = pd.read_excel('output.xlsx')
df


# In[ ]:


df['Text'][0]


# In[ ]:


# text analysis
get_ipython().system('pip install textblob')


# In[ ]:


# Sentiment analysis
# Cleaning the text using Stopwords and Punction marks
import re
import nltk
from nltk.corpus import stopwords
from textblob import TextBlob
# use custom stop words or stopword file
nltk.download('stopwords')


# In[ ]:


import os
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
    return chardet.detect(raw_data)['encoding']

def load_stopwords(folder_path):
    stopwords = set()

    if not os.path.exists(folder_path):
        print(f"Error: The folder {folder_path} does not exist.")
        return stopwords

    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            try:
                # Detect the file encoding
                encoding = detect_encoding(file_path)

                # If detection fails, try common encodings
                if not encoding:
                    encodings_to_try = ['utf-8', 'iso-8859-1', 'windows-1252']
                else:
                    encodings_to_try = [encoding]

                for enc in encodings_to_try:
                    try:
                        with open(file_path, 'r', encoding=enc) as file:
                            words = file.read().split()
                            stopwords.update(word.strip().lower() for word in words)
                        print(f"Successfully read {filename} with {enc} encoding")
                        break
                    except UnicodeDecodeError:
                        if enc == encodings_to_try[-1]:
                            print(f"Failed to read {filename} with all attempted encodings")
                        continue

            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")

    print(f"Loaded {len(stopwords)} unique stopwords from {folder_path}")
    return stopwords

# Usage
stopwords_folder = '/content/drive/MyDrive/Blackcoffer/StopWords/'  # Replace with the actual path to your stopwords folder
stopwords = load_stopwords(stopwords_folder)

# You can now use the 'stopwords' set in your text analysis


# In[ ]:


print(stopwords)


# In[ ]:


# find particular stopwrod is present or not
'cedi' in stopwords # currency stop word


# In[ ]:


# read the positive words from the positive txt file and store them in positive words
positive_words = set()
with open('/content/drive/MyDrive/Blackcoffer/MasterDictionary/positive-words.txt', 'r', encoding='latin-1') as file: # Try 'latin-1' encoding
    for line in file:
        positive_words.add(line.strip())
# for negative words
negative_words = set()
with open('/content/drive/MyDrive/Blackcoffer/MasterDictionary/negative-words.txt', 'r', encoding='latin-1') as file: # Try 'latin-1' encoding
    for line in file:
        negative_words.add(line.strip())


# In[ ]:


print(positive_words)


# In[ ]:


'accolade' in positive_words


# In[ ]:


# remove the stopwords from Text in output xlsx
def remove_stopwords(text):
    words = text.split()
    filtered_words = [word for word in words if word.lower() not in stopwords]
    return ' '.join(filtered_words)
# pass the Text from column wise to reove stop words


# In[ ]:


df['Text'] = df['Text'].apply(remove_stopwords)
df


# In[ ]:


df['Text'][0]


# In[ ]:


#  find any Stopword present in df[text][0]
for word in df['Text'][0].split():
    if word.lower() in stopwords:
        print(word)
    else:
        print('no')


# In[ ]:


# postive_word count for Text In df
def count_positive_words(text):
    count = 0
    for word in text.split():
        if word.lower() in positive_words:
            count += 1
    return count
df['Positive_Score'] = df['Text'].apply(count_positive_words)
df


# In[ ]:


# calculate negative_words
def count_negative_words(text):
    count = 0
    for word in text.split():
        if word.lower() in negative_words:
            count -= 1
    return count
df['Negative_Score'] = df['Text'].apply(count_negative_words)
df


# In[ ]:


# calculate popularity score
df['Popularity_Score'] = (df['Positive_Score'] - df['Negative_Score']) / (df['Positive_Score'] + df['Negative_Score'] + 0.000001)
df


# In[ ]:


# subjective score calculation
df['Word_Count'] = df['Text'].apply(lambda x: len(x.split()))
df['Subjective_Score'] = (df['Positive_Score'] + df['Negative_Score']) / (df['Word_Count'] + 0.000001)
df


# In[ ]:


# average sentence length clculation
df['Sentence_Count'] = df['Text'].apply(lambda x: len(re.split(r'[.!?]', x)))
df['Complex_Words'] = df['Text'].apply(lambda x: len([word for word in x.split() if len(word) > 2]))
df['Average_Sentence_Length'] = df['Word_Count'] / (df['Sentence_Count'] + 0.000001)
df['percentage_of_complex_words'] = (df['Complex_Words'] / (df['Word_Count'] + 0.000001)) * 100
df['fox_index'] = 0.4 * (df['Average_Sentence_Length'] + df['percentage_of_complex_words'])
df


# In[ ]:


df['average-number-of-words-per-sentence'] = df['Word_Count'] / df['Sentence_Count']


# In[ ]:


df


# In[ ]:



df['char'] = df['Text'].apply(lambda x: len(x))
df['Averge_Word_Length'] = df['char'] / (df['Word_Count'] + 0.000001)
df


# In[ ]:


#  extract personal pronouns
import re

def count_personal_pronouns(text):
    # List of personal pronouns to search for
    pronouns = [
        "I", "me", "my", "mine", "myself",
        "you", "your", "yours", "yourself", "yourselves",
        "he", "him", "his", "himself",
        "she", "her", "hers", "herself",
        "it", "its", "itself",
        "we", "us", "our", "ours", "ourselves",
        "they", "them", "their", "theirs", "themselves"
    ]

    # Convert text to lowercase for case-insensitive matching
    text = text.lower()

    # Count all pronouns in the text
    total_count = 0
    for pronoun in pronouns:
        # Use word boundaries to ensure we're matching whole words
        count = len(re.findall(r'\b' + re.escape(pronoun) + r'\b', text))

        # Handle exceptions (currently only for "us")
        if pronoun == "us":
            # Subtract occurrences of "US" referring to United States
            us_as_country = len(re.findall(r'\b(united states|u\.s\.)\b', text))
            count = max(0, count - us_as_country)

        total_count += count

    return total_count

df['Personal_Pronouns'] = df['Text'].apply(count_personal_pronouns)
df


# In[ ]:


import nltk
from nltk.corpus import cmudict

nltk.download('cmudict', quiet=True)
d = cmudict.dict()

def count_syllables(word):
    word = word.lower()

    # Check if the word is in the CMU dictionary
    if word in d:
        return max([len([y for y in x if y[-1].isdigit()]) for x in d[word]])

    # If the word is not in the dictionary, use the fallback method
    return fallback_syllable_count(word)

def fallback_syllable_count(word):
    word = word.lower()
    count = 0
    vowels = 'aeiouy'

    # Handle special cases
    if len(word) <= 3:
        return 1

    # Handle common endings
    if word.endswith('es') or word.endswith('ed'):
        # Remove 'es' or 'ed'
        word = word[:-2]
    elif word.endswith('e'):
        # Remove 'e' unless the word ends with 'le'
        if not word.endswith('le'):
            word = word[:-1]

    # Count vowel groups
    prev_char_was_vowel = False
    for char in word:
        if char in vowels:
            if not prev_char_was_vowel:
                count += 1
            prev_char_was_vowel = True
        else:
            prev_char_was_vowel = False

    # Handle special cases where counting vowel groups doesn't work well
    if word.endswith('le') and len(word) > 2 and word[-3] not in vowels:
        count += 1

    # Ensure at least one syllable
    return max(1, count)

df['Syllable_Count'] = df['Text'].apply(count_syllables)
df


# In[ ]:


# convert the df into xlsx file save to google drive
df.to_excel('Output_Data_Structure.xlsx', index=False)
from google.colab import files
files.download('Output_Data_Structure.xlsx')

