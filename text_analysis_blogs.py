import codecs
import csv
import os
import re
import nltk
import pandas as pd
import textstat
from bs4 import BeautifulSoup
from nltk.tokenize import word_tokenize
from selenium import webdriver

# load the selenium chrome driver to run the URLs
driver = webdriver.Chrome("C:/Users/ASUS/OneDrive/Desktop/chromedriver")

# Read the input.xlsx file
df = pd.read_excel("Input.xlsx")
urls = dict()

# Iterate over the rows in the DataFrame
for index, row in df.iterrows():
    # Get the URL and URL ID from the row
    url = row['URL']
    url_id = row['URL_ID']
    urls[url_id] = url
    driver.get(url)

    # Send a GET request to the URL and retrieve the HTML
    html = driver.page_source

    # Use BeautifulSoup to parse the HTML
    soup = BeautifulSoup(html, 'html.parser')
    
    # Extract the article text from the HTML
    article_text = soup.find(attrs={'class': "td-post-content"})
    
    # Save the article text to a file with the URL ID as the file name
    try:
        with open(f'{url_id}.txt', 'w') as file:
            file.write(article_text.get_text())
    except AttributeError:
        print(f'Error: Could not find article text for {url}')
    except UnicodeEncodeError:
        print(f'Error: Could not write text to file for {url}')

# Set the path to the new directory
dir_path = 'articles'

# Create the directory if it does not exist
if not os.path.exists(dir_path):
    os.mkdir(dir_path)

# Iterate over the text files
for file in os.listdir():
    # Check if the file is a text file
    if file.endswith('.txt') and file[:-4].isdigit():
        # Set the file path
        file_path = file

        # Delete the file if it exists in the new directory
        if os.path.exists(f'{dir_path}/{file_path}'):
            os.remove(f'{dir_path}/{file_path}')
        
        # Move the file to the new directory
        os.rename(file_path, f'{dir_path}/{file_path}')

# Include all stopword files in the directory
stop_word_files = ['StopWords_Names.txt', 'StopWords_Auditor.txt', 'StopWords_Generic.txt', 'StopWords_Currencies.txt',
                   'StopWords_Geographic.txt', 'StopWords_DatesandNumbers.txt', 'StopWords_GenericLong.txt']
stop_words = set()
for file in stop_word_files:
    with open(file) as f:
        stop_words.update(f.read().split())

# Add the NLTK stop words to the stop_words set
stop_words.update(nltk.corpus.stopwords.words('english'))

# Create an empty dictionary
words = {}

# Read in the positive words and add them to the dictionary
with open('positive-words.txt', 'r') as f:
    for word in f.read().split():
        if word not in stop_words:
            words[word] = 'positive'

# Read in the negative words and add them to the dictionary
with open('negative-words.txt', 'r') as f:
    for word in f.read().split():
        if word not in stop_words:
            words[word] = 'negative'

# function to count syllables
def count_syllables(text):
    syllable_count = 0
    for word in text.split():
        syllable_count += textstat.syllable_count(word)
    return syllable_count

sentences = 0
complex_words = 0
total_personal_pronouns = 0
average_word_length = 0
tokens = []
results = []

# Iterate over the text files in the articles directory
for file in os.listdir('articles'):
    # Check if the file is a text file
    if file.endswith('.txt'):
        # Read in the text file
        with codecs.open(f'articles/{file}', 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

            # Tokenize the text
            tokens = word_tokenize(text)
            
            # Count the number of personal pronouns in the text
            # Pronouns extraction
            pronounRegex = re.compile(r'\b(I|we|my|ours|(?-i:us))\b', re.I)
            for token in tokens:
                pronouns = pronounRegex.findall(token)
                total_personal_pronouns += len(pronouns)
            
            # Remove stop words and punctuation from the list of tokens
            tokens = [token for token in tokens if token.lower() not in stop_words and token.isalpha()]
            
            # Initialize the counters for positive and negative words
            positive_score = 0
            negative_score = 0
            
            # Iterate over the tokens and calculate the positive score and negative score
            total_syllables = 0
            for token in tokens:
                if token.lower() in words:
                    if words[token.lower()] == 'positive':
                        positive_score += 1
                    elif words[token.lower()] == 'negative':
                        negative_score += 1
                total_syllables += count_syllables(token)
            total_words = len(tokens)
            
            # Calculate the polarity score
            polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
            
            # Calculate the subjectivity score
            subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
            
            # Count the number of sentences
            sentence_count = textstat.sentence_count(text)
            
            # Calculate the average number of words per sentence
            avg_words_per_sentence = total_words / sentence_count
            
            # Count the number of complex words in the text
            complex_words = [token for token in tokens if count_syllables(token) > 2]
            total_complex_words = len(complex_words)
            
            # Calculate the average sentence length
            avg_sentence_length = len(tokens) / (sentence_count + 0.000001)
            
            # Calculate the percentage of complex words
            percent_complex_words = total_complex_words / (len(tokens) + 0.000001) * 100
            
            # Calculate the fog index
            fog_index = 0.4 * (avg_sentence_length + percent_complex_words)
            
            # Calculate the average syllable count per word
            average_syllables_per_word = total_syllables / (total_words + 0.000001)
            
            # Calculating average word length
            average_word_length = sum(len(token) for token in tokens) / (len(tokens) + 0.000001)
            
            # retrieving the url id from file name
            urlid = int(file[:-4])
            
            # calling the urls dictionary
            url = urls[urlid]
            results.append(
                [urlid, url, positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
                 percent_complex_words, fog_index,
                 avg_words_per_sentence, len(complex_words), len(tokens), average_syllables_per_word,
                 total_personal_pronouns, average_word_length])

            # reset
            total_personal_pronouns = 0

# writing the output in a csv file
with open('output.csv', 'w', newline='') as csv_file:
    # Create a CSV writer
    csv_writer = csv.writer(csv_file)
    
    # Write the header row
    csv_writer.writerow(['URL ID', 'URL', 'Positive score', 'Negative score', 'Polarity Score', 'Subjectivity Score',
                         'Average Sentence Length', 'Percent Complex Words', 'Fog Index',
                         'Average Number of Words per Sentence', 'Complex Words', 'Word Count',
                         'Average Syllables per Word', 'Total Personal Pronouns', 'Average Word Length'])
    # sort results
    sortedList = sorted(results, key=lambda row: row[0])
    csv_writer.writerows(sortedList)
