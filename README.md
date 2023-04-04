# Text-Analysis
This is a Python code that reads URLs from an Excel file, navigates to the pages using Selenium, scrapes the article text using Beautiful Soup, saves the text to individual files, and then analyzes the text using various Natural Language Processing techniques to calculate metrics such as subjectivity score, polarity score, average word length, and more.

The first part of the code reads the URLs from an Excel file, navigates to the pages using Selenium, scrapes the article text using Beautiful Soup, and saves the text to individual files. The second part of the code reads in the text files, tokenizes the text, removes stop words and punctuation, and then calculates various metrics using the text.

The code uses the NLTK and textstat libraries for Natural Language Processing and Beautiful Soup for HTML parsing. It also uses regular expressions to extract personal pronouns and count the number of syllables in each word.

Overall, this code provides a powerful way to analyze text data and extract insights from it. It could be useful for sentiment analysis, content analysis, or any other application that requires analyzing large amounts of text data.
