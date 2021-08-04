import re
import pandas as pd
from sklearn.model_selection import train_test_split
import nltk
from nltk.tokenize import word_tokenize
from nltk.stem import PorterStemmer
from nltk.corpus import stopwords
from IPython.display import display
from joblib import dump, load
import warnings
warnings.filterwarnings('ignore')


# Conduct simple cleaning on the desc variable, such that the descriptions:
#
# Contain only alphanumeric characters (use regular expression "[^0-9a-zA-Z\ ]")
# Are entirely lowercase
# Does not have leading or trailing spaces
# Also remove rows where desc is an empty string.
# Remove empty Strings

# read the csv file
df = pd.read_csv("email.csv")

def cleaning(phrase):
    phrase = re.sub("[^0-9a-zA-Z\ ]", "", str(phrase))
    phrase = phrase.lower()
    phrase = phrase.strip()

    return phrase


df['Text'] = df['Text'].apply(cleaning)
print('Before cleaning', df.shape)
df = df[df['Text'] != '']
display(df)

#word_tokenize to tokenize each description to a list of words.
#PorterStemmer to stem each word to get root form of a word.
#Join up the stemmed words as a string.

def stem(string):

    tokenized = word_tokenize(string)
    stemmed = []
    stemmer = PorterStemmer()

    for word in tokenized:
        stemmed.append(stemmer.stem(word))

    return ' '.join(stemmed)

# Actual DF
# Running the following code will take a while
df['text_stem'] = df['Text'].apply(stem)
display(df)

# Remove stopwords that have no meaning
#Create the list of stopwords in english
#Tokenize the each description
#Check for stopwords
#Join up non stopwords as a string

def remove_stopwords(string):
    STOP_WORDS = set(stopwords.words('english'))

    tokenized = word_tokenize(string)
    filtered = []

    for word in tokenized:
        if word not in STOP_WORDS:
            filtered.append(word)

    return " ".join(filtered)

df['text_cleaned'] = df['text_stem'].apply(remove_stopwords)
display(df)

# Split data up into training and testing sets
# put only the converted column 'text_cleaned' to new dataframe
text_cleaned = df[['text_cleaned']]
target = df['Class']

x_train, x_test, y_train, y_test = train_test_split(text_cleaned,
                                                    target,
                                                    test_size=0.20,
                                                    random_state=0)

#Show number of rows and columns in test and training sets
print("x_train", x_train.shape, "x_test", x_test.shape, "y_train", y_train.shape, "y_test", y_test.shape)

# In order to use Naive Bayes Classifier model, need to transform to 1s and 0s (known as one-hot encoding).

from sklearn.feature_extraction.text import CountVectorizer

vectorizer = CountVectorizer()
#transform training data
desc_matrix = vectorizer.fit_transform(x_train['text_cleaned'])

#Save vectorizer
dump(vectorizer, 'naivebayesVectorizer.joblib')
print("SAVEEDDDDDDDDDDDD")
#Train the model
from sklearn.naive_bayes import MultinomialNB


desc_classifier = MultinomialNB()

desc_classifier.fit(desc_matrix, y_train)

#Visualize the results of the model and performance

import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import confusion_matrix
from sklearn.metrics import plot_confusion_matrix

x_matrix = vectorizer.transform(x_test['text_cleaned'])

predicted_results = desc_classifier.predict(x_matrix)

cm_p = confusion_matrix(y_test, predicted_results)
print(cm_p)

disp = plot_confusion_matrix(desc_classifier, x_matrix, y_test,
                             cmap=plt.cm.Blues, values_format='d',
                             normalize=None)

disp.ax_.set_title('Confusion Matrix - Naive Bayes Classification of Loan Description')

plt.show()

#Save model
dump(desc_classifier, 'naivebayes.joblib')

#
# if output ==1:
#     result = 'Phishing'
# elif output == 0:
#     result = 'Non-Phishing'

