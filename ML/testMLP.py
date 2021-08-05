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

# read the csv file
df = pd.read_csv("email2.csv")

# Split data up into training and testing sets

x = df[['HtmlExists','DomainCount','DotCount',
                           'AccountExists','PaypalExists','LoginExists','BankExists']]
target = df['Class']

x_train, x_test, y_train, y_test = train_test_split(x,
                                                    target,
                                                    test_size=0.20,
                                                    random_state=0)

#Show number of rows and columns in test and training sets
print("x_train", x_train.shape, "x_test", x_test.shape, "y_train", y_train.shape, "y_test", y_test.shape)

#Train the model
from sklearn.neural_network import MLPClassifier

desc_classifier = MLPClassifier(solver='lbfgs', alpha=1e-5, hidden_layer_sizes=(5, 2)
                                , random_state=1)

desc_classifier.fit(x_train, y_train)

#Visualize the results of the model and performance

import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import confusion_matrix
from sklearn.metrics import plot_confusion_matrix

predicted_results = desc_classifier.predict(x_test)

cm_p = confusion_matrix(y_test, predicted_results)
print(cm_p)

disp = plot_confusion_matrix(desc_classifier, x_test, y_test,
                             cmap=plt.cm.Blues, values_format='d',
                             normalize=None)

disp.ax_.set_title('Confusion Matrix - MLP Classification')
plt.savefig('mlp.png')
plt.show()

from sklearn.metrics import accuracy_score
print('accuracy: ', accuracy_score(y_test, predicted_results))

#Save model
dump(desc_classifier, 'mlp.joblib')
print('SAVEEDDDDDDDDDDDDD')

