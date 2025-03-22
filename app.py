from flask import Flask, request, render_template
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.metrics import accuracy_score
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer
import os

nltk.download('stopwords')
nltk.download('punkt')
nltk.download('wordnet')

app = Flask(__name__)

# Preprocessing function
def preprocess_text(text):
    lemmatizer = WordNetLemmatizer()
    words = word_tokenize(text.lower())
    words = [lemmatizer.lemmatize(word) for word in words if word.isalpha()]
    return ' '.join(words)

# Load and preprocess the data
file_path = os.path.join(os.path.dirname(__file__), "movie_review.csv")
if not os.path.isfile(file_path):
    raise FileNotFoundError(f"The file at path {file_path} does not exist.")

data = pd.read_csv(file_path)
data['text'] = data['text'].apply(preprocess_text)

# Remove stopwords
stop_words = set(stopwords.words('english'))

# Split the dataset
X = data['text']
y = data['tag']
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Convert text to numerical data using TF-IDF
vectorizer = TfidfVectorizer(ngram_range=(1, 2), stop_words=list(stop_words), max_features=5000)
X_train_vectorized = vectorizer.fit_transform(X_train)
X_test_vectorized = vectorizer.transform(X_test)

# Hyperparameter tuning using GridSearchCV with limited parameters
param_grid = {'alpha': [0.5, 1.0]}
model = GridSearchCV(MultinomialNB(), param_grid, cv=3)
model.fit(X_train_vectorized, y_train)

# Calculate the model accuracy on the test set
accuracy = accuracy_score(y_test, model.predict(X_test_vectorized)) * 100
accuracy = f"{accuracy:.2f}"

# Function to predict sentiment of a sentence
def predict_sentence_sentiment(sentence):
    sentence_preprocessed = preprocess_text(sentence)
    sentence_vectorized = vectorizer.transform([sentence_preprocessed])
    return model.predict(sentence_vectorized)[0]

# Function to highlight sentences based on predicted sentiment
def highlight_sentences(text):
    sentences = text.split('. ')
    highlighted_text = ""
    for sentence in sentences:
        sentiment = predict_sentence_sentiment(sentence)
        if sentiment == "pos":
            highlighted_text += f'<span class="pos-sentence">{sentence}.</span> '
        elif sentiment == "neg":
            highlighted_text += f'<span class="neg-sentence">{sentence}.</span> '
        else:
            highlighted_text += f'<span class="neutral-sentence">{sentence}.</span> '
    return highlighted_text

@app.route('/')
def home():
    return render_template('index.html', prediction_accuracy=accuracy)

@app.route('/predict', methods=['POST'])
def predict():
    review = request.form['review']
    review_preprocessed = preprocess_text(review)
    review_vectorized = vectorizer.transform([review_preprocessed])
    prediction = model.predict(review_vectorized)[0]
    highlighted_text = highlight_sentences(review)
    analysis = "Your review was analyzed, and the prediction percentage indicates the confidence of the model in its classification."
    return render_template('index.html', prediction=prediction, prediction_accuracy=accuracy, review_text=review, analysis=analysis, highlighted_text=highlighted_text)

if __name__ == '__main__':
    app.run(debug=True)


