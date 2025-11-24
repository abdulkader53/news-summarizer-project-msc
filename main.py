import tkinter as tk
import nltk
import os
import pandas as pd
from datetime import datetime
from newspaper import Article
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

nltk.download('punkt')

# Excel file path
EXCEL_FILE = "summaries.xlsx"

# Function to create excel file
def create_excel_file():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=["Timestamp", "Title", "Summary", "Polarity", "Sentiment", "URL"])
        df.to_excel(EXCEL_FILE, index=False)

# Function to log details into that file
def log_summary_to_excel(title_text, summary_text, polarity, sentiment_result, url):
    create_excel_file()  # Ensure file exists

    # Read existing data or create a new DataFrame
    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception:
        df = pd.DataFrame(columns=["Timestamp", "Title", "Summary", "Polarity", "Sentiment", "URL"])

    # Add new summary with timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_entry = pd.DataFrame([[timestamp, title_text, summary_text, polarity, sentiment_result, url]], 
                              columns=df.columns)
    df = pd.concat([df, new_entry], ignore_index=True)

    # Save back to Excel
    df.to_excel(EXCEL_FILE, index=False)

# Function to summarize the article
def summarize():
    url = utext.get('1.0', "end-1c").strip()
    if not url:
        return  # Prevent errors if URL field is empty

    article = Article(url)
    
    try:
        article.download()
        article.parse()
        article.nlp()
    except Exception as e:
        summary.config(state='normal')
        summary.delete('1.0', 'end')
        summary.insert('1.0', f"Error: {e}")
        summary.config(state='disabled')
        return

    # Enable text fields for inserting new values
    title.config(state='normal')
    summary.config(state='normal')
    sentiment.config(state='normal')

    # Clear previous text
    title.delete('1.0', 'end')
    summary.delete('1.0', 'end')
    sentiment.delete('1.0', 'end')

    # Extract values
    title_text = article.title if article.title else "N/A"
    summary_text = article.summary if article.summary else "N/A"

    # Sentiment analysis using VADER
    analyzer = SentimentIntensityAnalyzer()
    sentiment_scores = analyzer.polarity_scores(article.text)
    compound_score = sentiment_scores['compound']

    if compound_score >= 0.05:
        sentiment_result = "Positive"
    elif compound_score <= -0.05:
        sentiment_result = "Negative"
    else:
        sentiment_result = "Neutral"

    # Insert into text fields
    title.insert('1.0', title_text)
    summary.insert('1.0', summary_text)
    sentiment.insert('1.0', f'Polarity: {compound_score:.2f}, Sentiment: {sentiment_result}')

    # Log to Excel
    log_summary_to_excel(title_text, summary_text, compound_score, sentiment_result, url)

    # Disable fields after inserting text
    title.config(state='disabled')
    summary.config(state='disabled')
    sentiment.config(state='disabled')

# Function to open the Excel file
def open_excel():
    create_excel_file()  # Ensure file exists
    os.system(f'start excel "{EXCEL_FILE}"' if os.name == "nt" else f'libreoffice "{EXCEL_FILE}"')

# GUI Setup
root = tk.Tk()
root.title("News Summarizer")
root.geometry('1200x550')

tk.Label(root, text='Title').pack()
title = tk.Text(root, height=1, width=140, bg='#dddddd', state='disabled')
title.pack()

tk.Label(root, text='Summary').pack()
summary = tk.Text(root, height=20, width=140, bg='#dddddd', state='disabled')
summary.pack()

tk.Label(root, text='Sentiment Analysis').pack()
sentiment = tk.Text(root, height=1, width=140, bg='#dddddd', state='disabled')
sentiment.pack()

tk.Label(root, text='URL').pack()
utext = tk.Text(root, height=1, width=140)
utext.pack()

tk.Button(root, text="Summarize", command=summarize).pack()
tk.Button(root, text="Open Excel", command=open_excel).pack()

root.mainloop()
