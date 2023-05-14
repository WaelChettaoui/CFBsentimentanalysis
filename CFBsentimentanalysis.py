import pandas as pd
from textblob import TextBlob

# Read the Excel file into a Pandas DataFrame
data = pd.read_excel(r'C:\Users\waelc\OneDrive\Desktop\CFBsentimentanalysisproj\CS102_Input.xlsx')

data

# Create empty lists to store the sentiment scores
polarity_scores = []
subjectivity_scores = []

# Perform sentiment analysis on each row in the DataFrame
for index, row in data.iterrows():
    # Concatenate the text from multiple columns if needed
    text = ' '.join(str(row[column]) for column in data.columns)
    
    # Create a TextBlob object and calculate sentiment scores
    blob = TextBlob(text)
    polarity = blob.sentiment.polarity
    subjectivity = blob.sentiment.subjectivity
    
    # Append the scores to the respective lists
    polarity_scores.append(polarity)
    subjectivity_scores.append(subjectivity)

# Add sentiment scores as new columns in the DataFrame
data['Polarity'] = polarity_scores
data['Subjectivity'] = subjectivity_scores

# Save the updated DataFrame to a new Excel file
data.to_excel(r'C:\Users\waelc\OneDrive\Desktop\CFBsentimentanalysisproj\CS102_output.xlsx', index=False)