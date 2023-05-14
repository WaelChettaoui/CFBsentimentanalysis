import pandas as pd
from textblob import TextBlob
import matplotlib.pyplot as plt
from io import BytesIO

# Read the Excel file into a Pandas DataFrame
data = pd.read_excel('CS102_Input.xlsx')

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

# Create a bar graph of polarity scores
plt.figure(figsize=(10, 6))
plt.bar(range(len(polarity_scores)), polarity_scores)
plt.xlabel('Row Index')
plt.ylabel('Polarity Score')
plt.title('Sentiment Analysis - Polarity Scores')
plt.xticks(range(len(polarity_scores)), range(len(polarity_scores)), rotation='vertical', ha='center')
plt.subplots_adjust(bottom=0.2)
plt.grid(True)

# Save the graph to a BytesIO object
graph_buffer = BytesIO()
plt.savefig(graph_buffer, format='png')
plt.close()

# Create an Excel writer object
writer = pd.ExcelWriter('output_file.xlsx', engine='xlsxwriter')

# Write the DataFrame to the Excel file on a new sheet
data.to_excel(writer, index=False, sheet_name='Sentiment Analysis')

# Create a new worksheet for the graph
workbook = writer.book
worksheet_graph = workbook.add_worksheet('Graph')

# Insert the graph image into the worksheet
worksheet_graph.insert_image('A1', 'graph.png', {'image_data': graph_buffer})

# Save the Excel file
writer.save()
