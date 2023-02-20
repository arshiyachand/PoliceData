import time
import pandas as pd
import numpy as np
import nltk
import io
import unicodedata
import re
import string
from numpy import linalg
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.tokenize import PunktSentenceTokenizer
from nltk.tokenize import PunktSentenceTokenizer
from nltk.corpus import webtext
from nltk.stem.porter import PorterStemmer
from nltk.stem.wordnet import WordNetLemmatizer
from nltk.tokenize import word_tokenize
import openpyxl
import pathlib

#List through all the files present in the folder, add the path where files are present

filesFromFolder = pathlib.Path("path")
listFiles = list(filesFromFolder.iterdir())
tempFiles = []

#Get the files that has comments

for file in listFiles:
	indexValue  = file.name.find('Comments')
	if indexValue > 0:
		tempFiles.append(file)
		print(file)

print(tempFiles)

#Extract the comments from the data dictionary extracted

#rowindex = 3
for j in tempFiles:	
	rowdata = pd.read_csv(j)
	comments_data = rowdata['comments']
	commentFiltered = []
	for comment in comments_data:
		if str(comment) != 'nan':
			data = re.findall(r'"message":(.*?)"}',comment)
			commentFiltered = commentFiltered + data
			
	print(commentFiltered)
	
	comment_textdata = ''
	for i in commentFiltered:
		comment_textdata = comment_textdata + " " + i

	#Clean up the comments by removing special characters

	comment_textdata = re.sub('[^A-Z a-z0-9]+','',comment_textdata)
	temp_commenttextdata = comment_textdata

	# Tokenizing the comments data which is spiliting the words 
 
	tokenize_sentence = PunktSentenceTokenizer(comment_textdata)
	sentense = tokenize_sentence.tokenize(comment_textdata)

	#nltk.download()

	#Normalize the comments text using porter stemmer and lematizer

	porter_stemmer = PorterStemmer()

	nltk_tokens = nltk.word_tokenize(comment_textdata)

	for w in nltk_tokens:
		print ("Actual: % s Stem: % s" % (w, porter_stemmer.stem(w)))
		

	wordnet_lemmatizer = WordNetLemmatizer()
	nltk_tokens = nltk.word_tokenize(comment_textdata)

	for w in nltk_tokens:
		print ("Actual: % s Lemma: % s" % (w, wordnet_lemmatizer.lemmatize(w)))
		
	comment_textdata = nltk.word_tokenize(comment_textdata)
	print(nltk.pos_tag(comment_textdata))

	#Running the data on Sentiment analyzer to get scores

	sid = SentimentIntensityAnalyzer()
	#nltk.download()
	nltk.download('all')

	comment_textdata = temp_commenttextdata
	
	#Get the positive, negative and neutral scores

	scores = sid.polarity_scores(comment_textdata)
	positiveresult = scores['pos']
	negativeresult = scores['neg']
	neutralresult = scores['neu']
	print(scores['pos'])

	departmentName = j.name.split("_",1)[0]

	# Load the Master Excel file, add the path to master file
	df = pd.read_excel("path")

	# Append the new row to the existing data
	new_row = {"Department": departmentName, "Positive Score": positiveresult, "Negative Score": negativeresult, "Neutral Score": neutralresult }
	df = df.append(new_row, ignore_index=True)

	# Save the data to the file, add the path to master file
	df.to_excel("path", index=False, engine='openpyxl')
	print("done")

	
				
	
