# -*- coding: utf-8 -*-
import requests,openpyxl
from bs4 import BeautifulSoup
import lxml
import pandas as pd
import pymongo
import json
excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title='Top Rated Movies'
print(excel.sheetnames)
sheet.append(["Movie Rank","Movie Name","Year of Release","IMBD Rating"])
try:
  source = requests.get('https://www.imdb.com/chart/top/')
  source.raise_for_status()
  soup = BeautifulSoup(source.text, 'lxml')
  movies = soup.find('tbody',class_="lister-list").find_all('tr')
  print(len(movies))
  
  for movie in movies:
    name=movie.find('td',class_="titleColumn").a.text
    rank=movie.find('td',class_="titleColumn").get_text (strip=True).split('.')[0]
    year=movie.find('td',class_="titleColumn").span.text.strip('()')
    rating=movie.find('td',class_="ratingColumn imdbRating").strong.text   
    print(rank,name,year,rating)
    sheet.append([rank,name,year,rating])
except Exception as e:
  print(e)
excel.save("IMBD Movie Rating.csv")
df = pd.DataFrame(sheet.values)
df.columns=["Movie_Rank","Movie_Name","Year_of_Release","IMBD_Rating"]
df=df[df['Movie_Rank']!='Movie Rank']
#establish connection with mongodb
client = pymongo.MongoClient("mongodb://localhost:27017")
data1 = df.to_dict(orient="records")
db = client["topratedmovies"]
db.imbd.insert_many(data1)