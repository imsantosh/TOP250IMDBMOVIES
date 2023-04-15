from bs4 import BeautifulSoup
import requests, openpyxl
import pandas as pd



excel = openpyxl.Workbook()
#print(excel)
#print(excel.sheetnames)

sheet = excel.active
sheet.title = 'Top Rated IMDB Movies'
#print(excel.sheetnames)

# Adding Headers into Excel sheet

sheet.append(['Movie Rank', 'Movie Name', 'Year Of Release', 'IMDB Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()
    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('tbody', class_='lister-list').find_all('tr')
    #print(len(movies))
    for movie in movies:
        movie_name = movie.find('td', class_ ='titleColumn').a.text
        rank = movie.find('td', class_ ='titleColumn').get_text(strip=True).split('.')[0] # To strip space new line character etc...
        year = movie.find('td', class_ ='titleColumn').span.text.strip('()')
        rating = movie.find('td', class_ ='ratingColumn imdbRating').strong.text
        #print(rank)
        #print(movie_name)
        #print(year)
        #print(rating)
        #print(f'{rank} {movie_name} {year} {rating}')
        #print(rank, movie_name, year, rating)
        sheet.append([rank, movie_name, year, rating])

except Exception as e:
    print(e)

excel.save('IMDB Movie Rating.xlsx')

# Reading excel file and printing saved values

movies_list = pd.read_excel('IMDB Movie Rating.xlsx', engine='openpyxl')
#print(movies_list.describe())

#print(movies_list.shape[0])
#print(movies_list.shape[1])

for row in range(250):
    for column in range(movies_list.shape[1]):
        #print(movies_list.loc['Movie Rank'])
        print(movies_list.iloc[row][column])
    if row == 1:
        break









