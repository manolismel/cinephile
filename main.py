import os
from imdb import IMDb, Person
import PTN
import xlsxwriter


root = "g:\\Users\\manolis\\Downloads\\0001]_[video\\__latest\\"

# directories (movies)
dir_list = [item for item in os.listdir(root) if os.path.isdir(os.path.join(root,item)) and item[0] != '_']
print("Parsed HDD directory films --- Found "+str(len(dir_list))+" titles")

# create an instance of the IMDb class
ia = IMDb()

workbook = xlsxwriter.Workbook('export.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Title")
worksheet.write(0, 1, "ImDB")
worksheet.write(0, 2, "Year")
worksheet.write(0, 3, "Director")
worksheet.write(0, 4, "Genres")

row = 1
col = 0

for entry in dir_list:
    info = PTN.parse(entry)

    if 'season' in info:
        continue

    if info['title'] == '':
        info['title'] = entry

    print(info['title'])

    worksheet.write(row, col, info['title'])

    movies_info = ia.search_movie(info['title'])
    if len(movies_info) != 0:
        print(movies_info[0], movies_info[0].movieID, ia.get_imdbURL(movies_info[0]))
        worksheet.write_url(row=row,col=col + 1, url=str(ia.get_imdbURL(movies_info[0])), string=str(movies_info[0]))

        # TODO: cross check release year to find the correct movie (I just use the first result for now with BOGUS results)
        imdb_info=ia.get_movie(movies_info[0].movieID)
        if 'year' in imdb_info:
            worksheet.write(row, col+2, str(imdb_info['year']))
        if 'directors' in imdb_info:
            # TODO: handle multiple directors & their hyperlinks
            person = imdb_info['directors'][0]
            worksheet.write_url(row=row, col=col + 3, url=ia.get_imdbURL(person),
                                string=person['name'])

        if 'genres' in imdb_info:
            worksheet.write(row, col+4, str(",".join(imdb_info['genres'])))


    else:
        print("ImDB returned no results")

    row +=1

workbook.close()
