import os
from imdb import IMDb
import PTN
import xlsxwriter


root = "g:\\Users\\manolis\\Downloads\\0001]_[video\\__latest\\"

# directories (movies)

dir_list = [item for item in os.listdir(root) if os.path.isdir(os.path.join(root,item)) and item[0] != '_']
print(dir_list)
print("Parsed HDD directory films --- Found "+str(len(dir_list))+" titles")

# create an instance of the IMDb class
ia = IMDb()

# TODO: - filename like: _contentsList_[date]
# - store the contents file in movie folder
# - logic to check for contents when running and ignore duplicate (using existing file)
# - handle series (different worksheet?)
# - include files (currently lists only folders)
# - introduce args (a. folder(s))
# - logging

workbook = xlsxwriter.Workbook('export.xlsx')
worksheet = workbook.add_worksheet()
# TODO: do this prettier:
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


        # TODO: cross check release year to find the correct movie
        # issues when a. there is no year in the torrent title, b. the year in the torrent title is wron
        imdb_info=ia.get_movie(movies_info[0].movieID)
        if 'year' in imdb_info and 'year' in info and imdb_info['year'] == info['year']:
            worksheet.write_url(row=row, col=col + 1, url=str(ia.get_imdbURL(imdb_info)),
                                string=str(imdb_info))
            worksheet.write(row, col + 2, str(imdb_info['year']))
        else:
            for movie in movies_info[1:]:
                imdb_info = ia.get_movie(movie.movieID)
                if 'year' in imdb_info and 'year' in info and imdb_info['year'] == info['year']:
                    worksheet.write_url(row=row, col=col + 1, url=str(ia.get_imdbURL(imdb_info)),
                                        string=str(imdb_info))
                    worksheet.write(row, col + 2, str(imdb_info['year']))
                    break


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
