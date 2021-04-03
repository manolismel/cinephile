import os
from imdb import IMDb, _logging
import PTN
import xlsxwriter
import argparse
import logging

cli_parser = argparse.ArgumentParser()
cli_parser.add_argument('--root', type=str, default='/Volumes/Untitled/Movies/',
                        help="Path of folder containing downloaded movies")
cli_parser.add_argument('--dir_format', type=bool, default=True,
                        help="Movies downloaded as folders with the mp4/mkv stored inside")
cli_parser.add_argument('--file_format', type=bool, default=False,
                        help="Movies downloaded directly as mp4/mkv files")
cli_parser.add_argument('--special_folders', type=str, default='_',
                        help="Any special folders structure, like movies stored per director folder, starting"
                             "with an underscore: e.g. _Aggelopoulos/")

logging.basicConfig(format='[%(asctime)s] %(levelname)s:: %(message)s', level=logging.INFO)
_logging.setLevel("error")  # only for imdbpy

args = cli_parser.parse_args()

# list with all movie titles as loaded from root
movies_list = []

if args.dir_format:
    # parses movies downloaded as folders
    dir_list = [item for item in os.listdir(args.root) if os.path.isdir(os.path.join(args.root, item))
                and item[0] != args.special_folders]

if args.dir_format:
    # parses movies downloaded directly as files (e.g. mp4)
    files_list = [f for f in os.listdir(args.root) if os.path.isfile(os.path.join(args.root, f))]

movies_list = dir_list + files_list
logging.info(f"Parsed {args.root} directory, found: {len(movies_list)} entries")

# create an instance of the IMDb class
ia = IMDb()

# TODO: - filename like: _contentsList_[date]
# TODO: - store the contents file in movie folder
# TODO: - logic to check for contents when running and ignore duplicate (using existing file)
# TODO: - handle series (different worksheet?)

workbook = xlsxwriter.Workbook('export.xlsx')
worksheet = workbook.add_worksheet()

column_dict = {"Title": 0, "ImDB": 1, "Year": 2, "Director": 3, "Genres": 4}

# Create header
for column_name, column_pos in column_dict.items():
    worksheet.write(0, column_pos, column_name)

row = 1

for entry in movies_list:
    info = PTN.parse(entry)

    if 'season' in info:
        continue

    if info['title'] == '':
        info['title'] = entry

    worksheet.write(row, column_dict["Title"], info['title'])

    movies_info = ia.search_movie(info['title'])
    if len(movies_info) != 0:
        logging.info(f"first result for {info['title']} :  -> {movies_info[0]} : {ia.get_imdbURL(movies_info[0])} ")

        # cross check release year to find the correct movie
        # issues when a. there is no year in the torrent title, b. the year in the torrent title is wrong
        # TODO: if no-match for all results: keep first
        imdb_info = ia.get_movie(movies_info[0].movieID)

        if ('year' in imdb_info and 'year' in info and imdb_info['year'] == info['year']) or 'year' not in info:
            logging.info(f"I decided to keep the first result :)")
            worksheet.write_url(row=row, col=column_dict["ImDB"], url=str(ia.get_imdbURL(imdb_info)),
                                string=str(imdb_info))

            if 'year' in imdb_info:
                worksheet.write(row, column_dict["Year"], str(imdb_info['year']))

        else:
            for movie in movies_info[1:]:
                imdb_info = ia.get_movie(movie.movieID)
                if 'year' in imdb_info and 'year' in info and imdb_info['year'] == info['year']:
                    logging.info(f"Matched year of release for: {imdb_info} -> {ia.get_imdbURL(imdb_info)}, ")
                    worksheet.write_url(row=row, col=column_dict["ImDB"], url=str(ia.get_imdbURL(imdb_info)),
                                        string=str(imdb_info))

                    worksheet.write(row, column_dict["Year"], str(imdb_info['year']))
                    break

        if 'directors' in imdb_info:
            # TODO: handle multiple directors & their hyperlinks
            #WIP
            person = imdb_info['directors'][0]
            worksheet.write_url(row=row, col=column_dict["Director"], url=ia.get_imdbURL(person),
                                string=person['name'])

        if 'genres' in imdb_info:
            worksheet.write(row, column_dict["Genres"], str(",".join(imdb_info['genres'])))

    else:
        logging.info(f"imdb returned no results for : {info['title']}")

    row += 1

workbook.close()
