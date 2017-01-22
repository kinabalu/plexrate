from plexapi.server import PlexServer
from plexapi.myplex import MyPlexAccount
import tmdbsimple as tmdb
import time
import xlsxwriter

try:
    import config
except ImportError:
    print("Config file config.py.tmpl needs to be copied over to config.py")
    sys.exit(1)

tmdb.API_KEY = config.THE_MOVIEDB_API_KEY

account = MyPlexAccount.signin(config.PLEX_USERNAME, config.PLEX_PASSWORD)
plex = account.resource(config.PLEX_HOST).connect()

workbook = xlsxwriter.Workbook(config.EXCEL_FILENAME)
worksheet = workbook.add_worksheet()

movies = plex.library.section(title=config.PLEX_LIBRARY_SECTION)

row = 0
col = 0
for entry in movies.all():
    search = tmdb.Search()
    response = search.movie(query=entry.title)

    vote_average = -1
    if len(response['results']) > 0:
        vote_average = response['results'][0]['vote_average']

    print("%s - %s" % (entry.title, vote_average))
    worksheet.write(row, col, entry.title)
    worksheet.write(row, col + 1, vote_average)
    row = row + 1
    # time.sleep(2)

workbook.close()
