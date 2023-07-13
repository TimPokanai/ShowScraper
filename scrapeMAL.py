import requests, openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Anime"

sheet.append(["Rank", "Name", "Release Date", "Rating"])

try:
    # accessing website and retrieving response object
    source = requests.get("https://myanimelist.net/topanime.php")

    # parsing using default Python installation parser
    parsedSoup = BeautifulSoup(source.text, "html.parser")
    
    # finding every entry in the top 50 list
    anime = parsedSoup.find("table", class_="top-ranking-table").findAll("tr", class_="ranking-list")
    
    for media in anime:

        mediaName = media.find("h3", class_="hoverinfo_trigger").a.text
        
        mediaRank = media.find("td", class_="rank ac").span.text
        
        mediaRelease = media.find("div", class_="information di-ib mt4").text.split("\n")[2].strip(" ")

        mediaRating = media.find("td", class_="score ac fs14").span.text

        print(mediaName, mediaRank, mediaRelease, mediaRating)

        # appending scraped data onto csv file
        sheet.append([mediaRank, mediaName, mediaRelease, mediaRating])

except Exception as e:
    print(e)

excel.save("MAL Scraping.xlsx")