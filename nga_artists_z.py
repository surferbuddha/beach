import requests
from bs4 import BeautifulSoup
import mysql.connector

mydb = mysql.connector.connect(
  host="surferbuddha.mysql.pythonanywhere-services.com",
  user="surferbuddha",
  passwd="yourMom81",
  database="surferbuddha$default"
)



alphabet = []
for letter in range(65, 66):
    alphabet.append(chr(letter))
    #print(alphabet)

for currLetter in alphabet:

    page = requests.get('https://web.archive.org/web/20121007172955/https://www.nga.gov/collection/an'+currLetter+'1.htm')

    # Create a BeautifulSoup object
    soup = BeautifulSoup(page.text, 'html.parser')

    # Remove bottom links
    last_links = soup.find(class_='AlphaNav')
    last_links_arr = last_links.find_all('a')
    for linkers in last_links_arr:

        last_links_vals = linkers.get('href')
        print (last_links_vals.split("an",1)[1])
        print(last_links_vals)

    last_links.decompose()

    # Pull all text from the BodyText div
    artist_name_list = soup.find(class_='BodyText')

    # Pull text from all instances of <a> tag within BodyText div
    artist_name_list_items = artist_name_list.find_all('tr')

    #print(artist_name_list_items)
    for artist_name in artist_name_list_items:
        artist_link = artist_name.find('a')

        artist_content = artist_name.find_all('td')
        country = artist_content[1].text
        #print(artist_content[1])
        #links = 'https://web.archive.org' + artist_name.get('href')
        #links = 'https://web.archive.org' + artist_link.get('href')
        names = artist_link.contents[0]
        links = 'https://web.archive.org' + artist_link.get('href')
        #place = artist_name.content[1]
        #print(artist_name.prettify())
        #print(names)
        #print(place)
        #print(names)
        #print(link)
        #mycursor = mydb.cursor()
        #print(country)

        #sql = "INSERT INTO posts_posts (title, body, link) VALUES (%s, %s, %s)"
        #val = (str(names), str(country), str(links))

        #mycursor.execute(sql, val)

        #mydb.commit()

    #print(mycursor.rowcount, "record inserted.")

