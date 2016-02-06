import requests
import json

li=["Title","Plot","Year","imdbRating","tomatoRating","Genre","Director","Actors","Runtime"]

li_show=["Title","Plot","Year","imdbRating","tomatoMeter"]

def search(movie):
    url = "http://www.omdbapi.com/?"

    try:
        r = requests.post(url,params={"tomatoes":"true","t":movie})
    except:
        print("Network Error !")
        search(movie)
        return 0

    jformat =json.loads(r.text)

    return jformat

class MovieData(object):

    def __init__(self,jformat,li):
        self.jformat=jformat
        self.data = {}
        self.li=li
        for i in self.li:
            self.data.setdefault(i,"N/A")
        try:
            self.data = {
			"Title":self.jformat["Title"],
			"Plot":self.jformat["Plot"],
			"Year":self.jformat["Year"],
			"imdbRating":self.jformat["imdbRating"],
			"Genre":self.jformat["Genre"],
			"tomatoRating":self.jformat["tomatoRating"],
            "Director":self.jformat["Director"],
            "Actors":self.jformat["Actors"],
            "Runtime":self.jformat["Runtime"]
			}

        except:
            pass

    def getData(self):
        return self.data


def getInfo(movie_name):
    count = 0
    result = {}

    while count <= 3:
        movieObject = MovieData(search(movie_name),li)
        result = movieObject.getData()
        if result == 0:
            count = count + 1
        else:
            return result
    return result
