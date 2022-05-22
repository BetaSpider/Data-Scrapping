from bs4 import BeautifulSoup
import requests
import openpyxl
spreadsheet=openpyxl.Workbook()
sheet=spreadsheet.active
sheet.title="Top50 SCI-FI Movies"
#print(sno,title,year,certificate,runtime,genre,rating,director,votes,Gross)
sheet.append(["S.no","Movie_name","Year","Certificate","Time","Genre","Movie_rating","Director_name","No of votes","Gross amount"])

try:

    url=requests.get("https://www.imdb.com/search/title/?genres=sci_fi&sort=user_rating,desc&title_type=feature&num_votes=25000,&pf_rd_m=A2FGELUUNOQJNL&pf_rd_p=5aab685f-35eb-40f3-95f7-c53f09d542c3&pf_rd_r=B8Q5YPKS54TX66WH5W63&pf_rd_s=right-6&pf_rd_t=15506&pf_rd_i=top&ref_=chttp_gnr_17")
    result=BeautifulSoup(url.text,"html.parser")
    #print(result)
    movies=result.find('div',class_="lister-list").find_all('div',class_="lister-item")
    for scifi in movies:
        print(scifi)
        sno=scifi.find('h3').find('span',class_="lister-item-index").get_text(strip=True).split(".")[0]
        title=scifi.find('h3').a.text
        year=scifi.find('h3').find('span',class_="lister-item-year").get_text(strip=True).replace("(","")
        year=year.replace(")","")
        certificate=scifi.find('p').span.text
        runtime=scifi.find('p').find('span',class_="runtime").get_text(strip=True)
        genre=scifi.find("p").find("span",class_="genre").get_text(strip=True)
        rating=scifi.find("div",class_="inline-block ratings-imdb-rating").strong.text
        director=scifi.find("p",class_="").a.text
        votes=scifi.find("p",class_="sort-num_votes-visible").find_all("span")[1].get_text()
        Gross=scifi.find("p",class_="sort-num_votes-visible").find_all("span")[-1].get_text()

        #print(sno,title,year,certificate,runtime,genre,rating,director,votes,Gross)
        sheet.append([sno,title,year,certificate,runtime,genre,rating,director,votes,Gross])
except Exception as e:
    print("Sorry")

spreadsheet.save("TOP 50 sci-fi movies IMDb.xlsx")


