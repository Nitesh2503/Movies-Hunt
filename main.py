from bs4 import BeautifulSoup
import requests, openpyxl
import pandas as pd
#generatiing excel file
# excel = openpyxl.Workbook()
# print(excel.sheetnames)
# sheet=excel.active
# sheet.append(['Movie Rank','Movie Name','Year of Release','About'])
movie_list=[]

try:
    source=requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()   #To show error if the website doesn't exist
    
    soup=BeautifulSoup(source.text,'html.parser')    #parsing HTML
 
    #Finding All Generes
    genre_list=soup.find('ul',class_='quicklinks').find_all('li',class_='subnav_item_main')   
    print("Total Generes available are: ",len(genre_list))
    
    print("_______GENRE LIST______ ")
    for genre in genre_list:
         item=genre.a.get_text(strip=True)
         print(item)
         
    print("\n")
    
    link=""   #Variable for newly searched Genere
    
    type=input("Enter Your Genre: ")
     
    for genre in genre_list:
        
        val=genre.a.get_text(strip=True)
        
        if(val.lower()==type.lower()):             #searching the user Genere in the Genere list
           
            
            # sheet.title='Top '+type+' Rated Movies'
            # print(excel.sheetnames)
            print("\n**_________________ [GENERE FOUND ] _________________**\n")
            
            link=genre.find('a').get('href').split('&')[0]
            link="https://imdb.com/"+link                  #Getting link in the desired form
            print("\n[Link Generated]")
            print(link)
            
            
            source=requests.get(link)
            source.raise_for_status()   #To show error if the website doesn't exist 
            soup=BeautifulSoup(source.text,'html.parser')
            
            movies=soup.find('div',class_="lister-list").find_all('div',class_='lister-item mode-advanced')
            
            print("\nTotal Movies of type [ "+type.upper()+" ] are: ",len(movies),"\n")
            lim=int(input("Enter Number of Rows to be printed [less than "+ str(len(movies))+" ]: "))
            
            i=0;
            
            for movie in movies:
              if i<lim:  
                rank=movie.find('h3').find('span','lister-item-index unbold text-primary').text.strip('.')
                name=movie.find('h3').find('a').text
                date=movie.find('h3').find('span','lister-item-year text-muted unbold').text.strip('()')
                # rating=movie.find('div',class_='ratings-bar').find('div','inline-block ratings-imdb-rating').find('strong').text
                # runtime=movie.find('p',class_='text-muted').find('span',class_='runtime').text
                about=movie.find_all('p',class_='text-muted')[-1].text
                
                print(rank,name,date,about,"\n")
                # sheet.append([rank,name,date,about])
                
                data={
                    'Movie Rank':rank,
                    'Movie Name':name,
                    'Year of Release':date,
                    'About':about
                }
                movie_list.append(data)
                
                i=i+1
                path = r'D:\data'
                df=pd.DataFrame(movie_list)
                df.to_csv(path+'IMDB_Movie_Info.csv')
                
              else:
                  break  #break if the limit of roms exceeded
              
            break    #break if the Genere isn't found
        
    else:
        print("!!_________________ [GENERE NOT FOUND ]_________________ !!")                   
        exit

    
except Exception as e:
    print("!!............ERROR...........!!")
    print(e) 
    
# excel.save('IMDB Movie Info.xlsx')      
