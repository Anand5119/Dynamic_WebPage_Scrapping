import requests
import openpyxl
from datetime import datetime

time_now=datetime.now().strftime("(%d-%m-%Y @ %I.%M_%p)")
#SETTING UP EXCEL
excel = openpyxl.Workbook()
#workbook=openpyxl.load_workbook()
sheet = excel.active

# ADDING NAME TO THE EXCEL SHEET
sheet.title = 'Udemy WebScrap'# type: ignore
# ADDING REQUIRED DATA CLOLUMNS IN EXCEL
sheet.append(['S.No', 'Image', 'Title', 'Url', 'Rating', 'Total_Duration']) # type: ignore

#TRY BLOCK
try:
    page =int(input('Enter the no. of page number: '))
    x = range(page)
    for i in x:
        i += 1
        Url = f"https://www.udemy.com/api-2.0/search-courses/?p={i}&q=python&src=ukw&skip_price=true"
    # SETTING UP HEADER TO REQUEST DATA FORM WEB
        HEADERS = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36 Edg/115.0.1901.183',
            'Referer': f'https://www.udemy.com/courses/search/?p={i}&q=python&src=ukw'}
        response = requests.get(url=Url, headers=HEADERS).json()
        # print(response['courses'])
    # USING FOR LOOP TO SCRAP ALL DATA FROM WEB
        for course in response['courses']:
            # print(course)
    # USING DICTIONARY TO GET DATA FROM JSON FILE
            course_info = {

                'image': course['image_304x171'],
                'url': 'https://www.udemy.com'+course['url'],
                'rating': course['rating'],
                'title': course['title'],
                'total_time': course['content_info']
            }
    # CREATING OBJECTS TO APPEND DISCTIONARY VALUES TO EXCEL
            image = course_info.get('image')
            url = course_info.get('url')
            rating = course_info.get('rating')
            title = course_info.get('title')
            total_time = course_info.get('total_time')
    # CREATING AUTO-INCREMENTOR TO INCLUDE SERIAL NO. FOR EACH ROW OF DATA
            cur = sheet
            s_no = 0
            for i in cur: # type: ignore
                s_no += 1
    # APPENDING ALL THE DATA TO THE EXCEL FILE
            sheet.append([s_no, image, title, url, rating, total_time]) # type: ignore

# EXCEPTION BLOCK
except Exception as e:
    print('Link not available')

excel.save(f'Udemy WebScrap-{time_now}.xlsx')    
