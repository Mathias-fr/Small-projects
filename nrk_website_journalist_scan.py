import bs4, requests, re, openpyxl

def nrk_website_journalist_scan(number_of_articles):
#Finds the authors of articles linked from the home page of the nrk website and displays
#them in excel
    
    wb = openpyxl.load_workbook("Workbook.xlsx")
    sheet = wb["Sheet1"]
    website_html = requests.get("https://www.nrk.no/")
    #website_html = open("nrk.no.html")
    bs4_object = bs4.BeautifulSoup(website_html.text, "html.parser")
    url_references = []

    for link in bs4_object.find_all('a',  attrs={'href':re.compile(
        "^https://www.nrk.no/(.*)\d.\d{8}$")}):
        
        url_references.append(link.get("href"))

    #number_of_articles = 3
    name_regex = re.compile("(\w+\s)+$")
    
    sheet.cell(row=1, column = 1).value = "Role"
    sheet.cell(row=1, column = 2).value = "Name"
    sheet.cell(row=1, column = 3).value = "Article number"
    row = 2
    column = 1
    for i in range(0, number_of_articles):
        
        article_html = requests.get(url_references[i])
        bs4_object = bs4.BeautifulSoup(article_html.text, "html.parser")
        author_names = bs4_object.select('a[class="author__name"]')
        author_roles = bs4_object.select('span[class="author__role"]')
        
        for j in range(0, len(author_names)):
            
            string_1 = author_names[j].getText()
            string_1 = name_regex.search(string_1)
            string_1 = string_1.group()
            string_1 = string_1.replace("\n", "")
            string_2 = author_roles[j].getText()
            
            sheet.cell(row=row, column=column+1).value = string_1
            sheet.cell(row=row, column=column).value = string_2
            sheet.cell(row=row, column=3).value = i+1
            #print(string_2)
            #print(string_1)
            row += 1

    wb.save("Workbook.xlsx")
