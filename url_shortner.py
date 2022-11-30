import pandas as pd

def read_excel(file):
    url_start, url_end, url_final= 'https://drive.google.com/uc?id=', '&export=download', ''
    sno_list, name_list, id_list, url_list = [], [], [], []
    pd.set_option('display.max_colwidth', None)
    urls_df = pd.read_excel(file)
    i = 0
    for _ in urls_df[['Name']].iterrows():
        sno_list.append(str(i + 1))
        i += 1
    for _ in urls_df[['Name']].iterrows():
        name_list.append(str(_[1])[8:-23].rstrip('\n'))
    for _ in urls_df[['Links']].iterrows():
        id_list.append(str(_[1])[41:-41].rstrip('\n'))
    for _ in id_list:
        url_final = url_start + _ + url_end
        url_list.append(url_final)
        url_final = ''
    with open('url_list.txt', 'w') as file:
        for i in range(len(name_list)):
            file.write(str(i + 1))
            file.write(". ")
            file.write(name_list[i])
            file.write(" - ")
            file.write(url_list[i])
            file.write("\n")
            
read_excel('books_urls.xlsx')