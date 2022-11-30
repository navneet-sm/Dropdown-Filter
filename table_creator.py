from re import X
import pandas as pd

def read_excel_research(file):
    research_str = ""
    sno_list, np_list, fy_list, i_list, status_list = [], [], [], [], []
    pd.set_option('display.max_colwidth', None)
    research_df = pd.read_excel(file, sheet_name=0)  
    i = 0
    for _ in research_df[['Name of Project']].iterrows():
        sno_list.append(str(i + 1))
        i += 1
    for _ in research_df[['Name of Project']].iterrows():
        np_list.append(str(_[1])[19:-23].rstrip('\n'))
    for _ in research_df[['Financial Year']].iterrows():
        fy_list.append(str(_[1])[18:-23].rstrip('\n').replace('NaN', 'NA'))
    #for _ in research_df[['Funding Agency']].iterrows():
        #fa_list.append(str(_[1])[18:-23].rstrip('\n'))
    #for _ in research_df[['Budget (Rs. In Lakh)']].iterrows():
        #b_list.append(str(_)[28:-25].rstrip('\n').replace('NaN', 'NA'))
    for _ in research_df[['Institution / Agency Name']].iterrows():
        i_list.append(str(_[1])[29:-23].rstrip('\n').replace('NaN', 'NA'))
    #for _ in research_df[['Category']].iterrows():
        #c_list.append(str(_[1])[12:-23].rstrip('\n'))
    for _ in research_df[['Status']].iterrows():
        status_list.append(str(_[1])[10:-23].rstrip('\n'))
    #for _ in research_df[['Project Complete']].iterrows():
        #p_list.append(str(_[1])[19:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    research_data = map(list, zip(sno_list ,np_list, fy_list, i_list, status_list))
    for x in research_data:
        research_str += "<tr>\n" + "    <td class='sno'>" + x[0] + "</td>\n    <td class='name-td'>" + x[1] + "</td>\n    <td class='year-td'>" + x[2] + "</td>\n    <td class='institution-td'>" + x[3] + "</td>\n    <td class='status-td'>" + x[4] + "</td>\n" + "\n</tr>\n"
    with open('research_table.html', 'w') as file:
        file.write(research_str)
    print(i_list)    

def read_excel_publications(file):
    publications_str = ""
    sno_list, n_list, a_list, y_list, c_list, p_list, l_list  = [], [], [], [], [], [], []
    pd.set_option('display.max_colwidth', None)
    publications_df = pd.read_excel(file, sheet_name=1)
    i = 0
    for _ in publications_df[['Title of Book']].iterrows():
        sno_list.append(str(i + 1))
        i += 1
    for _ in publications_df[['Title of Book']].iterrows():
        n_list.append(str(_[1])[17:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Author']].iterrows():
        a_list.append(str(_[1])[10:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Year']].iterrows():
        y_list.append(str(_[1])[8:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Category']].iterrows():
        c_list.append(str(_[1])[12:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Price']].iterrows():
        p_list.append(str(_[1])[14:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Download Link']].iterrows():
        l_list.append(str(_[1])[17:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    publications_data = map(list, zip(sno_list ,n_list, a_list, y_list, c_list, p_list, l_list))
    for x in publications_data:
        publications_str += "<tr>\n" + "    <td>" + x[0] + "</td>\n    <td class='title-td'>" + x[1] + "</td>\n    <td class='author-td'>" + x[2] + "</td>\n    <td class='year-td'>" + x[3] + "</td>\n    <td class='category-td'>"+ x[4] + "</td>\n    <td class='price-td'>" + x[5] + "</td>\n    <td class='download-td'><a href='" + x[6] + "'>Download</a></td>\n" + "\n</tr>\n"
    with open('publications_table.html', 'w') as file:
        file.write(publications_str)    

def read_excel_bulletins(file):
    bulletins_str = ""
    s_list, t_list, v_list, m_list, y_list, p_list, l_list = [], [], [], [], [], [], []
    pd.set_option('display.max_colwidth', None)
    publications_df = pd.read_excel(file, sheet_name=2)
    i = 0
    for _ in publications_df[['Theme/Title of Bulletin']].iterrows():
        s_list.append(str(i + 1))
        i += 1
    for _ in publications_df[['Theme/Title of Bulletin']].iterrows():
        t_list.append(str(_[1])[27:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Volume']].iterrows():
        v_list.append(str(_[1])[10:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Month']].iterrows():
        m_list.append(str(_[1])[9:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Year']].iterrows():
        y_list.append(str(_[1])[8:-26].rstrip('\n').replace('NaN', 'NA').rstrip('\nN').rstrip('.'))
    for x in range(len(y_list)):
        if (y_list[x] == '0'):
            y_list[x] = 'NA'
    for _ in publications_df[['Price']].iterrows():
        p_list.append(str(_[1])[10:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    for _ in publications_df[['Download Link']].iterrows():
        l_list.append(str(_[1])[17:-23].rstrip('\n').replace('NaN', 'NA').rstrip('\nN'))
    bulletins_data = map(list, zip(s_list ,t_list, v_list, m_list, y_list, p_list, l_list))
    print(l_list)
    for x in bulletins_data:
        bulletins_str += "<tr>\n" + "    <td>" + x[0] + "</td>\n    <td class='title-td'>" + x[1] + "</td>\n    <td class='volume-td'>" + x[2] + "</td>\n    <td class='month-td'>" + x[3] + "</td>\n    <td class='year-td'>"+ x[4] + "</td>\n    <td class='price-td'>" + x[5] + "</td>\n    <td class='download-td'><a href='" + x[6] + "'>Download</a></td>\n" + "\n</tr>\n"
    with open('bulletins_table.html', 'w') as file:
        file.write(bulletins_str)  
    
read_excel_research('TRI_Research and Publications-Nov-30.xlsx')