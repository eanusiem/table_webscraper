#Libraries needed to run the file
import os
import html5lib
import requests
import pandas as pd
from urllib.parse import urlparse,urlsplit


#URLs for the webpages 
urls=["https://en.wikipedia.org/wiki/List_of_films_considered_the_best"]

#Shows the number of URLs that were found
print(len(urls),"Urls Found")

#Changes the file name, removes _ and -, and puts the title case and removes spaces
def modify_name(my_str):
  new_title=my_str.replace("_", " ").replace("-", " ")
  return new_title.title().replace(" ","")

#Retrieves all the tables from the url 
def get_dataframes(url):
  html = requests.get(url).content
  df_list = pd.read_html(html)
  print(len(df_list),"Dataframes Returned")
  return df_list

#If the df is too small then it can be removed
def filter_dfs(dfs_list,min_rows=10):
  new_dfs_list=[]
  for each_df in dfs_list:
    if(len(each_df)>min_rows):
      new_dfs_list.append(each_df)
  return new_dfs_list

#The Excel worksheet name  must be less than 31 characters to avoid an error
def crop_name(name,thres=29):
  if len(name)<thres:
    return name
  else:
    return name[:thres]

#Retrieves the first n elements from list
def crop_list(lst,thres=29):
  if len(lst)<thres:
    return lst
  else:
    return lst[:thres]

#Converts urls to dataframes to excel sheets
#Get the maximum number of tables from each url
#Retrieves the minimum number of rows in each table to save it to the excel sheet
#Limites the number of excel sheets to prevent code from crashing

def urls_to_excel(urls,excel_path=None,get_max=10,min_rows=0,crop_name_thres=29):
  excel_path=os.path.join(os.getcwd(),"Michigan_Medicine_Reviews.xlsx") if excel_path==None else excel_path
  writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
  i=0
  for url in urls:
    parsed=urlsplit(url)
    sheet_name=parsed.path.split('/')[-1]
    mod_sheet_name=crop_name(modify_name(sheet_name),thres=crop_name_thres)

    dfs_list=get_dataframes(url)
    filtered_dfs_list=filter_dfs(dfs_list,min_rows=min_rows)
    filtered_dfs_list=crop_list(filtered_dfs_list,thres=get_max)
  for each_df in filtered_dfs_list:
    print("Parsing Excel Sheet "," : ",str(i)+mod_sheet_name)
    i+=1
    each_df.to_excel(writer, sheet_name=str(i)+mod_sheet_name, index=True)
  writer.save()
urls_to_excel(urls,get_max=1,min_rows=10)