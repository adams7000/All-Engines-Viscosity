#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime


# In[2]:


def wrangle(filepath):
    df = pd.read_excel(filepath, sheet_name = "Viscosity")
    
    # drop ENGINE LUBE OIL DAILY VISCOSITY ANALYSIS REPORT FOR AKSA ENERGY GHANA - 2023 column
    df.drop(
            columns = "ENGINE LUBE OIL DAILY VISCOSITY ANALYSIS REPORT FOR AKSA ENERGY GHANA - 2023",
        inplace = True
    )
    
    # rename index "1" in column "Unnamed: 0" to DG_No
    df.at[1, "Unnamed: 0"] = "DG_No"  
   
    # drop irrelevant index namely "0, 2" and drop columns Unnamed: 1", "Unnamed: 3", "Unnamed: 4"
    df.drop(index = [0, 2], inplace = True)
    df.drop(columns = ["Unnamed: 1", "Unnamed: 3", "Unnamed: 4"], inplace = True)
    
    df = df.T.fillna(0)
        
    # Assign columns with "analysis_date" as new header
    new_header = df.iloc[0]
    df = pd.DataFrame(df.values[1:], columns = new_header)
    
    # set DG_No as index
    df.set_index("DG_No", inplace = True) 
    
    # replace "0" cells in last row with "1"
    df.iloc[-1] = df.iloc[-1].replace(0, 1)
    return df   


# In[3]:


df = wrangle(r"C:\Users\HP\Desktop\Engine Lube Oil Analysis Report - 2023.xlsx")
df.head()


# In[4]:


# select "last_but_one_non_zero" value of each row and its index and save to dictionary
def last_but_one_non_zero(series):
    non_zero_values = series[series != 0]
    if len(non_zero_values) >= 2:
        last_non_zero_index = non_zero_values.index[-2]
        last_non_zero_value = non_zero_values.iloc[-2]
        return last_non_zero_value, last_non_zero_index
    return None, None

prev_vis_values = {col: last_but_one_non_zero(df[col]) for col in df.columns}
prev_vis_values;


# In[5]:


# convert "prev_vis_values" dictionary to dataframe 
df1 = pd.DataFrame.from_dict(prev_vis_values, orient = "index")

# attribute it to the name "Previous Viscosity Value"
df1[["Previous Viscosity Value", "analysis_date"]] = df1

# drop columns "0, 1"
df1 = df1.drop(columns = [0, 1])

# change viscosity values to floating values 
pre_vis_data = df1["Previous Viscosity Value"].astype(float)

# change "datetime" dtype to string of format dd-mm-yy
analysis_date = df1["analysis_date"].dt.strftime("%d-%b-%y")

analysis_date.head()


# In[6]:


# set criteria for graph size 
figure = plt.figure(figsize = (7.5,6.8))

# set criteria for "upper limit" vertical line and plot in graph
upper_limit = 130*1.4
plt.axvline(upper_limit, linestyle = "--", color = "green", label = "Upper Limit")

# set criteria for "lower limit" vertical line and plot in graph
lower_limit = 130 - 130*0.25
plt.axvline(lower_limit, linestyle = "--", color = "red", label = "Lower Limit")

# plot viscosity horizontal bar graph
pre_vis_data.plot(kind = "barh", color = "b")

# indicate viscosity values inside barchart close to the top
for index, value in enumerate(pre_vis_data):
    plt.text(value, index, str(value),
             ha='right', va='center', color = "white", fontsize = 8, weight = "bold")
    
# use this code to indicate "analysis_date" inside barchart at the middle
for i, annotation in enumerate(analysis_date):
    plt.annotate(annotation, (pre_vis_data[i]/1.75, i),
    ha ='right', va='center', color = "white", weight = "bold", fontsize = 9)
    
# label axes, titles and position legend
plt.xlabel("Viscosity @ 40deg [cSt]")
plt.ylabel("Engine Number")
plt.title(
    "Previous Distribution of Lube Oil Viscosity of All Engines",
         weight = "bold"
        )
plt.legend(loc = "upper left", fontsize = "8");

plt.show()


# In[ ]:





# In[ ]:




