# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 18:25:40 2020

@author: Administrator

Let us see if I can code something to create an automation tool
that can automatically visualise data from an excel sheet

Made some updates to clean up the code
"""



#%% Importing libraries we need
import pandas as pd
import pyyed
import os
from datetime import datetime
import matplotlib
import sys

now = datetime.now()
dt_string = now.strftime("%Y.%m.%d - %H.%M")


print("Libraries imported and base variables defined.")
    
# Read folder for source file
# Make sure that when you're running this, it's referencing the correct folder
files = [f for f in os.listdir(os.curdir) if os.path.isfile(f) and f.endswith((".xls", ".xlsm", ".xlsx"))]
files.sort()
index_files = {i: files[i] for i in range(0, len(files))}

target = ''
target_ID = ''

if len(index_files) == 0:
    print("No files found in current folder, please double check what it's currently referring to.")
    sys.exit()    
if len(index_files) == 1:
    print("Only one file found in directory, assuming it to be target and importing directly.")
    target_ID = 0
    target = index_files[int(target_ID)]
    # Importing dataframe
    df = pd.read_excel(target, sheet_name = "Links")
    print(f"\n{target} imported.")
if len(index_files) > 1:
    for i in index_files:
        print(i, index_files[i])
    target_ID = input('Using the index number, which excel will we import from?: ')
    if target_ID.isalpha() or int(target_ID) > len(files):
        print('Please enter valid ID for import target.')
        target_ID = ''
        if KeyError or ValueError:
            print('Please enter valid ID for import target.')
            target_ID = ''
    target = index_files[int(target_ID)]
    
    # Importing dataframe
    df = pd.read_excel(target, sheet_name = "Links")
    print(f"\n{target} imported.")
    
# Defining the nature of the entities, refining those dfs to just relevant data
# Because pyyed is amazing and already can deal with these linkages
# Entity list generation

entity_list = df.iloc[:,0].values.tolist() + df.iloc[:,1].values.tolist()
entity_list = list(set(entity_list))

# Entity dict generation (Entity - Colour pairing). Filtering as well.

entity_colour = {k: g["Colour"].tolist() for k,g in df.groupby("Subsidiary")}

for v in entity_colour.values():
    v[:] = list(set(v))
    
entity_colour = {k: str(v) for k, v in entity_colour.items()}

for k in entity_list:
    if k not in entity_colour.keys():
        entity_colour.update({k: ""})

for k, v in entity_colour.items():
    entity_colour[k] = entity_colour[k].replace("[","")
    entity_colour[k] = entity_colour[k].replace("nan","")
    entity_colour[k] = entity_colour[k].replace("]","")
    entity_colour[k] = entity_colour[k].replace("'","")
    entity_colour[k] = entity_colour[k].replace(",","")
    entity_colour[k] = entity_colour[k].replace(" ","")
    
for k, v in entity_colour.items():
    if v == "":
        entity_colour[k] = 'white'

for k, v in entity_colour.items():
    try:
        entity_colour[k] = matplotlib.colors.cnames[entity_colour[k]]
    except:
        entity_colour[k] = '#FFFFFF' #default background to white

# Creating nodes
target_name = os.path.splitext(target)[0]
g = pyyed.Graph()

for name, colour in entity_colour.items():
    if name != 'Placeholder':
        g.add_node(name, label=name, label_alignment="center",
                   shape="rectangle", shape_fill = colour)
    else:
        g.add_node(name, label="", label_alignment="center",
                   shape="rectangle", shape_fill = "#FFFFFF",
                   transparent = "true", edge_color = "#FFFFFF")

g.add_node("Milestone", label = f"Made {dt_string} \n Based on: {target}",
           font_style = "bold")


# Sorting entities

style_df = df.iloc[:, [0,1,2]]
style_df = style_df.dropna(axis = 0) #any entities with special linkages? defining link style

label_df = df.iloc[:, [0,1,3]]
label_df = label_df.dropna(axis = 0) #defining link labels (e.g. % ownership)
label_df["Ownership"] = label_df["Ownership"].astype(str)

standard_link_df = df[df["Connection"].isna()]
standard_link_df = standard_link_df[standard_link_df["Ownership"].isna()]

# Creating links

for row in standard_link_df.itertuples():
        g.add_edge(row.Parent, row.Subsidiary,
                   color="#000000", arrowhead = "standard", arrowfoot = "none",
                   line_type = "line")

for row in label_df.itertuples():
    g.add_edge(row.Parent, row.Subsidiary, label = row.Ownership+'%',
                   color="#000000", arrowhead = "standard", arrowfoot = "none",
                   line_type = "line", )

for row in style_df.itertuples():
    if row.Connection == 1:
        g.add_edge(row.Parent, row.Subsidiary,
                   color="#000000", arrowhead = "standard",
                   arrowfoot = "standard", line_type = "dashed")
    elif row.Connection == 0:
        g.add_edge(row.Parent, row.Subsidiary,
                   color="#FFFFFF", arrowhead = "none",
                   arrowfoot = "none", line_type = "dotted")

g.add_edge("Placeholder", "Milestone",
           color="#FFFFFF", arrowhead = "none",
           arrowfoot = "none", line_type = "dotted")
#Exporting

g.write_graph(f'{target_name} Chart (ver. {dt_string}).graphml')
print(f"\nGraph exported as {target_name} Chart({dt_string}).graphml")