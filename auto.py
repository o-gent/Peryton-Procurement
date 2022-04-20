"""
Load rows from project costs and export as purchase forms
can be run easily in VScode notebooks for debugging, or just python auto.py

build with:
pyinstaller auto.py -F
clean up build folders afterwards plz
"""

print("""
   _____________   ________
  / ____/ ____/ | / /_  __/
 / / __/ __/ /  |/ / / /   
/ /_/ / /___/ /|  / / /    
\____/_____/_/ |_/ /_/     
                          systems
""")

# %%
print("loading modules")

import json
import os
import sys
import time
from typing import Dict

import pandas
from openpyxl import load_workbook

print("modules loaded")

# %% load in the json config file
try:
    with open("config.json", "r") as f:
        config: Dict[str, str] = json.load(f)
    PR_TEMPLATE = config["PR_TEMPLATE"]
    PROCUREMENT = config["PROCUREMENT"]
    NAME = config["Name"]
    NUMBER = config["Phone Number"]
    EMAIL = config["Email"]
    print("loaded config.json")
    print(PR_TEMPLATE, PROCUREMENT, NAME, NUMBER, EMAIL)
except Exception as e:
    print("config file not set up properly, see the 'config_base.json' file for examples")
    print("more details:")
    print(e)
    print("this will exit after 60 seconds")
    time.sleep(60)
    sys.exit()

time.sleep(2)

# %% load in project costs
try:
    pc = pandas.read_excel("0-1-b project costs.xlsx", "main", engine='openpyxl')
    print(f"found {len(pc)} rows")
except Exception as e:
    print(e)
    print("loading project costs failed, it is probably open")
    raise KeyboardInterrupt

# %% load in template
def load_template():
    try:
        template = load_workbook(filename=PR_TEMPLATE)
        template_sheet = template["PR form"]
    except Exception as e:
        print(e)
        print("loading PR template failed, it is probably open")
        raise KeyboardInterrupt
    return template_sheet, template

# %% find which forms need writing
already_done = os.listdir("./Purchase Forms")
done_order_nums = []
for name in already_done:
    try:
        done_order_nums.append(int(name.split("-")[1]))
    except:
        pass

# %% run through each order group and write the PR
for group in pc.groupby(["order group"]):
    if int(group[0]) in done_order_nums:
        print(f"skipping {group[0]} as it has already been written")
        continue
    
    template_sheet, template = load_template()
    
    index = 0
    rows = group[1]

    for row in rows.iterrows():
        
        row = row[1]
        row_number = str(18 + index)
        
        # set the Catelogue number
        template_sheet['B' + row_number] = row['link']
        # set the description
        template_sheet['C' + row_number] = row['description'] + f" ~ at {row['discount']*100}% discount"
        # set the Quantity
        template_sheet['H' + row_number] = row['quantity']
        # Set cost incl vat
        template_sheet['I' + row_number] = row['cost']*(1-row['discount'])
        # set cost exl vat
        template_sheet['J' + row_number] = row['cost excl vat']*(1-row['discount'])
        # set equipment/COSHH/Unit
        template_sheet['E' + row_number] = "No"
        template_sheet['F' + row_number] = "No"
        template_sheet['G' + row_number] = "box"

        index += 1

    template_sheet['I26'] = rows['shipping'].sum()
    
    # set company
    template_sheet['C5'] = row['company']

    # Requester details
    template_sheet['F10'] = row['date sent']
    template_sheet['F5'] = NAME
    template_sheet['F8'] = NUMBER
    template_sheet['F9'] = EMAIL

    template.save(f"Purchase Forms/PR - {int(row['order group'])} - {row['company']}.xlsx")
    print(f"written PR - {int(row['order group'])} - {row['company']}")


print("close when ready")
time.sleep(10)

# %%
