# -- IMPORTS --
import pandas as pd
import numpy as np

import tkinter as tk
import tkinter.font as font
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from datetime import datetime
import time

# -- script to improve clarity of text in window (for windows 10 only!) --
try:
    from ctypes import windll

    windll.schore.SetProcessDpiAwareness(1)
except:
    pass

# -- global variables --
ussaar_file = ""
wards_file = ""
polk_file_prev = ""
polk_file_curr = ""
output_folder = ""


# -- FUNCTIONS --

# -- function that runs the background code that cleans, sorts, and compiles the data --
def run():
    # -- read in all of the files uploaded by the user --
    global month
    ussaar = pd.read_excel(ussaar_file, header=7)
    fleets_prev = pd.read_excel(polk_file_prev, header=5)
    fleets_curr = pd.read_excel(polk_file_curr, header=5)
    incentives = pd.read_excel(motor_incentives, sheet_name=0, header=3)
    sales = pd.read_excel(wards_sales_file, header=4)
    production = pd.read_excel(wards_production_file, header=3)

    # -- USSAAR DATA CLEANING --
    # -- Define the current date which we will use later to index the original dataset properly --
    now = datetime.now()
    # -- Reset the index (row names) which gives us the ability to rename them to the months of the year later --
    ussaar.reset_index(inplace=True)
    # -- Selecting the only 3 columns we need from the dataset --
    ussaar = ussaar[["level_0", "level_1", "Lt. Veh..2"]]
    # -- From the current date that we saved in a variable earlier, we just want to know the current year which is saved in a variable called year.
    # From there we check for only the items in the 'level_1' column that are equal to the current year and get the subsequent data from column' Lt. Veh..2' --
    year = now.year
    monthly_data_current_year = ussaar[ussaar["level_1"] == year]["Lt. Veh..2"]
    # -- Now we can do a similar step for getting the information for the previous year and we need to create our new index which will be in numerical order
    # matching the exact number of months in a year so we can rename them later. --
    previous_year = year - 1
    monthly_data_past_year = ussaar[ussaar["level_1"] == previous_year]["Lt. Veh..2"]
    # -- We create a new set of data by combining all of the data points we pulled together above. --
    data = pd.concat([monthly_data_past_year, monthly_data_current_year], axis=1)
    # -- Now we take the combined data and create a new dataframe called ussaar_final. We then rename the indeces based on the which number correlates to which month
    # Then we name the remaining columns based on which year the data represents. --
    ussaar_final = pd.DataFrame(data)
    ussaar_final.rename(
        index={
            0: "Jan",
            1: "Feb",
            2: "Mar",
            3: "Apr",
            4: "May",
            5: "Jun",
            6: "Jul",
            7: "Aug",
            8: "Sep",
            9: "Oct",
            10: "Nov",
            11: "Dec",
        },
        inplace=True,
    )
    ussaar_final.columns = [previous_year, year]

    # -- POLK FLEET DATA CLEANING --
    # -- Remove last four rows of dataset because they contain irrelevant information. --
    fleets_prev = fleets_prev.iloc[:-4]
    fleets_curr = fleets_curr.iloc[:-4]
    # -- Removing 'Lotus', 'Isuzu' and 'Polaris' indeces from the dataset. --
    fleets_prev = fleets_prev[
        (fleets_prev["Corporation"] != "LOTUS")
        & (fleets_prev["Corporation"] != "ISUZU COMMERCIAL TRUCK")
        & (fleets_prev["Corporation"] != "POLARIS")
    ]

    fleets_curr = fleets_curr[
        (fleets_curr["Corporation"] != "LOTUS")
        & (fleets_curr["Corporation"] != "ISUZU COMMERCIAL TRUCK")
        & (fleets_curr["Corporation"] != "POLARIS")
    ]
    # -- Since we don't need any data from the 'Governemtn' category we can filter out rows that contain that data in the 'New Category'
    # column and then only keep the 4 columns we need to move forward with in our dataset. --
    fleets_prev = fleets_prev[fleets_prev["New Category"] != "GOVERNMENT"][
        ["Corporation", "New Category", "Body Style", "CYTD"]
    ]
    fleets_curr = fleets_curr[fleets_curr["New Category"] != "GOVERNMENT"][
        ["Corporation", "New Category", "Body Style", "CYTD"]
    ]
    # -- Now we can define lists that contain the categorization labels for groups of brands. Once the lists are defined, we define a
    # function that references the lists and returns the updated category for a row. --
    other_asian_fleet = ["MAZDA", "MITSUBISHI", "SUBARU", "VOLVO"]
    euro_fleet = ["BMW", "DAIMLER", "PORSCHE", "MCLAREN", "MERCEDES BENZ", "TATA"]
    ev_fleet = ["TESLA MOTORS", "KARMA"]
    ford_fleet = ["ASTON MARTIN", "FORD"]
    fca_fleet = ["FCA", "ALFA ROMEO"]
    # -- function that takes 'corporation' column and changes names based on key and lists above --
    def corporation_group(x):
        if x in other_asian_fleet:
            return "Other Asian"
        elif x in euro_fleet:
            return "Euro"
        elif x in ev_fleet:
            return "Tesla + EV OEMs"
        elif x in ford_fleet:
            return "Ford"
        elif x in fca_fleet:
            return "FCA"
        elif x == "TOYOTA":
            return "Toyota"
        elif x == "VOLKSWAGEN":
            return "VW"
        elif x == "HYUNDAI":
            return "Hyundai"
        elif x == "NISSAN":
            return "Nissan"
        elif x == "HONDA":
            return "Honda"
        elif x == "VOLVO CORP":
            return "Volvo"
        elif x == "GENERAL MOTORS":
            return "GM"
        else:
            return x

    # -- Define function that changes the name of 'body style' to a standardized name which will be used in the Power BI. --
    def body_type(x):
        if x == "Passenger Vans":
            return "Mini-Van"
        elif x == "Pickups":
            return "Pick-Up"
        elif x == "Sport Utility":
            return "SUV/CUV"
        elif x == "Station Wagon":
            return "Other Car"
        elif x == "Truck":
            return "Large Van"
        elif x == "Truck Wagon":
            return "SUV/CUV"
        elif x == "Van":
            return "Large Van"
        else:
            return x

    # -- Using the corporation_group function we defined above, we can change all of the names in the 'Corporation' column to reflect the new
    # categories defined by the lists we defined above. --
    fleets_prev["Corporation"] = fleets_prev["Corporation"].apply(corporation_group)
    fleets_curr["Corporation"] = fleets_curr["Corporation"].apply(corporation_group)
    # -- Now we can use the body_type function we defined above to change the values in the 'Body Style' column to reflect the standardized
    # names which will be used in the Power BI. --
    fleets_prev["Body Style"] = fleets_prev["Body Style"].apply(body_type)
    fleets_curr["Body Style"] = fleets_curr["Body Style"].apply(body_type)

    # -- MOTORTREND INCENTIVES DATA CLEANING --
    # -- Remove last 6 rows because they contain irrelevant information. --
    incentives = incentives[:-6]
    # -- Reset dataset index in order to set it up for manipulation in later step. --
    incentives.reset_index(inplace=True)
    # -- Remove all of the rows with NaN values. --
    incentives = incentives[pd.notnull(incentives["index"])]
    # -- Define function that will take in string and remove `*` from final value. --
    def parser_star(item):
        string = str(item)
        output = string.split("*")[0]
        return output

    # -- Applied function we just defined to the first column ('index') in order to return just the portion of the string we need. --
    incentives["index"] = incentives["index"].apply(parser_star)
    # -- The original dataset contained several 'summary' lines for each brand so this part of the script allows us to only return
    # the brand names and not all of the unncessary summary lines. This way, in the future the brands can change when we pull in new
    # datasets but as long as the summary lines don't change the script will always just return the brand names. --
    incentives = incentives[
        (incentives["index"] != "      Total Cars ")
        & (incentives["index"] != "      Industry Truck ")
        & (incentives["index"] != "      Industry Car ")
        & (incentives["index"] != "      Total Light Trucks ")
    ]
    # -- We only need the first couple of columns in the dataset so this pulls those columns and removes the rest. --
    incentives = incentives.iloc[:, :3]
    # -- One of the 3 columns is unncessary so we will drop the 'Date, Year' column and just keep the 2 columns we care about. --
    incentives.drop(["Sep 19"], axis=1, inplace=True)

    # -- WARD'S SALES DATA CLEANING
    # -- All of these variables need to be defined for Ward's Sales and Production data in order to make sure we pull the correct
    # columns given the user indicated month. The `num_month` dictionary is only needed for naming the 'sum' columns in the output excel file. --
    # -- Define several variables to be used in automating which month of report should be evaluated --
    num_month = {
        1: "Jan",
        2: "Feb",
        3: "Mar",
        4: "Apr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Aug",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dec",
    }
    now = datetime.now()
    current_year = now.year
    previous_year = current_year - 1
    month = int(
        month_result.get()
    )  # *******************USER NEEDS TO CHANGE NUMBER OF MONTH MANUALLY **********************
    user_month = num_month.get(month)
    previous_year_index_sales = 7 + month
    previous_year_index_production = 6 + month
    # -- Removing the last four rows because they contain irrelevant data --
    sales = sales[:-4]
    # -- Removing 'Source' and 'Country' columns from dataset because they aren't necessary --
    sales.drop(["Source", "Country"], axis=1, inplace=True)
    # -- The initial Ward's dataset contains column names that are on two lines. When we read the data into python we have to cut off one
    # of those lines. For this reason some columns have similar names so python automatically names them `column`, `column.1`, `column.2`, etc.
    # The code below is taking one of the column names where this happened and renames it. --
    sales.rename(columns={"Group.1": "Segment Group"}, inplace=True)
    # -- The Ward's dataset contains data for Jan to Dec of the previous year but only Jan to current month of current year. For this reason
    # we need to select only the columns that are needed for the Power BI. The `labels` variable grabs the identifiers, the `prev_data` grabs
    # the previous year's data but only up to the month indicated by the user, and the `curr_data` takes all of the columns that are there for
    # the current year. Then we join everything together to make our new dataset. --
    labels = sales.iloc[:, 0:5]
    prev_data = sales.iloc[:, 5:previous_year_index_sales]
    curr_data = sales.iloc[:, 19:]
    sales = labels.join(prev_data)
    sales = sales.join(curr_data)
    # -- Some of the columns in the dataset have the data wrapped in parentheses. This function will allow us to take one of those elements as an
    # input and output the data without any parentheses. --
    # -- Function that parses columns that have extra parentheses and returns second half of parsed string --
    def parser_parathenses(x):
        name = x.split(")")[1]
        return name[1:]

    # -- Since our Power BI will have groupings, such as 'Euro' and 'Other Asian', we need to define lists that contain all of the brands that fall
    # under a certain category. Once we have our lists, we can define a function that will take a data point as an input and output either a new category
    # or the original depending on if the brand name exists in our defined lists. --
    # -- Define lists that contain all of the companies that will be grouped into other_asian and euro categories --
    other_asian = [
        "Mazda",
        "Mitsubishi",
        "Subaru",
        "Suzuki",
        "Volvo",
        "Beijing AIC",
        "Zhejiang Geely",
        "Isuzu",
        "Tata Motors",
    ]
    euro = [
        "Audi",
        "BMW",
        "Daimler",
        "Jaguar Land Rover",
        "Peugeot Citroen",
        "Porsche",
        "Volkswagen",
    ]
    # -- Function that takes "Group" column and references defined lists to change names of companies to final graph versions --
    def categories_1(x):
        if x in other_asian:
            return "Other Asian"
        elif x in euro:
            return "Euro"
        elif x == "Hyundai Group" or x == "Kia Motors":
            return "Hyundai/Kia"
        elif x == "Tesla Motors":
            return "Tesla"
        elif x == "Fiat Chrysler":
            return "FCA"
        elif x == "General Motors":
            return "GM"
        elif x == "Renault":
            return "RN"
        else:
            return x

    # -- Similar to the function we defined above except the next four cells define lists and functions that allow us to group the 'segment'
    # column into previously defined categories for Power BI. --
    # -- Define lists for Cross Utility as luxory or non-luxory. Use lists to define function that changes categories in
    # "Segment" column to be "luxury" or "non-luxory" --
    luxury = ["Middle Luxury CUV", "Small Luxury CUV", "Large Luxury CUV"]
    not_luxury = ["Middle CUV", "Small CUV", "Large CUV"]

    def cross_utility_luxury(x):
        if x in luxury:
            return "Luxury"
        elif x in not_luxury:
            return "Non-Luxury"
        else:
            return " "

    # -- Define lists for Cross Utility as small, middle, large. Use lists to define function that changes categories in
    # "Segment" column to be "small", "middle" or "large" --
    large = ["Large Luxury CUV", "Large CUV"]
    middle = ["Middle Luxury CUV", "Middle CUV"]
    small = ["Small Luxury CUV", "Small CUV"]

    def cross_utility_lms(x):
        if x in large:
            return "Large CUVs"
        elif x in middle:
            return "Middle CUVs"
        elif x in small:
            return "Small CUVs"
        else:
            return " "

    # -- Define lists for Sport Utility as luxory or non-luxory. Use lists to define function that changes categories in
    # "Segment" column to be "luxury" or "non-luxory" --
    luxury = ["Middle Luxury SUV", "Large Luxury SUV", "Small Luxury SUV"]
    not_luxury = ["Middle SUV", "Small SUV", "Large SUV"]

    def suv_luxury(x):
        if x in luxury:
            return "Luxury"
        elif x in not_luxury:
            return "Non-Luxury"
        else:
            return " "

    # -- Define lists for Sport Utility as small, middle, large. Use lists to define function that changes categories in
    # "Segment" column to be "small", "middle" or "large" --
    large = ["Large Luxury SUV", "Large SUV"]
    middle = ["Middle Luxury SUV", "Middle SUV"]
    small = ["Small SUV"]

    def suv_lms(x):
        if x in large:
            return "Large SUVs"
        elif x in middle:
            return "Middle SUVs"
        elif x in small:
            return "Small SUVs"
        else:
            return " "

    # -- To make working with the data easier in Power BI, we can create 'sum' columns for all of the months included in the previous year and
    # in the current year. We define which portion of the new dataset has the previous and current year data and assign them both to their respective variables.
    # Once the variables are created we can add columns to the dataset that show the `sum from previous year`, `sum from current year`, and the `total sum`. --

    # -- Defining part of data program should pull from when calculating sums for groups in previous/current year --
    group_sum_current_year = sales.iloc[:, previous_year_index_sales:]
    group_sum_previous_year = sales.iloc[:, 5:previous_year_index_sales]

    # -- Create columns containing totals for each row up to "user indicated" month. Including previous year,
    # current year, and sum of those two columns --
    sales[
        f"Sum from Jan to {user_month} {previous_year}"
    ] = group_sum_previous_year.sum(axis=1)
    sales[f"Sum from Jan to {user_month} {current_year}"] = group_sum_current_year.sum(
        axis=1
    )
    sales["Overall Total"] = (
        sales[f"Sum from Jan to {user_month} {previous_year}"]
        + sales[f"Sum from Jan to {user_month} {current_year}"]
    )

    # -- Now we can apply the `categories` function we defined above to the 'Group' column which will bucket all of the brand names into the appropriate category where necessary. --
    sales["Group"] = sales["Group"].apply(categories_1)

    # -- We then apply the `parser` function we defined above to get rid of the parentheses in the 'Segment Group' and 'Segment' columns. --
    sales["Segment Group"] = sales["Segment Group"].apply(parser_parathenses)
    sales["Segment"] = sales["Segment"].apply(parser_parathenses)

    # -- Now we can apply the other four functions we defined above to categorize the 'segment' column. --
    sales["CUV - Lux vs. Non-Lux"] = sales["Segment"].apply(cross_utility_luxury)
    sales["CUV - Lar/Mid/Sml"] = sales["Segment"].apply(cross_utility_lms)
    sales["SUV - Lux vs. Non-Lux"] = sales["Segment"].apply(suv_luxury)
    sales["SUV - Lar/Mid/Sml"] = sales["Segment"].apply(suv_lms)

    # -- The last bit of cleaning we will do is get rid of the 'Comm. Chassis' from our dataset. --
    sales = sales[sales["Segment Group"] != "Comm. Chassis"]

    # -- WARD'S PRODUCTION DATA
    # -- Removing the last four rows because they contain irrelevant data --
    production = production[:-4]
    # -- The initial Ward's dataset contains column names that are on two lines. When we read the data into python we have to cut
    # off one of those lines. For this reason some columns have similar names so python automatically names them `column`, `column.1`,
    # `column.2`, etc. The code below is taking one of the column names where this happened and renames it. --
    production.rename(columns={"Group.1": "Segment Group"}, inplace=True)

    # --The Ward's dataset contains data for Jan to Dec of the previous year but only Jan to current month of current year.
    # For this reason we need to select only the columns that are needed for the Power BI. The `labels` variable grabs the identifiers,
    # the `prev_data` grabs the previous year's data but only up to the month indicated by the user, and the `curr_data` takes all of the
    # columns that are there for the current year. Then we join everything together to make our new dataset. --
    labels = production.iloc[:, 0:5]
    prev_data = production.iloc[:, 6:previous_year_index_production]
    curr_data = production.iloc[:, 18:]
    production = labels.join(prev_data)
    production = production.join(curr_data)

    # -- Some of the columns in the dataset have the data wrapped in parentheses. This function will allow us to take one of those elements
    # as an input and output the data without any parentheses. --
    # -- Function that parses columns that have extra parentheses and returns second half of parsed string --
    def parser(x):
        name = x.split(")")[1]
        return name[1:]

    # -- Since our Power BI will have groupings, such as 'Euro' and 'Other Asian', we need to define lists that contain all of the brands
    # that fall under a certain category. Once we have our lists, we can define a function that will take a data point as an input and
    # output either a new category or the original depending on if the brand name exists in our defined lists. --
    # -- Define lists that contain all of the companies that will be grouped into other_asian and euro categories --
    other_asian = [
        "Mazda",
        "Mitsubishi",
        "Subaru",
        "Suzuki",
        "Volvo",
        "Beijing AIC",
        "Zhejiang Geely",
        "Isuzu",
        "Tata Motors",
    ]
    euro = [
        "Audi",
        "BMW",
        "Daimler",
        "Jaguar Land Rover",
        "Peugeot Citroen",
        "Porsche",
        "Volkswagen",
    ]

    # -- Function that takes "Group" column and references defined lists to change names of companies to final graph versions --
    def categories(x):
        if x in other_asian:
            return "Other Asian"
        elif x in euro:
            return "Euro"
        elif x == "Hyundai Group" or x == "Kia Motors":
            return "Hyundai/Kia"
        elif x == "Tesla Motors":
            return "Tesla"
        elif x == "Fiat Chrysler":
            return "FCA"
        elif x == "General Motors":
            return "GM"
        elif x == "Renault":
            return "RN"
        else:
            return x

    # -- To make working with the data easier in Power BI, we can create 'sum' columns for all of the months included in the previous year and in
    # the current year. We define which portion of the new dataset has the previous and current year data and assign them both to their respective variables.
    # Once the variables are created we can add columns to the dataset that show the `sum from previous year`, `sum from current year`, and the `total sum`. --
    # -- Defining part of data program should pull from when calculating sums for groups in previous/current year --
    group_sum_current_year = production.iloc[:, previous_year_index_production:]
    group_sum_previous_year = production.iloc[:, 5:previous_year_index_production]

    # -- Create columns containing totals for each row up to "user indicated" month. Including previous year,
    # current year, and sum of those two columns --
    production[
        f"Sum from Jan to {user_month} {previous_year}"
    ] = group_sum_previous_year.sum(axis=1)
    production[
        f"Sum from Jan to {user_month} {current_year}"
    ] = group_sum_current_year.sum(axis=1)
    production["Overall Total"] = (
        production[f"Sum from Jan to {user_month} {previous_year}"]
        + production[f"Sum from Jan to {user_month} {current_year}"]
    )

    # -- Now we can apply the `categories` function we defined above to the 'Group' column which will bucket all of the brand names into the appropriate category where necessary. --
    production["Group"] = production["Group"].apply(categories)

    # -- We then apply the `parser` function we defined above to get rid of the parentheses in the 'Segment Group' and 'Segment' columns. --
    production["Segment Group"] = production["Segment Group"].apply(parser)
    production["Segment"] = production["Segment"].apply(parser)

    # -- The last bit of cleaning we will do is get rid of the 'Comm. Chassis' from our dataset. --
    production = production[production["Segment Group"] != "Comm. Chassis"]

    # -- WRITE AGGREGATED CLEANED DATA TO NEW EXCEL FILE --
    # -- Now we can take all of our cleaned datasets and write them to a single excel file with each dataset on a separate tab.
    # -- Define today's date to be used for output file name --
    todaysdate = datetime.today().strftime("%m-%d-%Y")
    # -- Define filename of output excel file --
    path = output_folder
    name = f"MKR_Market_Update_{todaysdate}.xlsx"
    outpath = path + "/" + name
    # -- Define method to export dataframe --
    writer = pd.ExcelWriter(outpath, engine="xlsxwriter")

    # -- Write dataframe to sheet --
    ussaar_final.to_excel(writer, sheet_name="USSAAR")
    fleets_prev.to_excel(writer, sheet_name="Prev Year Polk Fleet Data", index=False)
    fleets_curr.to_excel(writer, sheet_name="Curr Year Polk Fleet Data", index=False)
    incentives.to_excel(writer, sheet_name="Motortrend Incentives", index=False)
    sales.to_excel(writer, sheet_name="Wards Sales", index=False)
    production.to_excel(writer, sheet_name="Wards Production", index=False)

    writer.save()


# -- function that allows user to import USSAAR History file --
def ussaar_open():
    global ussaar_file
    ussaar_file = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    ussaar_data = ttk.Label(
        root, text=ussaar_file, width=40, borderwidth=2, relief="groove"
    )
    ussaar_data.grid(row=1, column=1, padx=10, pady=10, sticky="NSEW")


# -- function that allows user to import Ward's Sales data --
def wards_sales_open():
    global wards_sales_file
    wards_sales_file = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    wards_sales_data = ttk.Label(
        root, text=wards_sales_file, width=40, borderwidth=2, relief="groove"
    )
    wards_sales_data.grid(row=2, column=1, padx=10, pady=10, sticky="NSEW")


# -- function that allows user to import Ward's Sales data --
def wards_production_open():
    global wards_production_file
    wards_production_file = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    wards_production_data = ttk.Label(
        root, text=wards_production_file, width=40, borderwidth=2, relief="groove"
    )
    wards_production_data.grid(row=3, column=1, padx=10, pady=10, sticky="NSEW")


# -- function that allows user to import Polk Fleet data from previous year --
def polk_open_prev():
    global polk_file_prev
    polk_file_prev = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    polk_data_prev = ttk.Label(
        root, text=polk_file_prev, width=40, borderwidth=2, relief="groove"
    )
    polk_data_prev.grid(row=4, column=1, padx=10, pady=10, sticky="NSEW")


# -- function that allows user to import Polk Fleet data from current year --
def polk_open_curr():
    global polk_file_curr
    polk_file_curr = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    polk_data_curr = ttk.Label(
        root, text=polk_file_curr, width=40, borderwidth=2, relief="groove"
    )
    polk_data_curr.grid(row=5, column=1, padx=10, pady=10, sticky="NSEW")


# -- function that allows user to import Polk Fleet data from current year --
def open_incentives():
    global motor_incentives
    motor_incentives = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    polk_data_curr = ttk.Label(
        root, text=polk_file_curr, width=40, borderwidth=2, relief="groove"
    )
    polk_data_curr.grid(row=6, column=1, padx=10, pady=10, sticky="NSEW")


# -- function that allows user to select location to place output file --
def find_folder():
    global output_folder
    output_folder = filedialog.askdirectory(
        initialdir="/Users/jeffjakinovich/Desktop", title="Select a folder"
    )
    folder = ttk.Label(
        root, text=output_folder, width=40, borderwidth=2, relief="groove"
    )
    folder.grid(row=9, column=1, padx=10, pady=10, sticky="NSEW")


# -- START OF GUI --
# -- create main frame where widgets will exist --
root = tk.Tk()
root.title("MKR Monthly Market Update")
root.configure(background="white")

# -- WIDGETS FOR GUI --
title_label = ttk.Label(
    root, text="MKR Monthly Market Update", background="white", font=("Arial Bold", 18)
)

ussaar_label = ttk.Label(root, text="USSAAR's Data", background="white")
ussaar_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
ussaar_import = ttk.Button(root, text="Import", width=10, command=ussaar_open)

wards_sales_label = ttk.Label(root, text="Ward's Sales Data", background="white")
wards_sales_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
wards_sales_import = ttk.Button(root, text="Import", width=10, command=wards_sales_open)

wards_production_label = ttk.Label(
    root, text="Ward's Production Data", background="white"
)
wards_production_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
wards_production_import = ttk.Button(
    root, text="Import", width=10, command=wards_production_open
)

polk_prev_label = ttk.Label(
    root, text="Polk's Data From Previous Year", background="white"
)
polk_prev_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
polk_prev_import = ttk.Button(root, text="Import", width=10, command=polk_open_prev)

polk_curr_label = ttk.Label(
    root, text="Polk's Data From Current Year", background="white"
)
polk_curr_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
polk_curr_import = ttk.Button(root, text="Import", width=10, command=polk_open_curr)

incentives_label = ttk.Label(root, text="Incentive's Data", background="white")
incentives_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
incentives_import = ttk.Button(root, text="Import", width=10, command=open_incentives)

month_label = ttk.Label(root, text="Pick the month being analyzed", background="white")
month_result = tk.StringVar()
month_combo = ttk.Combobox(root, textvariable=month_result)
month_combo["values"] = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
month_combo.current(0)

separator = ttk.Separator(root, orient="horizontal")

source_label = ttk.Label(root, text="Output Folder", background="white")
source_blank = ttk.Label(root, width=40, borderwidth=2, relief="groove")
source_import = ttk.Button(root, text="Find", width=10, command=find_folder)

run_button = ttk.Button(root, text="Run", width=15, command=run)

# progress_bar = ttk.Progressbar(root, length=100, mode='determinate')

# -- LAYOUT OF GUI --
title_label.grid(row=0, column=1, padx=10, pady=10, sticky="NSEW")

ussaar_label.grid(row=1, column=0, padx=10, pady=10, sticky="NSEW")
ussaar_blank.grid(row=1, column=1, padx=10, pady=10, sticky="NSEW")
ussaar_import.grid(row=1, column=2, padx=10, pady=10, sticky="NSEW")

wards_sales_label.grid(row=2, column=0, padx=10, pady=10, sticky="NSEW")
wards_sales_blank.grid(row=2, column=1, padx=10, pady=10, sticky="NSEW")
wards_sales_import.grid(row=2, column=2, padx=10, pady=10, sticky="NSEW")

wards_production_label.grid(row=3, column=0, padx=10, pady=10, sticky="NSEW")
wards_production_blank.grid(row=3, column=1, padx=10, pady=10, sticky="NSEW")
wards_production_import.grid(row=3, column=2, padx=10, pady=10, sticky="NSEW")

polk_prev_label.grid(row=4, column=0, padx=10, pady=10, sticky="NSEW")
polk_prev_blank.grid(row=4, column=1, padx=10, pady=10, sticky="NSEW")
polk_prev_import.grid(row=4, column=2, padx=10, pady=10, sticky="NSEW")

polk_curr_label.grid(row=5, column=0, padx=10, pady=10, sticky="NSEW")
polk_curr_blank.grid(row=5, column=1, padx=10, pady=10, sticky="NSEW")
polk_curr_import.grid(row=5, column=2, padx=10, pady=10, sticky="NSEW")

incentives_label.grid(row=6, column=0, padx=10, pady=10, sticky="NSEW")
incentives_blank.grid(row=6, column=1, padx=10, pady=10, sticky="NSEW")
incentives_import.grid(row=6, column=2, padx=10, pady=10, sticky="NSEW")

month_label.grid(row=7, column=0, padx=10, pady=10, sticky="NSEW")
month_combo.grid(row=7, column=1, padx=10, pady=10, sticky="W")

separator.grid(row=8, columnspan=3, padx=10, pady=10, sticky="NSEW")

source_label.grid(row=9, column=0, padx=10, pady=10, sticky="NSEW")
source_blank.grid(row=9, column=1, padx=10, pady=10, sticky="NSEW")
source_import.grid(row=9, column=2, padx=10, pady=10, sticky="NSEW")

run_button.grid(row=10, column=0, columnspan=3, padx=275, pady=10, sticky="EW")

# progress_bar.grid(row=9, column=0, columnspan=4, sticky='NSEW')


root.mainloop()

print(month)
