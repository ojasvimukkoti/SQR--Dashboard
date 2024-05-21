"""
Finalized Code Date: 05/06/2024
Author: Ojasvi Mukkoti

This script generates Supplier Quality Ratios (SQR) based on DMR (Defective Material Report) and Supplier PO (Purchase Order) data,
and exports the results to an Excel.

It includes the functions of:
1. Generate unique years or months from a Dataframe's date column.
2. Generate counts of DMRs or POs for each unique year or month.
3. Generate a bar chart of the top suppliers by SQR percentage for a specfic year. 

"""
import pandas as pd
import os
import matplotlib.pyplot as plt
import streamlit as st

#Setting Streamlit page configuration
st.set_page_config(layout='wide',
                   initial_sidebar_state="expanded")

#Function generates a dictionary of counts for each unqiue year or month
@st.cache_data
def generate_unique_list(df, option, date_column):
    """
    Generate a list of unique years or months from a Dataframe's date column.

    Parameters:
        df: Dataframe
            Dataframe contianing the data column needed
        option: string
            'Year' or 'Month' to generate list of unqiue years or months
        date_column: string
            Name of the data column in the Dataframe

    Returns:
    List of unique years or months
    """
    #If-statement that checks the option parameter
    if option == 'Year': #Calculates the unique years and complies into a list if option is set to 'Year'
        date_ = df[date_column].unique()
        unique_DMRs_years = [] #creates empty list
        for date_string in date_: #goes through every string in the date_ dataframe
            year = date_string[:4] #extracts the year
            if year not in unique_DMRs_years: #checks if current year is in unique list
                unique_DMRs_years.append(year) #adds the current year into list
        return unique_DMRs_years

#Function that generates a dictionary of counts for each unique year or month
@st.cache_data
def generate_count_dict(df, option, date_column, unique_list):
    """
    Generates a dictionary of counts for each unique year or month.

    Parameters:
        df: dataframe
            Dataframe containing the date column.
        option: string
            'Year' or 'Month' to generate counts for unique years or months.
        date_column: string
            Name of the date column in the Dataframe.
        unqiue_list: list
            list of unique years or month
    
    Returns:
        dict: Dictionary containing counts for each unique year or month
    """
    count_dict = {} #creating an empty dictionary
    if option == 'Year':
        # Going through each year in unique list
        for year in unique_list:
                # Getting the count for the year in the DataFrame
            count = (df[date_column].str[:4] == year).sum()
                # Putting the year and its count into the dictionary
            count_dict[year] = count
        return count_dict

    elif option =="Month":
        unique_years = df[date_column].str[:4].unique()
        for yr in unique_years:
            # dictionary that stores the month counts for current year
            month_counts = {}
            # iterates through the unique months
            for month in unique_list:
                # gets the counts for each month
                count = (df[df[date_column].str[:4]==yr][date_column].str[5:7]==month).sum()
                # stores the month counts in month dictionary
                month_counts[month] = count
            # stores month_counts dictionary for current year in official dictionary
            count_dict[yr] = month_counts
        return count_dict

#Function that generates a bar chart of the top 8 suppliers by SQR percentages for a specific year 
@st.cache_data
def generate_SQR_bar_chart(df_SQR, year):
    """
    Generate a horizontal bar chart of the top 8 suppliers with the highest SQR %

    Parameters:
        df_SQR: dataframe
            Contains the SQR percentages for each vendor and year
        year: string
            Year for whcih bar chart will be generated
    Returns:
        None
    """
    #need to do through the DF for specfic year
    #find the top 8 suppliers w/ highest SQR%
    sorted_year_SQR = df_SQR.sort_values(by=year,ascending= False)
    top8_SQRs = sorted_year_SQR.head(8)

    plt.figure(figsize=(15,8))
    bars = plt.barh(top8_SQRs['Vendor'], top8_SQRs[year])
    plt.xlabel('SQR Ratio')
    plt.ylabel("Vendor")
    plt.title(f"{year} Supplier Quality Ratios")
    plt.gca().invert_yaxis() #inverting y-axis to have highest ratios at top
    
    plt.yticks(fontsize=7)
    plt.yticks(rotation=55)

    for bar, pct in zip(bars, top8_SQRs[year]):
        plt.text(bar.get_width(), bar.get_y()+bar.get_height()/2, f'{pct:.2f}', ha='left', va='center')

    plt.axvline(color='white', zorder=2)

    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    plt.savefig(f'bar_chart_{year}.png')

#UNC path to access the DMR Log
server_name= "empowering.apcd.local"
share_name='public'
folder_name = "Quality"
DMR_folder_name = 'DMR'
DMR_filename = 'DMR MASTER LIST - USE THIS LOG.xlsx'

#getting the file path for DMR Log
file_path = os.path.join(r'\\', server_name, share_name, folder_name, DMR_folder_name, DMR_filename)
#reading in the Log into a dataframe
DMR_df = pd.read_excel(file_path)

#getting rid of any unnamed columns
unnamed_columns = [col for col in DMR_df.columns if 'Unnamed' in col]
DMR_df = DMR_df.drop(columns=unnamed_columns)
DMR_df['Date'] = DMR_df['Date'].astype('str')

#getting the Supplier PO data into a dataframe
PO_filename = "Supplier PO List.csv"
supplier_PO_df = pd.read_csv(PO_filename, encoding='latin1')
supplier_PO_df['P.O. Date'] = pd.to_datetime(supplier_PO_df['P.O. Date'])
supplier_PO_df['P.O. Date'] = supplier_PO_df['P.O. Date'].astype('str')

#getting the unique years list for the DMR and PO data
dmrs_years = generate_unique_list(DMR_df, 'Year', 'Date')
PO_years = generate_unique_list(supplier_PO_df, 'Year', 'P.O. Date')

#getting the count dictionary for each year for DMR and PO data
DMRs_yrs_count = generate_count_dict(DMR_df, 'Year', 'Date', dmrs_years)
PO_yrs_count = generate_count_dict(supplier_PO_df, 'Year', 'P.O. Date', PO_years)

#creating DMR to PO percentage ratio dictionary
DMR_PO_perc_yr = {}
#calcualting DMR to PO percentage ratio for each year
for year in dmrs_years:
    #checking if year is in PO_count dictionary
    if year in PO_yrs_count:
        #gets percentage ratio
        ratio = (DMRs_yrs_count[year]/PO_yrs_count[year])*100
        DMR_PO_perc_yr[year] = ratio 
    else:
        DMR_PO_perc_yr[year]=None

#need to go through every month IN EACH YR and get the SQR calculations
unique_months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

#getting count dictionary for each month for DMR and PO data
DMRs_month_count = generate_count_dict(DMR_df, 'Month', 'Date', unique_months)
PO_month_count = generate_count_dict(supplier_PO_df, 'Month', 'P.O. Date', unique_months)

#creating DMR to PO percentage ratio dictionary
DMR_PO_perc_month = {}

#calculating the DMR to PO ratio but for months. Same logic used as the DMR to PO ratio for years.
for year in sorted(set(DMRs_month_count.keys()) & set(PO_month_count.keys())):
    year_ratio = {}
    for month in PO_month_count[year]:
        if month in PO_month_count[year]:
            dmr_count = DMRs_month_count[year][month]
            po_count = PO_month_count[year][month]
            if po_count !=0:
                ratio = (dmr_count/po_count) *100
                year_ratio[month]=ratio
    DMR_PO_perc_month[year]=year_ratio

#converting the yearly and monthly ratios to dataframes
df_yr = pd.DataFrame(DMR_PO_perc_yr, index = ['SQR Percentage'])
#making sure that the only years are numbers & nothing else
df_yr = df_yr[[col for col in df_yr.columns if col.isdigit()]]
df_month = pd.DataFrame(DMR_PO_perc_month)
df_month.index.name='Month'

#rounding the SQR percents to 1 decimal point
df_yr = df_yr.round(1)
df_month = df_month.round(1)

#creating dataframe that holds the vendors DMR count and the coresponding year
DMR_df['Year'] = DMR_df['Date'].str[:4]
# Filtering out rows where 'Year' is not a valid year
DMR_df = DMR_df[DMR_df['Year'].str.isdigit()]
# Converting 'Year' to integer type
DMR_df['Year'] = DMR_df['Year'].astype(int)
# Grouping the DataFrame by 'Year' and 'Vendor' and counting the occurrences
vendors_yr_DMR_count_df = DMR_df.groupby(['Year', 'Vendor']).size().reset_index(name='DMR Count')

#creating Dataframe for the POs vendor count 
supplier_PO_df['Year'] = supplier_PO_df['P.O. Date'].str[:4]
# Filtering out rows where 'Year' is not a valid year
supplier_PO_df = supplier_PO_df[supplier_PO_df['Year'].str.isdigit()]
# Converting 'Year' to integer type
supplier_PO_df['Year'] = supplier_PO_df['Year'].astype(int)
# Grouping the DataFrame by 'Year' and 'Vendor' and counting the occurrences
vendors_yr_PO_count_df = supplier_PO_df.groupby(['Year', 'Vendor Name']).size().reset_index(name='DMR Count')

#creating dictionaries to hold the DMR + PO vendors counts
vendor_DMR_counts = {}
vendor_PO_counts = {}

#iterates through each row in the dataframe
#row = pandas series of the vendors year data
for index, row in vendors_yr_DMR_count_df.iterrows():
    #extracting the year, vendor, + count from the row Series
    year = row['Year']
    vendor = row['Vendor']
    count = row['DMR Count']
    #checking if the given vendor is present as a key in 'vendor_DMR_counts'
    if vendor not in vendor_DMR_counts:
        #initialiszes an empty dictionary for that vendor
        vendor_DMR_counts[vendor] = {}
    #adds an entry to the dictionary
    vendor_DMR_counts[vendor][year] = count

#same thing as above but for the PO dictionary
for index, row in vendors_yr_PO_count_df.iterrows():
    year = row['Year']
    vendor = row['Vendor Name']
    count = row['DMR Count']
    if vendor not in vendor_PO_counts:
        vendor_PO_counts[vendor]={}

    vendor_PO_counts[vendor][year] = count

#creating DMR to PO percentage ratio dictionary
vendor_SQR_ratios = {}
#iterates over the items in 'vendor_DMR_counts'
for vendor, dmr_counts_per_year in vendor_DMR_counts.items():
    #checking if vendor is a part of 'vendor_SQR_ratios'
    if vendor not in vendor_SQR_ratios:
        vendor_SQR_ratios[vendor] = {} #creates new dictionary slot for new vendor
    #iterates through the years in dictionary
    for year in dmr_counts_per_year:
        #checks if vendor exists in 'vendor_PO_counts' dictionary and if year exists for that vendor
        if vendor in vendor_PO_counts and year in vendor_PO_counts[vendor]:
            #retrives the PO count for current vendor and year
            po_count = vendor_PO_counts[vendor][year]
            #checks if po_count is not zero
            if po_count !=0:
                #gets dmr count for current vendor and year
                dmr_count = dmr_counts_per_year[year]
                #calculates the SQR percentage
                ratio = (dmr_count/po_count) *100
                #stores the calculated SQR percentage in the 'vendor_SQR_ratio' dictionary
                vendor_SQR_ratios[vendor][year]=ratio
        else:
            #assigns None to the SQR ratio
            vendor_SQR_ratios[vendor][year]= None

#converting the 'vendor_SQR_ratios' dictionary to a dataframe
df_vendor_SQR_ratios = pd.DataFrame(vendor_SQR_ratios)
#transposing dataframe so it is easier to use
df_vendor_SQR_ratios = df_vendor_SQR_ratios.T

df_vendor_SQR_ratios.reset_index(inplace=True)
#renaming columns to have vendor
df_vendor_SQR_ratios.rename(columns={'index':'Vendor'}, inplace = True)

#using UNC path to export the excel sheet to the Supplie Quality Ratios folder
server_name= "empowering.apcd.local"
share_name='public'
folder_name = "Quality"
subfolder1 = "Quality Co-op - Winter 2024"
subfolder2 = "Supplier Quality Ratios"
file_name = 'SQR Calculation.xlsx'

#creating file_path
file_path = os.path.join(r'\\', server_name, share_name, folder_name,subfolder1 ,subfolder2, file_name)

#exporting dataframes and bar charts to Excel
with pd.ExcelWriter(file_path) as writer:
    df_yr.to_excel(writer, sheet_name='Yearly SQR Percentages')
    df_month.to_excel(writer, sheet_name='Monthly SQR Percentages')
    df_vendor_SQR_ratios.to_excel(writer, sheet_name = "Vendor SQR Percentages")

   # Getting the unique years in PO and DMR files to create bar charts
    unique_years = sorted(set(PO_years).intersection(dmrs_years))

    for year in unique_years:
        #generates the SQR bar chart
        generate_SQR_bar_chart(df_vendor_SQR_ratios, int(year))

        # Creating new sheet in Excel to place the chart
        bar_chart_sheet = pd.DataFrame()
        bar_chart_sheet.to_excel(writer, sheet_name=f'Bar Chart {year}')

        # Inserting the chart image into the Excel sheet
        workbook = writer.book
        worksheet = writer.sheets[f'Bar Chart {year}']
        worksheet.insert_image('A1', f'bar_chart_{year}.png')
