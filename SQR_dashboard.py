"""
Finalized Code Date: 05/06/2024
Author: Ojasvi Mukkoti

This script generates a dashboard for the Supplier Quality Ratios (SQR).

Dashboard contains:
1. DMR to PO ratio percentages
2. Monthly and yearly SQR %'s
3. Bar charts of the top 10 SQR Vendor Ratios
4. Process Behavior Charts for specfic vendors and Years. 

- SQR(SupplierQuality Ratio): the # of DMRs generated by a supplier
    to the numper of POs submitted w/ a supplier
"""
import pandas as pd
import plotly.express as px
import streamlit as st
from io import BytesIO

st.set_page_config(layout='wide',
                   initial_sidebar_state="expanded")
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
    if option == 'Year':
        date_ = df[f'{date_column}'].unique()
        unique_DMRs_years = []
        for date_string in date_:
            year = date_string[:4]
            if year not in unique_DMRs_years: 
                unique_DMRs_years.append(year)
        return unique_DMRs_years
    elif option == 'Month':
        date_ = df[f'{date_column}'].unique()
        unique_DMRs_months = []
        for date_string in date_:
            month = date_string[:4]
            if month not in unique_DMRs_months: 
                unique_DMRs_months.append(month)
        return unique_DMRs_months
    
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
    count_dict = {}
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
        plotly.graph_objs.Figure
            Plotly figure object containing the horizontal bar chart.
    """
    #need to do through the DF for specfic year
    #find the top 8 suppliers w/ highest SQR%
    sorted_year_SQR = df_SQR[['Vendor', year]].dropna(subset=[year]).sort_values(by=year, ascending=False)
    top8_SQRs = sorted_year_SQR.head(8)

    # Create the horizontal bar chart using Plotly Express
    fig = px.bar(top8_SQRs, x=year, y='Vendor', orientation='h', title=f'{year} Supplier Quality Ratios',
                 labels={'Vendor': 'Vendor', year: 'SQR Ratio'})

    fig.update_traces(text=top8_SQRs[year].map('{:.2f}'.format), textposition='outside')

    # Customize layout
    fig.update_layout(
        yaxis={'categoryorder': 'total ascending'},  # Ensure bars are sorted by SQR Ratio
        xaxis_title='SQR Ratio',
        yaxis_title='Vendor',
        yaxis_tickfont=dict(size=10),
        height=500,
        width=500
    )
    return fig


@st.cache_data
# Function to calculate and plot Process Behavior Chart (PBC) for specific vendors and year
def calculate_and_plot_pbc(vendor, year, raw_df):
    """
    Calculates and plots the Process Behavior Chart (PBC) for specific vendors and year.

    Parameters:
        vendor: string
            Name of the vendor.
        year: int
            Year for which the chart is generated
        raw_df: Dataframe
            Raw dataframe containing the DMRs and dates. It is the DMR log.
    Returns:
        BytesIO: Buffer contianing the plotted PBC chart image. Will be used to download the PBC in the dashbaord. 
    """
    # Convert 'Date' column to datetime
    raw_df['Date'] = pd.to_datetime(raw_df['Date'], errors='coerce')

    # Gather data for the specified vendor and year
    vendor_data = raw_df[(raw_df['Vendor'] == vendor) & (raw_df['Date'].dt.year == year)]

    # Extract 'Year-Month' from 'Date'
    vendor_data['Year-Month'] = vendor_data['Date'].dt.to_period('M').astype(str)

    # Grouping by 'Year-Month' and counting the number of DMRs
    dmr_per_month = vendor_data.groupby('Year-Month').size().reset_index(name='DMRs Count')
    # Sorting the dataframe by 'Year-Month' order
    dmr_per_month = dmr_per_month.sort_values('Year-Month')

    # Calculating MEAN
    mean_dmr = dmr_per_month['DMRs Count'].mean()

    # Calculating moving ranges
    moving_ranges = [0] + [abs(next_val - current_val) for current_val, next_val in zip(dmr_per_month['DMRs Count'][:-1], dmr_per_month['DMRs Count'][1:])]

    dmr_per_month['Moving Ranges'] = moving_ranges
    average_moving_ranges = moving_ranges[-1]

    # Calculating the process limits
    UPL = mean_dmr + (2.66 * average_moving_ranges)
    LPL = mean_dmr - (2.66 * average_moving_ranges)
    if LPL < 0:
        LPL = 0
    else:
        LPL = mean_dmr
    URL = average_moving_ranges * 3.27

    dmr_per_month['UPL'] = UPL
    dmr_per_month['LPL'] = LPL
    dmr_per_month['URL'] = URL

    # Plotting PBC chart using Plotly Express
    fig = px.line(dmr_per_month, x='Year-Month', y='DMRs Count', title=f'PBC For {vendor} in {year}',
                  labels={'DMRs Count': 'Counts', 'Year-Month': 'Index'},
                  template='plotly_white',
                  markers=True)
    fig.add_hline(y=mean_dmr, line_dash="dash", line_color="blue", annotation_text=f'Mean: {mean_dmr:.3f}', annotation_position="bottom right")
    fig.add_hline(y=UPL, line_dash="dash", line_color="red", annotation_text=f'UPL: {UPL:.3f}', annotation_position="bottom right")
    fig.add_hline(y=LPL, line_dash="dash", line_color="red", annotation_text=f'LPL: {LPL:.3f}', annotation_position="bottom right")

    # Set the figure layout to add a black border
    fig.update_layout(
        margin=dict(l=20, r=20, t=60, b=20),
        paper_bgcolor="white",  # Set background color to transparent
        plot_bgcolor="white",   # Set plot color to transparent
        xaxis=dict(linecolor="black", linewidth=2),  # Set X-axis line color and width
        yaxis=dict(linecolor="black", linewidth=2),  # Set Y-axis line color and width
        )

    st.plotly_chart(fig, use_container_width=True)

    buffer = BytesIO()
    fig.write_image(buffer, format='png', scale=2)
    buffer.seek(0)
    return buffer

DMR_df = pd.read_excel("Example_DMR_Log.xlsx")
supplier_PO_df = pd.read_csv("Example_Supplier PO List.csv", encoding='latin1')

#checking if the DMR and PO dataframe has data in them
if DMR_df is not None and supplier_PO_df is not None:
    #getting rid of any unnamed columns
    unnamed_columns = [col for col in DMR_df.columns if 'Unnamed' in col]
    DMR_df = DMR_df.drop(columns=unnamed_columns)
    DMR_df['Date'] = DMR_df['Date'].astype('str')

    #getting the Supplier PO data into a dataframe
    supplier_PO_df['P.O. Date'] = pd.to_datetime(supplier_PO_df['P.O. Date'])
    supplier_PO_df['P.O. Date'] = supplier_PO_df['P.O. Date'].astype('str')

    dmrs_years = generate_unique_list(DMR_df, 'Year', 'Date')
    PO_years = generate_unique_list(supplier_PO_df, 'Year', 'P.O. Date')

    # print("\n", "\n")
    DMRs_yrs_count = generate_count_dict(DMR_df, 'Year', 'Date', dmrs_years)
    # print("1: ",DMRs_yrs_count)
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
    # print(DMR_PO_perc_yr)

    #need to go through every month IN EACH YR and get the SQR calculations
    unique_months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    DMRs_month_count = generate_count_dict(DMR_df, 'Month', 'Date', unique_months)
    PO_month_count = generate_count_dict(supplier_PO_df, 'Month', 'P.O. Date', unique_months)

    #creating DMR to PO percentage ratio dictionary
    DMR_PO_perc_month = {}

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

    #START OF: Code for the dashboard continued
    st.info("This is a sample of what the dashboard looks like. The official one requires users to upload\
    the needed files.")
    st.title("Supplier Quality Ratios Dashboard")
    col = st.columns([1,1,1], gap='small')
    #creating row that contains the SQR percentages from the SQR calcualtions excel sheet
    with col[0]:
        with st.expander("Yearly Supplier Quality Ratios Percentages"):
            st.write(df_yr)
    with col[1]:
        with st.expander("Monthly Supplier Quality Ratios Percentages"):
            st.write(df_month)
    with col[2]:
        with st.expander("Vendor SQR Percentages"):
            st.write(df_vendor_SQR_ratios)

    st.subheader("Yearly Top 8 Supplier Quality Ratios")
    col1 = st.columns([1,1])

    # Display the bar charts for each year
    unique_years = sorted(set(PO_years).intersection(dmrs_years))
    n = 0
    for year in unique_years:
        with col1[n]:
            # Generate the bar chart
            bar_chart = generate_SQR_bar_chart(df_vendor_SQR_ratios, int(year))
            # Display the bar chart using Streamlit's st.plotly_chart function
            st.plotly_chart(bar_chart, use_container_width=True)
            n+=1
            if n ==2:
                n=0
    
    st.subheader('Process Behvaior Charts of DMR Log for Specfic Vendors and Years')
    #reading in the top 20 key suppliers to filter through the suppliers in the DMR log for the PBC 
    top_20_key_suppliers = pd.read_csv("Example_Top_Key_Suppliers.csv")
    #making the suppliers capitalized for consistency
    top_20_key_suppliers["Top 20 Key Suppliers"] = top_20_key_suppliers['Top 20 Key Suppliers'].str.upper()
    DMR_df["Vendor"] = DMR_df['Vendor'].str.upper()
    #filtering the suppliers in the DMR dataframe
    filtered_dmr_df = DMR_df[DMR_df['Vendor'].isin(top_20_key_suppliers['Top 20 Key Suppliers'])]

    col2 = st.columns([1,1,1])
    #crating selectboxes for the use to select a vendor and a respective year
    with col2[0]:
        vendor_interested = st.selectbox("Select a Vendor: ", filtered_dmr_df['Vendor'].unique())
    with col2[1]:
        year_interested = st.selectbox("Select Year: ", pd.to_datetime(filtered_dmr_df['Date'], errors='coerce').dt.year.dropna().unique())
    try:
        #making the PBC chart
        plot_buffer = calculate_and_plot_pbc(vendor_interested, year_interested, filtered_dmr_df)
        with col2[2]:
            #making a Download button for the PBC chart
            st.write("Download PBC Chart")
            st.download_button(
                label = "Click to download the PBC Chart",
                data = plot_buffer,
                file_name = f"PBC_{vendor_interested}_{year_interested}.png",
                mime='image/png'
            )
    #if an error while making the PBC chart occurs, an error message comes up
    # except ValueError:
    #     st.warning("Error occured while plotting. Please check files. *Select a different vendor or year.*")
    except ValueError as e:
        st.warning(f"ValueError: {str(e)} - Error occurred while plotting. Please check files. *Select a different vendor or year.*")
    except Exception as e:
        st.warning(f"Exception: {str(e)} - An unexpected error occurred. Please check the data and try again.")

    colf = st.columns([1,1])
    #creating expanders for the raw DMR and PO data
    with colf[0]:
        with st.expander(label="Expand to see the Raw DMR Data"):
            st.write(DMR_df)
    with colf[1]:
        with st.expander(label="Expand to see the Raw PO Data"):
            st.write(supplier_PO_df)
else:
    st.warning('Please go to: **"P:\Quality\DMR"** to upload the *DMR Log*.\nPlease go to **"P:\Quality\Quality Co-op - Winter 2024\Supplier Quality Ratios"** to upload the *Supplier PO List*.')
