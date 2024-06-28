import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from numpy import sort 
import calendar
import io
from st_aggrid import GridOptionsBuilder, AgGrid, ColumnsAutoSizeMode
import math
from PIL import Image

# Set page configuration to wide format
st.set_page_config(layout="wide")

#Add GEA image
img_1 = Image.open(r"logogea.png")


st.markdown(
    """
    <style>
    html,body{
    padding:0;
    margin:0;
    }
    
      .title-wrapper {
           font-size: 3em; /* Adjust font size */
            color: transparent; /* Transparent text color */
            margin: 10px 60px;
            font-weight: bold; /* Font weight */
            background-image: linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%) ; /* Gradient background */
             -webkit-background-clip: text; /* Clip text to background */
            background-clip: text;
        }

    /* Style for tabs */
    button[data-baseweb="tab"] {
    color: black; /* Text color */
    font-size: 25px; /* Font size */
    padding: 10px 20px; /* Padding inside the button */
    cursor: pointer; /* Pointer cursor on hover */
    transition: transform 0.3s ease, box-shadow 0.3s ease; /* Smooth transition for hover effects */
    font-weight: bold; /* Bold text */
    outline: none; /* Remove default outline */
    text-decoration:none;
}
 /* Active tab */
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #304CD7;
        font-weight: bold;  /* Make the font bold for the active tab */
        transition: font-weight 0.3s;
    }
      /* Hover effect */
    button[data-baseweb="tab"]:hover {
        font-weight: bold;  /* Ensure the text is bold on hover */
        color: #304CD7;
    }


div[data-testid="stFileUploader"]
{

align-items:center;
justify-content:center;
width: 100%;
padding: 5px;
}

#select-the-sofi-excel-output-file{
   color:transparent;
   font-size: 30px;
   letter-spacing: 1px;
   width:40%;
    background-image: linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%);
   -webkit-background-clip: text; /* Clip text to background */
    background-clip: text;
    padding: 5px; /* Add padding */
    font-weight:bold;
  
   
}
#adjust-the-input-file-parameters
{
color:transparent;
   
   font-size: 30px;
   letter-spacing: 1px;
   width:40%;
   background-image: linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%) ; /* Gradient background */
             -webkit-background-clip: text; /* Clip text to background */
            background-clip: text;
    padding: 5px; /* Add padding */
    
    font-weight:bold;
}
/* For first container*/
.st-emotion-cache-bct9qr.e1f1d6gn0 {
    height: 240px;
    background: white;
    border: 1px solid #304CD7;
    border-radius: 12px; /* Rounded corners */
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Lighter box shadow */
    color: #304CD7; /* Text color changed to match border */
    font-size: 18px; /* Text size */
    margin-bottom:15px;
    transition: transform 0.3s ease, box-shadow 0.3s ease; 
}

.st-emotion-cache-bct9qr.e1f1d6gn0:hover {
    transform: translateY(-2px); /* Move up on hover */
    box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2); /* Increase shadow on hover */
}
/*For second Container*/
.st-emotion-cache-isgwfk.e1f1d6gn0{
 height: 290px;
    background: white;
    border: 1px solid #304CD7;
    border-radius: 12px; /* Rounded corners */
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Lighter box shadow */
    color: #304CD7; /* Text color changed to match border */
    font-size: 18px; /* Text size */
    margin-bottom:15px;
    transition: transform 0.3s ease, box-shadow 0.3s ease; 
}
.st-emotion-cache-isgwfk.e1f1d6gn0:hover {
    transform: translateY(-2px); /* Move up on hover */
    box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2); /* Increase shadow on hover */
}


div[data-testid="stFileUploader"] p {
    margin: 0; /* Remove default margin */
    padding: 2px 0; /* Add padding to paragraph */
    font-size: 18px; /* Increase font size */
    color: black; /* Set paragraph text color to white */
    font-weight: bold; /* Make the paragraph text bold */
}
button[kind="secondary"][data-testid="baseButton-secondary"] {
    
    border: 1px solid #304CD7; /* Border */
    border-radius: 12px; /* Rounded corners */
    color: black; /* Text color */
    font-size:18px; /* Font size */
    padding: 10px 20px; /* Padding inside the button */
    cursor: pointer; /* Pointer cursor on hover */
    transition: transform 0.3s ease, box-shadow 0.3s ease; /* Smooth transition for hover effects */
      
}
button[kind="secondary"][data-testid="baseButton-secondary"]:hover {
     background:linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%);
     color:white;
  
}
   .st-emotion-cache-1whx7iy.e1nzi1vr4 p{
        border: 2px solid #3498db;
        padding: 15px;
        border-radius: 10px;
        color: #2c3e50;
        font-size: 18px;
    }
       [data-testid="stWidgetLabel"] p{
        font-size: 18px;
        padding: 5px;
        
    }
     [data-testid="stTextInput-RootElement"]{
        border: 1px solid lightgrey ;
        margin-left:5px;
        padding: 10px;
        border-radius: 5px;
        width:99%;
        }
           [data-testid="stTextInput-RootElement"]:hover {
        border-color: #304CD7;
      
    }
     [data-testid="stNumberInputContainer"] {
        margin-left:6px;
        font-size: 18px;
        border-radius: 5px;
        border: 1px solid lightgrey;
    }
      [data-testid="stNumberInputContainer"]:hover{
        border-color: #304CD7 ;
        }
   
     [data-testid="stNumberInput-StepDown"] {
        background-color: #f0f0f0;
        transition: border-color 0.3s, background-color 0.3s, color 0.3s;
        color: #3498db;
        cursor: pointer;
    }
    [data-testid="stNumberInput-StepDown"]:hover:enabled {
        background-color:#304CD7 ;
        color:white;
    }
   

       [data-testid="stNumberInput-StepUp"] {
        
        background-color: #f0f0f0;
        transition: border-color 0.3s, background-color 0.3s, color 0.3s;
        color: #3498db;
        cursor: pointer;
    }
    [data-testid="stNumberInput-StepUp"]:hover:enabled {
        background-color:#304CD7 ;
        color:white;
    }
    #upload-comments-excel-file {
      color:transparent;
   font-size: 30px;
   margin-bottom:20px;
   margin-top:5px;
   letter-spacing: 1px;
   width:40%;
   background-image: linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%);             -webkit-background-clip: text; /* Clip text to background */
            background-clip: text;
    padding: 5px; /* Add padding */
    font-weight:bold;    }
   }
   #input-data{
       color:transparent;
   font-size: 30px;
   margin-bottom:20px;
   margin-top:5px;
   letter-spacing: 1px;
   width:40%;
   background-image: linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%);             -webkit-background-clip: text; /* Clip text to background */
            background-clip: text;
    padding: 5px; /* Add padding */
    font-weight:bold; 
   
   }
   /*Styling For Tab2*/
   div[data-testid="stText"].st-emotion-cache-y4bq5x ewgb6651{
     color: black;
        font-style: italic;
   }
   .gradient-text {
   
   margin-left:5px;
   color:black;
   font-style:italic;
}
 ul.custom-bullets {
        list-style: none;
        padding:5px;
        
    }
     ul.custom-bullets li{
     font-size:18px;
     }
    ul.custom-bullets li::before {
        content: "➔ ";  /* Unicode arrow character */
        color:#304CD7 ;   /* Custom bullet color */
        font-weight: bold;  /* Custom bullet weight */
        display: inline-block; 
        width: 2em;
        margin-left: -1em;
    }
  
    .st-emotion-cache-1f5832m.e1f1d6gn0
    {
      
        height: 640px;
    background: white;
    border: 1px solid #304CD7;
    border-radius: 12px; /* Rounded corners */
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Lighter box shadow */
    color: black; /* Text color changed to match border */
    font-size: 18px; /* Text size */
    margin-bottom:15px;
    transition: transform 0.3s ease, box-shadow 0.3s ease; 
    }
   .st-emotion-cache-1f5832m.e1f1d6gn0:hover {
    transform: translateY(-2px); /* Move up on hover */
    box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2); /* Increase shadow on hover */
}
    .st-emotion-cache-32pt9z.e1f1d6gn0
    {
       height: 140px;
    background: white;
    border: 1px solid #304CD7;
    border-radius: 12px; /* Rounded corners */
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Lighter box shadow */
    color: black; /* Text color changed to match border */
    font-size: 18px; /* Text size */
    margin-bottom:15px;
    transition: transform 0.3s ease, box-shadow 0.3s ease;
    }
    .st-emotion-cache-32pt9z.e1f1d6gn0:hover {
    transform: translateY(-2px); /* Move up on hover */
    box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2); /* Increase shadow on hover */
}
h3.subheader#decision-criteria {
    width:20%;
    background: linear-gradient(to right,rgb(2,0,36) 0%,rgb(63, 76, 183) 35%,rgb(0,212,255) 100%); /* Gradient background */
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 30px;
    font-weight: bold;
}
div[data-testid="stImage"].st-emotion-cache-1v0mbdj {
    display: flex;
    align-items: center;
    justify-content: center;
}

/* Targeting the img within the specified div */
.st-emotion-cache-1kyxreq.e115fcil2 {
    width: 30px;
  
}

    </style>
    """,
    unsafe_allow_html=True
)


#st.image(img_!,width=300,use_column_width=True)
header_col_1,header_col_2=st.columns([1,4])
header_col_1.image(img_1,width=300,use_column_width=True)
#title
header_col_2.markdown('<div class="title-wrapper">Environment Reporting Control Testing App</div>', unsafe_allow_html=True)

tab1,tab2,tab3=st.tabs(["Upload SOFI EXL File","Decision Criteria","Upload Comments "])

with tab1:
    with st.container(height=250,border=True):
        # file uploader
        st.markdown('<div class="file-uploader">',unsafe_allow_html=True)
        st.markdown('<h3 id="select-the-sofi-excel-output-file">Select the SOFI Excel Output File</h3>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("", type=["xlsx", "xls"])
       
    with st.container(height=300,border=True):
        #input parameters container
        st.subheader("Adjust the input file parameters")
        # Adding Columns for Year Inputs
        col_y1, col_y2 = st.columns(2)
        # Get the current year
        current_year = datetime.now().year
        # Get the previous year
        previous_year = current_year - 1
        pre_yr = col_y1.text_input('Previous Year', value=previous_year)
        curr_yr = col_y2.text_input('Current Year', value=current_year)
        # Adding Columns for Inputs in the next row
        col1, col2, col3, col4 = st.columns(4)
        # Specify the starting row and columns
        start_row = col1.number_input('Row starting number', value=6, min_value=0)
        start_column = col2.number_input('Column starting number', step=1, value=1, min_value=1, max_value=1)
        end_column = col3.number_input('Column Ending number', step=1, value=2, min_value=2, max_value=10)
        env_select = col4.selectbox('Select Environment', ['Energy', 'Water', 'Waste'])
       

    # Define new column names
    new_column_names = ['Raw_Column', 'Value']
    if 'button_1' not in st.session_state:
        st.session_state.button_1 = False
    if 'button_2' not in st.session_state:
        st.session_state.button_2 = False
    def cb1():
        st.session_state.button_1 = True
    def cb2():
        st.session_state.button_2 = True
    # Define a button to display the DataFrame

    
    display_button = st.button("Display DataFrame", on_click=cb1)

with tab2:
    with st.container(height=620,border=True):
        st.markdown('<h3 class="subheader">Decision Criteria</h3>', unsafe_allow_html=True)
        st.markdown('<p class="gradient-text">All figures are in percentage</p>', unsafe_allow_html=True)
        with st.container(height=140,border=True):
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("""
                <ul class="custom-bullets">
                    <li> If Variance is BLANK - - - > FALSE</li>
                </ul>
                """, unsafe_allow_html=True)
                with col2:
                    st.markdown("""
                      <ul class="custom-bullets">
                       <li>If Variance = 0 - - - > "Same Value as Last Year"</li>
                      </ul
                      """, unsafe_allow_html=True)
                    with col3:
                        st.markdown("""
                <ul class="custom-bullets">
                    <li>variance = -100% and Current Year = 0 - - - > "Non Reporting"</li>
                </ul>
                """, unsafe_allow_html=True)
            col6, col7 = st.columns(2)
            with col6:
                st.markdown("""
            <ul class="custom-bullets">
                <li>variance = 100% and Previous Year = 0 - - - > "New Reporting"</li>
            </ul>
            """, unsafe_allow_html=True)
                with col7:
                     st.markdown("""
            <ul class="custom-bullets">
                <li>variance < -100% or variance > 100% - - - > "Abnormal Deviation"</li>
            </ul>
            """, unsafe_allow_html=True)
                         
        col4, col5 = st.columns(2)
        with col4:
            min_var_criteria = st.number_input("Lower Range", min_value=-1000000, max_value=-1, value=-100, step=1)
            with col5:
                max_var_criteria = st.number_input("Max Range", min_value=1, max_value=1000000, value=100, step=1)
        st.markdown("""
                      <ul class="custom-bullets">
                       <li>Exclude Criteria: Select the range (Note that both numbers are included in the Exclude range</li>
                      </ul
                      """, unsafe_allow_html=True)
        
        # Define the range for the slider
        default_exclude_min_val = -15
        default_exclude_max_val = 15

        # Create the two-way slider
        exclude_criteria = st.slider('Exclude Criteria:', 
                                 min_value=-30, 
                                 max_value=30, 
                                 value=(default_exclude_min_val, default_exclude_max_val), step=1)

        selected_exclude_min = exclude_criteria[0]
        selected_exclude_max = exclude_criteria[1]

        st.markdown("""
                      <ul class="custom-bullets">
                       <li>If Variance does not fall under above criteria - - - > Sample Population</li>
                      </ul
                      """, unsafe_allow_html=True)

        

    # Load and display the image
    img_2 = Image.open(r"DecisionCriteria.png")
   
    st.image(img_2,width=20, use_column_width=True)


with tab1:
    #if uploaded_file is not None and display_button:
    if st.session_state.button_1:
        # Read the Excel file into a DataFrame
        df_raw = pd.read_excel(uploaded_file)
    
        df = pd.read_excel(uploaded_file, skiprows=start_row, usecols=range(start_column, end_column + 1))
    
        # Forcefully change all column names
        df.columns = new_column_names
    
        # Display the DataFrame in Streamlit
        #st.text('Raw Environment Data')
        #st.dataframe(df)
    
        # Drop rows with missing values in the 'Raw_Column' column
        df = df.dropna(subset=['Raw_Column'])
        df.drop(df.index[-1], inplace=True)
    
        # Function to count spaces at the beginning of each entry
        def count_spaces(entry):
            return len(entry) - len(entry.lstrip())
    
        # Apply the function to the column
        df['spaces_count'] = df['Raw_Column'].apply(count_spaces)
    
    
        # Cleaning the dataframe
        # Removing the YTD values
        # Filter rows where 'Raw_Column' column contains 'YTD'
        filtered_df = df[~df['Raw_Column'].str.contains('YTD')]
    
        # Remove leading and trailing white spaces from 'Column1'
        filtered_df['Raw_Column'] = filtered_df['Raw_Column'].str.strip()
    
        filtered_df = filtered_df[~filtered_df['Raw_Column'].isin([pre_yr, curr_yr])]
    
        # Create a new column based on conditions
        filtered_df['GEA_Sites'] = np.where(filtered_df['spaces_count'] == 0, filtered_df['Raw_Column'], None)
        filtered_df['GEA_Sites'].fillna(method='ffill', inplace=True)
    
        filtered_df['Category'] = np.where(filtered_df['spaces_count'] == 2, filtered_df['Raw_Column'], None)
        filtered_df['Category'].fillna(method='ffill', inplace=True)
    
        # Removing the 'space_count' = 0 and 2
        filtered_df = filtered_df[~filtered_df['spaces_count'].isin([0,2])]
        filtered_df.drop('spaces_count', axis=1, inplace=True)
    
        # Split the "Raw_Column" into "Month" and "Year" columns
        filtered_df[['Month', 'Year']] = filtered_df['Raw_Column'].str.split(expand=True)
    
        # Map month names to their numerical values
        # month_to_number = {month: index + 1 for index, month in enumerate(calendar.month_abbr[1:])}
        month_to_number = {month: index + 1 for index, month in enumerate(calendar.month_name[1:])}
    
        # Add a new column "Month_Number" with numerical values for months
        filtered_df['Month_Number'] = filtered_df['Month'].map(month_to_number)
    
        # Pivot the 'Year' column into separate columns
        pivot_df = filtered_df.pivot_table(index=['GEA_Sites', 'Category', 'Month', 'Month_Number'], columns='Year', values='Value', aggfunc='first').reset_index()
    
        # Sort the DataFrame
        pivot_df = pivot_df.sort_values(by=['GEA_Sites', 'Category', 'Month_Number'])
        pivot_df['Variance%'] = round((pivot_df[curr_yr]-pivot_df [pre_yr]) * 100/ pivot_df[pre_yr],2)
        # pivot_df['Variance%'] = pivot_df.apply(lambda x: 
                                          # 0 if x[pre_yr] == 0 else 
                                          # round((x[curr_yr] - x[pre_yr]) * 100 / x[pre_yr], 2),
                                          # axis=1)
    
        
    
    
    
        # Function to apply the logic
        def formula(row):
            variance = row['Variance%']
            p_yr = row[pre_yr]
            c_yr = row[curr_yr]
            #if pd.isna(variance) or variance == "":
            #if pd.isna(variance):
             #   return "False"
            #else:
                # if variance == -100.0 and c_yr == 0:
                    # return "Non Reporting"
            if p_yr > 0 and pd.isnull(c_yr):
                return "Non Reporting"
            elif variance == 0.0:
                return "Same Value as Last Year"
            #elif variance == 100.0 and p_yr == 0:
            #elif np.isinf(variance):
                #return "New Reporting"
            elif (p_yr == 0 or pd.isna(p_yr) or p_yr == "") and c_yr > 0:
                return "New Reporting"
            # elif variance >= max_var_criteria or variance <= min_var_criteria:
            elif variance > max_var_criteria or variance < min_var_criteria:
                return "Abnormal Deviation"
            elif 0 < variance <= selected_exclude_max:
                return "Exclude"
            elif selected_exclude_min <= variance < 0:
                return "Exclude"
            elif min_var_criteria < variance < selected_exclude_min:
                return "Sample Population"
            elif selected_exclude_max < variance < max_var_criteria:
                return "Sample Population"
            else:
                return "False"
    
    
        # Apply the function to create a new column
        pivot_df['Decision'] = pivot_df.apply(lambda row: formula(row), axis=1)
    
        # Perform data transformation based on user input
        #pivot_df = transform_data(start_row)
    
    
    
    try:
        
        if st.session_state.button_1:
            st.subheader('Input Data:')
            #Adding Columns for Dataframes
            col4, col5 = st.columns(2)
            
            col4.text('Raw Environment Data')
            col4.dataframe(df_raw)
            
    
            col5.text('Trimmed Data')
            col5.dataframe(df)
    
            #st.text('Transformed Data')
            #st.dataframe(pivot_df)
    
    
    
            # Dropdown for Category selection
            if env_select == 'Energy':
                GEA_Sites_select = st.multiselect("Select Category", 
                                                    options=pivot_df['Category'].unique(),
                                                    default = ['Electricity' ,'Heat' ,'Natural Gas' ,'Diesel (100% mineral diesel)' ,'Diesel (retail station biofuel blend)' ,'Petrol (100% mineral petrol)' ,'Petrol (retail station biofuel blend)' ,'Burning Oil' ,'Fuel Oil' ,'Gas Oil' ,'GEA Private Jet Fuel' ,'LPG (liquified petroleum gas)' ,'Wood Pellets'])
            else:
                GEA_Sites_select = st.multiselect("Select Category", 
                                                    options=pivot_df['Category'].unique())
                                                
            final_df = pivot_df[pivot_df['Category'].isin(GEA_Sites_select)]
            final_df = final_df[final_df['GEA_Sites'] != 'GEA Group']       # Removing the GEA Group
            
            st.text('Transformed Data')
            st.dataframe(final_df)
            
            # Add a download button to download the DataFrame as an Excel file
            output_1 = io.BytesIO()
            # excel_file = pd.ExcelWriter(output_1, engine='xlsxwriter')
            # final_df.to_excel(excel_file, index=False, sheet_name='Sheet1')
            # excel_file.save()
            
            with pd.ExcelWriter(output_1, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
            output_1.seek(0)
    
    
            st.download_button(label='Download DataFrame as Excel',
                               data=output_1,
                               file_name='Full Data.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
            
    except:
        st.text("Please select the input parameters correctly and then click on the 'Display DataFrame' button to view the data.")
    
with  tab3:
    with st.container(height=250,border=True):
        # Creating a form to get the Comment file
        st.subheader('Upload Comments Excel File')
        # comment_form = st.form("Get the Comment File")

        # Comment and Attachment File Uploader
        uploaded_comment = st.file_uploader("Upload Excel File",type=["xlsx", "xls"],label_visibility="collapsed")

        # Initialize comment_df outside the condition to avoid scope issues
        # comment_df = None

        # if uploaded_comment is not None:
        # Get the list of sheet names
        xls = pd.ExcelFile(uploaded_comment)
        sheet_names = xls.sheet_names
    with st.container(height=250,border=True):
        # Display the sheet names
        selected_sheet = st.selectbox("Select Sheet", sheet_names)
        uploaded_comment_col_1, uploaded_comment_col_2, uploaded_comment_col_3 = st.columns(3)

    # Specify the starting row and columns
        uploaded_comment_start_row = uploaded_comment_col_1.number_input('Row starting number', value=0, min_value =0)  # 0-based index (row 10)
        uploaded_comment_start_column = uploaded_comment_col_2.number_input('Column starting number',step=1, value=0, min_value=0, max_value=1)  # 0-based index (column B)
        uploaded_comment_end_column = uploaded_comment_col_3.number_input('Column Ending number', step=1, value=10, min_value=0, max_value=10)  # 0-based index (column D)
    # Read the selected sheet as a DataFrame
        if uploaded_comment_start_row == 0:
            comment_df = pd.read_excel(uploaded_comment, sheet_name=selected_sheet, skiprows=uploaded_comment_start_row, usecols=range(uploaded_comment_start_column, uploaded_comment_end_column + 1))
        else:
            comment_df = pd.read_excel(uploaded_comment, sheet_name=selected_sheet, skiprows=uploaded_comment_start_row, usecols=range(uploaded_comment_start_column, uploaded_comment_end_column + 1), header=uploaded_comment_start_row)
            # Now you can work with the DataFrame 'comment_df'
            # st.dataframe(comment_df)
        # The form submission button should be inside the 'if' block
        # comment_submit_button = comment_form.form_submit_button("Submit")
        # if comment_submit_button and comment_df is not None:
    st.markdown("### Below is the Comments file preview:")
    st.dataframe(comment_df)
    final_df = pd.merge(final_df, comment_df[['Site', 'Position', 'Cycle', 'Attachments', 'Comment']], left_on=['GEA_Sites', 'Category', 'Month'], right_on=['Site', 'Position', 'Cycle'], how='left')
    final_df.drop(['Site', 'Position', 'Cycle'], axis=1, inplace=True)
    # Applying condition for Missing Attachment
    def missing_attachment(row):
        variance = row['Variance%']
        attachments = row['Attachments']
        p_yr = row[pre_yr]
        c_yr = row[curr_yr]
        if c_yr > 0 and pd.isna(attachments):
            return "Missing Attachment"
        else:
            return None
    # Apply the function to create a new column
    final_df['Missing Attachment'] = final_df.apply(lambda row: missing_attachment(row), axis=1)
    st.markdown("### Below is Final file preview with added Comments and Attachment details:")
    st.dataframe(final_df)
    # AgGrid(final_df, height = 500)
    # Add a download button to download the DataFrame as an Excel file
    output_7 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_1, engine='xlsxwriter')
    # final_df.to_excel(excel_file, index=False, sheet_name='Sheet1')
    # excel_file.save()

with pd.ExcelWriter(output_7, engine='xlsxwriter') as writer:
    final_df.to_excel(writer, index=False, sheet_name='Sheet1')
output_7.seek(0)


st.download_button(label='Download DataFrame with Comments as Excel',
                   data=output_7,
                   file_name='Full Data with Comments.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# st.dataframe(uploaded_comment)




col_year, col_month = st.columns(2)
yr_select = col_year.number_input('Year', value = current_year)

# Get the unique months from df_SP
unique_months = pivot_df['Month'].unique()

# Convert month names to their corresponding indices (1-12)
month_indices = [list(calendar.month_name).index(month) for month in unique_months]

# Sort the unique months based on their corresponding indices
sorted_unique_months = [month for _, month in sorted(zip(month_indices, unique_months))]

# Use the sorted unique months in your multiselect widget
# month_select = col_month.selectbox("Select Month", options=sorted_unique_months) 
# final_df = final_df[final_df['Month']==month_select]

month_select = col_month.multiselect("Select Month", options=sorted_unique_months) 
final_df = final_df[final_df['Month'].isin(month_select)]






tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Abnormal Deviation", "Non-Reporting", "Same Value as Last Year", "Sample Population", "New Reporting", "Missing Attachment"])
#data = np.random.randn(10, 1)

# Abnormal Deviation tab
tab1.subheader("Abnormal Deviation data")
df_AD = final_df[final_df['Decision']=='Abnormal Deviation']

if df_AD.empty:
    tab1.text("The filtered DataFrame is empty")
else:
    tab1.dataframe(df_AD)
    # if tab1.button('Download Excel'):
        # Download the DataFrame as an Excel file
        # df_AD.to_excel(r'C:\Users\das.su\dataframe.xlsx', index=False)
        # st.success('DataFrame downloaded successfully as "dataframe.xlsx"')

    # Add a download button to download the DataFrame as an Excel file
    output_2 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_2, engine='xlsxwriter')
    # df_AD.to_excel(excel_file, index=False, sheet_name='Abnormal Deviation')
    # excel_file.save()
    with pd.ExcelWriter(output_2, engine='xlsxwriter') as writer:
        df_AD.to_excel(writer, index=False, sheet_name='Sample Population')
    output_2.seek(0)

    tab1.download_button(label="Download 'Abnormal Deviation' as Excel",
                       data=output_2,
                       file_name='Abnormal Deviation.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Non-Reporting tab
tab2.subheader("Non Reporting data")
df_NR = final_df[final_df['Decision']=='Non Reporting']

if df_NR.empty:
    tab2.text("The filtered DataFrame is empty")
else:
    tab2.dataframe(df_NR)

    # Add a download button to download the DataFrame as an Excel file
    output_3 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_3, engine='xlsxwriter')
    # df_NR.to_excel(excel_file, index=False, sheet_name='Non-Reporting')
    # excel_file.save()
    with pd.ExcelWriter(output_3, engine='xlsxwriter') as writer:
        df_NR.to_excel(writer, index=False, sheet_name='Sample Population')
    output_3.seek(0)

    tab2.download_button(label="Download 'Non-Reporting' as Excel",
                       data=output_3,
                       file_name='Non-Reporting.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# New Reporting tab
tab5.subheader("New Reporting data")
df_NewR = final_df[final_df['Decision']=='New Reporting']

if df_NewR.empty:
    tab5.text("The filtered DataFrame is empty")
else:
    tab5.dataframe(df_NewR)

    # Add a download button to download the DataFrame as an Excel file
    output_3 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_3, engine='xlsxwriter')
    # df_NewR.to_excel(excel_file, index=False, sheet_name='Non-Reporting')
    # excel_file.save()
    with pd.ExcelWriter(output_3, engine='xlsxwriter') as writer:
        df_NewR.to_excel(writer, index=False, sheet_name='New Reporting')
    output_3.seek(0)

    tab5.download_button(label="Download 'New Reporting' as Excel",
                       data=output_3,
                       file_name='New Reporting.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

#Same Value as Last Year Tab
tab3.subheader("Same Value as Last Year data")
df_SVLY = final_df[final_df['Decision']=='Same Value as Last Year']

if df_SVLY.empty:
    tab3.text("The filtered DataFrame is empty")
else:
    tab3.dataframe(df_SVLY)

    # Add a download button to download the DataFrame as an Excel file
    output_4 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_4, engine='xlsxwriter')
    # df_SVLY.to_excel(excel_file, index=False, sheet_name='Same Value as Last Year')
    # excel_file.save()
    with pd.ExcelWriter(output_4, engine='xlsxwriter') as writer:
        df_SVLY.to_excel(writer, index=False, sheet_name='Sample Population')
    output_4.seek(0)

    tab3.download_button(label="Download 'Same Value as Last Year' as Excel",
                       data=output_4,
                       file_name='Same Value as Last Year.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# Missing Attachment tab
tab6.subheader("Missing Attachment data")
df_MA = final_df[final_df['Missing Attachment']=='Missing Attachment']

if df_MA.empty:
    tab6.text("The filtered DataFrame is empty")
else:
    tab6.dataframe(df_MA)

    # Add a download button to download the DataFrame as an Excel file
    output_6 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_3, engine='xlsxwriter')
    # df_NewR.to_excel(excel_file, index=False, sheet_name='Non-Reporting')
    # excel_file.save()
    with pd.ExcelWriter(output_6, engine='xlsxwriter') as writer:
        df_MA.to_excel(writer, index=False, sheet_name='Missing Attachment')
    output_6.seek(0)

    tab6.download_button(label="Download 'Missing Attachment' as Excel",
                       data=output_6,
                       file_name='Missing Attachment.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')



#Sample Population Tab
tab4.subheader("Sample Population data")
df_SP = final_df[final_df['Decision']=='Sample Population']
# df_SP.sort_values(by=['GEA_Sites', 'Category', 'Month_Number', 'Variance%'], ascending=[True, True, True, False], inplace=True)
df_SP.sort_values(by=['Category', 'Month_Number', 'Variance%'], ascending=[True, True, True], inplace=True)


if df_SP.empty:
    tab4.text("The filtered DataFrame is empty")
else:
    # tab4.dataframe(df_SP)
    with tab4:
        AgGrid(df_SP, height = 500)
    # Add a download button to download the DataFrame as an Excel file
    output_5 = io.BytesIO()
    # excel_file = pd.ExcelWriter(output_5, engine='xlsxwriter')
    # df_SP.to_excel(excel_file, index=False, sheet_name='Sample Population')
    # excel_file.save()
    with pd.ExcelWriter(output_5, engine='xlsxwriter') as writer:
        df_SP.to_excel(writer, index=False, sheet_name='Sample Population')
    output_5.seek(0)

    tab4.download_button(label="Download 'Sample Population' as Excel",
                       data=output_5,
                       file_name='Sample Population.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                       

    
    #Creating 3 columns in tab4
    col_t4_1, col_t4_2, col_t4_3, col_t4_4, col_t4_5 = tab4.columns(5)
    
    # Selecting the month
    # df_SP = df_SP[df_SP['Month']==month_select]
    df_SP = df_SP[df_SP['Month'].isin(month_select)]
    
    # Selection
    col_t4_1.markdown('<div style="text-align: right; font-size: 20px; font-weight: bold;">Total Population Count : </div>', unsafe_allow_html=True)
    gea_sites = col_t4_2.number_input('No. of GEA Sites', value = final_df['GEA_Sites'].nunique())
    n_category = col_t4_3.number_input('No. of Categories', value = final_df['Category'].nunique())
    data_points = col_t4_4.number_input('Max Data Points', value = int(gea_sites) * int(n_category))
    sample_size = col_t4_5.number_input('Max Sample Size', value = round(data_points*6/100,0), format="%.0f")
    
    col_t4_6, col_t4_7, col_t4_8, col_t4_9, col_t4_10 = tab4.columns(5)
    # tab4.text("Actual Count")
    col_t4_6.markdown('<div style="text-align: right; font-size: 20px; font-weight: bold;">Sample Population Count : </div>', unsafe_allow_html=True)
    actual_gea_sites = col_t4_7.number_input('Max Sample No. of GEA Sites', value = df_SP['GEA_Sites'].nunique())
    actual_n_category = col_t4_8.number_input('Max Sample No. of Categories', value = df_SP['Category'].nunique())
    # actual_n_category = col_t4_8.number_input('Actual No. of Categories', value = df_SP['Category'].count())
    actual_data_points = col_t4_9.number_input('Max Sample Data Points', value = int(actual_gea_sites) * int(actual_n_category))
    actual_sample_size = col_t4_10.number_input('Max Sample Size from Sample Population', value = round(actual_data_points*6/100,0), format="%.0f")

    col_t4_11, col_t4_12, col_t4_13, col_t4_14, col_t4_15 = tab4.columns(5)
    col_t4_11.markdown('<div style="text-align: right; font-size: 20px; font-weight: bold;">Actual Sample Population Count : </div>', unsafe_allow_html=True)
    sample_data_points = col_t4_14.number_input('Actual Sample Data Points', value = df_SP['Category'].count())
    sample_sample_size = col_t4_15.number_input('Actual Sample Size from Sample Population', value = round(sample_data_points*6/100,0), format="%.0f")
    
    manual_sample = col_t4_15.number_input('Manual Sample 2 to pick:', value = round(sample_data_points*6/100,0), format="%.0f")
    
    # tab4.text(sample_size)
    
    # Calculate the sum of yr_select grouped by 'Category'
    # df_SVLY_category_summary = df_SP.groupby('Category')[str(yr_select)].sum()
    # Calculate the sum and count by 'Category'
    # df_SVLY_category_summary = df_SP[df_SP['Month']==month_select]
    df_SVLY_category_summary = df_SP.groupby('Category').agg({str(yr_select): 'sum', 'Category': 'size'})
    # df_SVLY_category_summary = df_SVLY_category_summary.applymap(lambda x: '{:,.2f}'.format(x))
    # df_SVLY_category_summary[str(yr_select)] = df_SVLY_category_summary[str(yr_select)].apply(lambda x: '{:,.2f}'.format(x))
    
    # Apply formatting to the sum column
    # df_SVLY_category_summary[str(yr_select)] = df_SVLY_category_summary[str(yr_select)].apply(lambda x: '{:,.2f}'.format(round(x, 2)))
    # Rename the columns
    df_SVLY_category_summary.columns = [str(yr_select), 'Actual Count']
    
    df_SVLY_category_summary_1 = df_SVLY_category_summary.reset_index()

    df_SVLY_category_summary_1[str(yr_select)] = round(df_SVLY_category_summary_1[str(yr_select)],2)
    df_SVLY_category_summary_1['Percentage'] = round(df_SVLY_category_summary_1[str(yr_select)]*100/df_SVLY_category_summary_1[str(yr_select)].sum(),2)
    # df_SVLY_category_summary_1['Percentage'] = round(df_SVLY_category_summary_1[str(yr_select)]*100/df_SVLY_category_summary_1[str(yr_select)].sum(),2)
    
    df_SVLY_category_summary_1['Max Data Points'] = round(df_SVLY_category_summary_1['Percentage']*data_points/100,0)
    df_SVLY_category_summary_1['Max Sample Data Points'] = round(df_SVLY_category_summary_1['Percentage']*actual_data_points/100,0)
    

    # df_SVLY_category_summary_1['Auto Sample'] = df_SVLY_category_summary_1[str(yr_select)].apply(lambda x: 1 if round(x * 0.06, 0) == 0 else round(x * 0.06, 0))
    
    df_SVLY_category_summary_1['Actual Sample'] = np.where(round(df_SVLY_category_summary_1['Actual Count'] * 0.06) == 0, 
                                                                0, 
                                                                round(df_SVLY_category_summary_1['Actual Count'] * 0.06))
    
    
    df_SVLY_category_summary_1['Max Sample'] = np.where(round(df_SVLY_category_summary_1['Max Data Points'] * 0.06) == 0, 
                                                                0, 
                                                                round(df_SVLY_category_summary_1['Max Data Points'] * 0.06))
                                                                
    df_SVLY_category_summary_1['Sample'] = np.where(round(df_SVLY_category_summary_1['Max Sample Data Points'] * 0.06) == 0, 
                                                            0, 
                                                            round(df_SVLY_category_summary_1['Max Sample Data Points'] * 0.06))
    
    
    df_SVLY_category_summary_1['Manual Sample'] = df_SVLY_category_summary_1['Actual Sample']
    df_SVLY_category_summary_1['Manual Sample 2'] = round(df_SVLY_category_summary_1['Percentage']*manual_sample/100,0)
    # df_SVLY_category_summary_1['Fixed Sample'] = [4,4,12,2,2,2,14,4,4,2]
    # tab4.dataframe(df_SVLY_category_summary_1)
    
    # df_SVLY_category_summary_1['Under Report'] = ''
    # df_SVLY_category_summary_1['Over Report'] = ''
    
    # To enable the entry for Manual Sample column, other column remains freezed
    # create a grid options builder object
    gb = GridOptionsBuilder.from_dataframe(df_SVLY_category_summary_1)

    # configure the second column to be editable
    gb.configure_column('Manual Sample', editable=True)
    gb.configure_column('Manual Sample 2', editable=True)
    # gb.configure_column('Under Report', editable=True)
    # gb.configure_column('Over Report', editable=True)
    
    # configure multiple columns at once
    # gb.configure_columns({
        # 'Manual Sample': {'editable': True},
        # 'Under Report': {'editable': True},
        # 'Over Report': {'editable': True}
    # })

    # Calculate the sum of each column
    totals = df_SVLY_category_summary_1.sum(numeric_only=True)
    totals_df = pd.DataFrame(totals, columns=['Total'])  
    totals_df = totals_df.T
    
    # print(totals)
    # print(totals_df.reset_index())
    # print(df_SVLY_category_summary_1)
    
    # Append the totals to the bottom of the DataFrame
    # df_SVLY_category_summary_1 = df_SVLY_category_summary_1.append(totals, ignore_index=True)
    df_SVLY_category_summary_1 = pd.concat([df_SVLY_category_summary_1, totals_df], axis=0)

    # You might want to fill the non-numeric 'Category' field for the total row
    df_SVLY_category_summary_1.at[df_SVLY_category_summary_1.index[-1], 'Category'] = 'Total'


    # build the grid options
    gridOptions = gb.build()

    # display the aggrid dataframe with the grid options
    st.write('<style>.css-vxshay .stDataFrame {width: 100% !important;}</style>', unsafe_allow_html=True)
    with tab4:
        summary_data = AgGrid(df_SVLY_category_summary_1, gridOptions=gridOptions, width='100%')
        # AgGrid(df_SVLY_category_summary_1, gridOptions=gridOptions, columns_auto_size_mode=ColumnsAutoSizeMode.FIT_ALL_COLUMNS_TO_VIEW)
        
    # Extract DataFrame from summary_data
    summary_data = summary_data['data']
    
    # tab4.dataframe(summary_data)
    # Picking the Actual Sample
    tab4.subheader("Sampled data")
    # tab4.text('Select the column from where you need to pick the data sample:')
    column_select = tab4.selectbox('Select the column from where you need to pick the data sample:', ['Actual Sample', 'Manual Sample', 'Manual Sample 2', 'Sample'], index=0)
    # column_select = tab4.selectbox('Select the column from where you need to pick the data sample:', ['Actual Sample', 'Manual Sample', 'Manual Sample 2', 'Sample', 'Fixed Sample'], index=0)
        
    population_df = df_SP
    # sample_numbers_df = df_SVLY_category_summary_1[['Category', 'Actual Sample']][:-1]
    # sample_numbers_df = summary_data[['Category', column_select]][:-1]
    sample_numbers_df = summary_data[['Category', column_select]][:-1]
    # tab4.dataframe(sample_numbers_df)
    # sample_numbers_df['Actual Sample'] = sample_numbers_df['Actual Sample'].astype(int)
    sample_numbers_df['Target Sample'] = sample_numbers_df[column_select].astype(int)
    sample_numbers_df = sample_numbers_df.drop(columns=[column_select])

    # tab4.dataframe(sample_numbers_df)
    # tab4.text(sample_numbers_df.info)

    # Create an empty dataframe to store the samples
    sampled_df = pd.DataFrame(columns=population_df.columns)
    print("SD_1", sampled_df)
    # tab4.dataframe(sampled_df)

    # Iterate over each group in sample_numbers_df
    # for group, num_samples in zip(sample_numbers_df['Category'], sample_numbers_df['Actual Sample']):
        # Sample from the group based on the specified number of samples
        # sampled_group = population_df[population_df['Category'] == group].sample(n=num_samples, replace=True)
        
        # Append the sampled group to the sampled dataframe
        # sampled_df = sampled_df.append(sampled_group)

    # Iterate over each group in sample_numbers_df
    for _, row in sample_numbers_df.iterrows():
        group = row['Category']
        num_samples = row['Target Sample']
        
        # Filter population dataframe by group
        group_population = population_df[population_df['Category'] == group]
        
        # Sample from the group based on the specified number of samples
        # sampled_group = group_population.sample(n=num_samples, replace=False)
        if column_select == 'Manual Sample' or  column_select == 'Manual Sample 2':
            # Round up for odd number of samples
            half_num_samples = math.ceil(num_samples / 2)
            sampled_group = pd.concat([group_population.head(half_num_samples), group_population.tail(half_num_samples)])
        else:
            sampled_group = group_population.sample(n=num_samples, replace=False)
        
        # Append the sampled group to the list
        # sampled_df = sampled_df.append(sampled_group)
        sampled_df = pd.concat([sampled_df, sampled_group], axis=0)

    print("SD_2", sampled_group)
    
    # Reset index of the sampled dataframe
    sampled_df.reset_index(drop=True, inplace=True)

    # Reset index of the sampled dataframe

    sampled_df.reset_index(drop=True, inplace=True)
    # tab4.dataframe(sampled_df)
    with tab4:
        final_sample = AgGrid(sampled_df, index=True)
    
    final_sample = final_sample['data']
    
    # Add a download button to download the DataFrame as an Excel file
    output_6 = io.BytesIO()
    with pd.ExcelWriter(output_6, engine='xlsxwriter') as writer:
        final_sample.to_excel(writer, index=False, sheet_name='Final Sample')
    output_6.seek(0)

    tab4.download_button(label="Download 'Final Sample' as Excel",
                       data=output_6,
                       file_name='Final Sample.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
