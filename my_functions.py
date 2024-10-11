from IPython.display import display, HTML
import pandas as pd
import numpy as np
from datetime import datetime
import seaborn as sns
import matplotlib.pyplot as plt
import math

# Define a function to display a DataFrame with a scrollbar and locked headers
def display_scrollable_dataframe(df, max_rows=100, max_height=700):
    # Limit the DataFrame to the specified number of rows
    df_limited = df.head(max_rows)
    
    style = f"""
    <style>
    .dataframe-div {{
        max-height: {max_height}px;
        overflow: auto;
    }}
    .dataframe thead th {{
        position: sticky;
        top: 0;
        background-color: white;
        z-index: 1;
    }}
    </style>
    """
    html_content = df_limited.to_html(classes='dataframe-div')
    pd.options.display.float_format = '{:,.2f}'.format
    display(HTML(style + '<div class="dataframe-div">' + html_content + '</div>'))

def export_pop_sample_tables(prog_name, sample_name, raw_data, project_list, meas2proj_dict):

    # Get the current date and format it as YYYY_MM_DD
    current_date = datetime.now().strftime("%Y_%m_%d")

    # Create the filename with the dynamic date
    file_name = f"{prog_name}_{sample_name}_{current_date}.xlsx"

    # Write DataFrame to Excel
    with pd.ExcelWriter(file_name) as writer:
        raw_data.to_excel(writer, sheet_name='RawData', index=True)
        project_list.to_excel(writer, sheet_name='PopByProject', index=True)
        meas2proj_dict.to_excel(writer, sheet_name='MeasToProj_Dict', index=True)



def import_tabs_from_excel(directory_name, survey_files):
    # Initialize an empty dictionary to hold DataFrames
    dfs = {}
    
    # For-loop to iterate through survey_info
    for survey_name, sheet_name, skip_row, query, df_name in survey_files:

        df = pd.read_excel(directory_name + '/' + survey_name, sheet_name=sheet_name, skiprows=skip_row)
        
        # Apply the query if it's provided (i.e., not None or an empty string)
        if query != '':
            df = df.query(query)
        
        # Assign row ID based on order of compiling (optional)
        df['ObservationID'] = df.index
        
        # Store the DataFrame in the dictionary with the sheet name as the key
        if df_name != '':
            dfs[df_name] = df
        else:
            dfs[sheet_name] = df

    return dfs


def identify_outliers(df, threshold=3):
    outliers_zscore = identify_outliers_zscore(df)
    
    # Filter to show only rows with at least one True value
    rows_with_outliers = outliers_zscore.any(axis=1)
    outliers_only = df[rows_with_outliers]
    
    # Count the number of outliers by field
    outliers_count = outliers_zscore.sum()
    
    print("\nNumber of outliers by field:")
    print(outliers_count)
    print(f"\nTotal outlier count: {outliers_count.sum()}")
    print(f"Unique outlier row count: {len(rows_with_outliers[rows_with_outliers == True])}")

    return outliers_zscore, rows_with_outliers, outliers_only, outliers_count


def identify_outliers_zscore(df, threshold=3):
    outliers = pd.DataFrame(index=df.index)
    for col in df.select_dtypes(include=[np.number]).columns:
        z_scores = (df[col].replace([np.inf, -np.inf], np.nan) - df[col].replace([np.inf, -np.inf], np.nan).mean()) / df[col].replace([np.inf, -np.inf], np.nan).std(skipna=True)
        outliers[col] = np.abs(z_scores) > threshold
    return outliers


def boxplot_numerical_fields(df_numcols, max_plots_per_row = 5):
    # Parameters for layout
    num_cols = df_numcols.select_dtypes(include=[np.number]).shape[1]
    num_rows = (num_cols // max_plots_per_row) + int(num_cols % max_plots_per_row > 0)
    
    # Create subplots
    fig, axes = plt.subplots(num_rows, max_plots_per_row, figsize=(5 * max_plots_per_row, 6 * num_rows))
    
    # Flatten the axes array for easy iteration
    axes = axes.flatten()
    
    # Create individual box plots for each numerical field
    for i, col in enumerate(df_numcols.select_dtypes(include=[np.number]).columns):
        sns.boxplot(y=df_numcols[col], ax=axes[i])
        axes[i].set_title(f'Box Plot of {col}')
        axes[i].set_xlabel('')  # Remove x-label to avoid clutter
        axes[i].set_ylabel('Values')
    
    # Remove any unused subplots
    for j in range(i + 1, len(axes)):
        fig.delaxes(axes[j])
    
    # Adjust layout
    plt.tight_layout()
    plt.show()


def field_exists_in_other_field(field1, field2, df1, df2):
    # Check if the project IDs exist in the 'Project List Project ID' column
    exists_df2_in_df1 = field2.isin(field1)
    exists_df1_in_df2 = field1.isin(field2)

    # Count the number of rows for each condition
    count_df1_in_df2 = df1[exists_df1_in_df2].shape[0]
    count_df1_not_in_df2 = df1[~exists_df1_in_df2].shape[0]
    count_df2_in_df1 = df2[exists_df2_in_df1].shape[0]
    count_df2_not_in_df1 = df2[~exists_df2_in_df1].shape[0]

    # Create a summary table
    merge_summary = pd.DataFrame({
        'Condition': ['df1 in df2', 'df1 not in df2',
                      'df2 in df1', 'df2 not in df1'],
        'Count': [count_df1_in_df2, count_df1_not_in_df2,
                  count_df2_in_df1, count_df2_not_in_df1]
    })
    
    # Reshape the DataFrame to have two rows and two columns
    merge_summary = merge_summary.pivot_table(index='Condition', values='Count')
    
    # Display the summary table
    print("Merge Summary:")
    display_scrollable_dataframe(merge_summary)
    
    # Filter the DataFrame to show only matching rows
    print("\nRows from df1 in df2: \n")
    display_scrollable_dataframe(df1[exists_df1_in_df2])
    print("\nRows from df1 not in df2: \n")
    display_scrollable_dataframe(df1[~exists_df1_in_df2])
    print("\ndf2 in df1: \n")
    display_scrollable_dataframe(df2[exists_df2_in_df1])
    print("\ndf2 not in df1: \n")
    display_scrollable_dataframe(df2[~exists_df2_in_df1])


# Define a function to categorize based on keywords
def categorize_measure(description):
    if 'lighting' in description.lower():
        return 'Lighting'
    elif 'foodsvc' in description.lower():
        return 'Food Service'
    elif 'led' in description.lower():
        return 'Lighting'
    elif 'water savings' in description.lower():
        return 'Water Heater'
    elif 'energy star' in description.lower():
        return 'Energy Star Appliances'
    elif 'energystar' in description.lower():
        return 'Energy Star Appliances'
    elif 'heat pump' in description.lower():
        return 'HVAC'
    elif 'hvac' in description.lower():
        return 'HVAC'
    elif 'insulation' in description.lower():
        return 'Insulation'
    elif 'windows' in description.lower():
        return 'Windows'
    elif 'custom' in description.lower():
        return 'Custom'
    elif 'refrigerant' in description.lower():
        return 'Refrigeration'
    elif 'refrigeration' in description.lower():
        return 'Refrigeration'
    elif 'batter' in description.lower():
        return 'Batteries'
    else:
        return 'Other'

