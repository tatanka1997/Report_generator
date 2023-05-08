import pandas as pd
import plotly.express as px
import streamlit as st
import weasyprint
import base64


# Define a function to generate the PDF report
def generate_pdf_report(df, salary, deductions, projects, check_date):
    df = df[df['Check Date'] == check_date]  # filter data by check_date
    job_name = ', '.join(df['Job Name'].unique())
    last_name = ', '.join(df['Last Name and Suffix'].unique())
    df = df.fillna(0)  # replace NaN values with 0
    html = f"<h1>Paychex Report ({check_date})</h1><p>Salary: {salary}</p><p>Deductions: {deductions}</p><p>Projects: {projects}</p><p>Employees: {last_name}</p>"
    html += "<table style='border-collapse: collapse; font-size: 12px;'>"
    html += "<tr style='background-color: #f2f2f2;'>"
    for col in df.columns:
        html += f"<th style='border: 1px solid black; padding: 5px;'>{col}</th>"
    html += "</tr>"
    for i, row in df.iterrows():
        html += "<tr style='border: 1px solid black; page-break-inside: avoid;'>"
        for val in row:
            html += f"<td style='border: 1px solid black; padding: 5px;'>{val}</td>"
        html += "</tr>"
    html += "</table>"
    report_pdf = weasyprint.HTML(string=html).write_pdf()
    b64 = base64.b64encode(report_pdf).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="paychex_report_{check_date}.pdf">Download PDF Report ({check_date})</a>'
    return href


st.set_page_config(page_title="Paychex Data",
                   page_icon=":bar_chart:",
                   layout="wide")

uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
if uploaded_files:
    # Combine the dataframes and sort by job name
    df_combined = pd.concat(
        [pd.read_excel(file, engine='openpyxl', sheet_name='Paychex Data') for file in uploaded_files],
        ignore_index=True, sort=False)
    df_combined.loc[(df_combined['Job Name'] == 'Unassigned') &
                    (df_combined['Primary Org Unit'] == '3 Maintenance'), 'Job Name'] = 'Misc'

    # Replace 'Unassigned' Job Name and non-'1 Administrative' Primary Org Unit with 'Misc'
    df_combined.loc[(df_combined['Job Name'] == 'Unassigned') &
                    (df_combined['Primary Org Unit'] != '3 Maintenance'), 'Job Name'] = 'Office'
    df_combined = df_combined.sort_values(by=['Job Name'])

    st.sidebar.header("Please Filter Here:")

    filters = {
        "Job Name": "ALL",
        "Last Name and Suffix": "ALL",
        "Check Date": "ALL",
    }

    for key, values in filters.items():
        options = ["ALL"] + list(df_combined[key].unique())
        default = values
        filter_values = st.sidebar.multiselect(
            f"Select the {key}:",
            options=options,
            default=default,
        )
        filters[key] = filter_values

    df_selection = df_combined
    for key, values in filters.items():
        if "ALL" not in values:
            df_selection = df_selection[df_selection[key].isin(values)]

    st.dataframe(df_selection.set_index('Job Name'), height=500)

    if st.button('Create Report'):
        # Get unique check dates
        keywords = ['PX401 ', 'Child ']
        keywords1 = ['Union', 'misc']
        check_dates = df_selection['Check Date'].unique()
        for check_date in check_dates:
            # Filter data by check_date
            df_by_date = df_selection[df_selection['Check Date'] == check_date]
            df_by_date['tac'] = 0
            df_by_date['misc'] = 0
            for keyword in keywords:
                df_selected = df_by_date[
                    df_by_date['Withholding-Deduction Name'].str.contains(keyword, na=False)]
                df_by_date.loc[df_selected.index, 'Child Support/401k'] = df_selected['Withholding-Deduction Amt']
            for keyword1 in keywords1:
                df_selected = df_by_date[
                    df_by_date['Withholding-Deduction Name'].str.contains(keyword1, na=False)]
                df_by_date.loc[df_selected.index, 'Other Ded'] = df_selected['Withholding-Deduction Amt']


            # Group data by Check Date and Job Name and calculate sum of certain columns
            group_by_columns = ['Job Name']
            df_grouped = df_by_date.groupby(group_by_columns).agg(
                {'Earning Amount': 'sum', 'Reimbursement-Other Payment Amount': 'sum', 'Withholding-Deduction Amt': 'sum', 'Child Support/401k': 'sum','Combined Company and Employee Tax Amount':'sum', 'Other Ded': 'sum'})
            df_grouped['Salary'] = df_grouped['Earning Amount'] + df_grouped['Reimbursement-Other Payment Amount'] - df_grouped['Other Ded']
            df_grouped['Tax'] = df_grouped['Combined Company and Employee Tax Amount'] - df_grouped['Withholding-Deduction Amt'] + df_grouped['Other Ded'] + df_grouped['Child Support/401k']
            df_grouped = df_grouped.drop(['Earning Amount', 'Reimbursement-Other Payment Amount','Withholding-Deduction Amt','Child Support/401k','Combined Company and Employee Tax Amount','Other Ded'], axis=1)
            salary_total = df_grouped['Salary'].sum()
            tax_total = df_grouped['Tax'].sum()
            df_grouped.loc['Total'] = pd.Series({'Salary': salary_total, 'Tax': tax_total})
            st.write(f"Salary and Deduction by Check Date and Job Name for {check_date}:")
            st.write(df_grouped)
            filename = f"QB Recap {check_date}.xlsx"  # create filename using check date
            df_grouped.to_excel(filename) # save dataframe to excel file with the created filename
            st.write('File was successfuly generated in the folder where you imported the Paychex Data file')

