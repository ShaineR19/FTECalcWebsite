"""
Streamlit App – Faculty FTE Report Generator

This application reads course enrollment and FTE data, then allows users
to generate customized reports and visualizations for:

- Divisions
- Individual Courses
- Instructor Performance
- Section Enrollment Ratios

Outputs can be previewed, graphed, and downloaded as Excel files.

Modules:
--------
- `web_functions` (wf): contains data preprocessing and report logic
- `options4` (opfour): utility functions for formatting and cleaning
"""

# -*- coding: utf-8 -*-
import io
import streamlit as st
import pandas as pd
import web_functions as wf
import options4 as opfour
import seaborn as sns
import matplotlib.pyplot as plt


@st.cache_data
def load_data():
    """
    Loads and caches core datasets for the app.

    Returns
    -------
    tuple(pd.DataFrame, pd.DataFrame, pd.DataFrame)
        - dean_df: Full dataset
        - unique_df: Dataset with unique course sections
        - fte_tier: Tier mapping for FTE generation
    """

    dean_df = wf.readfile()
    unique_df = pd.read_excel('unique_deansDailyCsar_FTE.xlsx')
    fte_tier = pd.read_excel('FTE_Tier.xlsx')
    dean_df.columns = dean_df.columns.str.strip()
    unique_df.columns = unique_df.columns.str.strip()
    return dean_df, unique_df, fte_tier


dean_df, unique_df, fte_tier = load_data()


def save_report(df_full, filename, image=None):
    """
    Prompts the user to name and download an Excel report.

    Parameters
    ----------
    df_full : pd.DataFrame
        The DataFrame to export.
    fig : png
        png of plot to export
    filename : str
        Suggested default filename for the export.
    """

    user_filename = st.text_input("Enter a filename (e.g., my_report.xlsx):",
                                  value=filename)

    if filename:

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_full.to_excel(writer, sheet_name='Full Report', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Full Report']

            # Auto-size each column based on content
            for i, column in enumerate(df_full.columns):
                # Get max length of data and column name
                col_len = max(
                    df_full[column].astype(str).map(len).max(),
                    len(column)
                )
                worksheet.set_column(i, i, col_len + 2)  # +2 for padding

            if image:
                workbook = writer.book
                chart_sheet = workbook.add_worksheet("Graph Report")
                writer.sheets["Graph Report"] = chart_sheet
                chart_sheet.insert_image('A1', image)

        st.download_button("Save Report", data=output.getvalue(),
                           file_name=user_filename)


# We are generating the sidebar here
menu = [
    "Home",
    "Sec Division Report",
    "Course Enrollment Percentage",
    "FTE by Division",
    "FTE per Instructor",
    "FTE per Course"
]

st.sidebar.title("Navigation")
choice = st.sidebar.selectbox("Choose Report Option", menu)

# Title that will stay at the top of the program
st.title('FTE Report Generator')

# Branches for the sidebar menu
if choice == "Sec Division Report":
    st.header("Sec Division Report")
    if 'Sec Divisions' in dean_df.columns:
        division = st.selectbox("Select Division",
                                dean_df['Sec Divisions'].dropna().unique())

        run = st.button("Run Report")
        if run:
            filtered = dean_df[dean_df['Sec Divisions'] == division]
            filtered.index = range(1, len(filtered) + 1)

            st.dataframe(filtered.head(10))
            save_report(filtered, f"{division}_Division_Report.xlsx")
    else:
        st.warning("This feature will run when 'Sec Divisions' is available in the dataset.")

elif choice == "Course Enrollment Percentage":
    st.header("Course Enrollment Percentage")
    if 'Sec Name' in dean_df.columns:
        course = st.selectbox("Select Course",
                              dean_df['Course Code'].dropna().unique())
        run = st.button("Run Report")
        if run:
            filtered = dean_df[dean_df['Course Code'] == course]
            filtered = filtered.drop_duplicates(subset="Sec Name")
            filtered["Enrollment Percentage"] = filtered.apply(
                wf.calc_enrollment, axis=1)

            st.dataframe(filtered.head(10))
            save_report(filtered, f"{course}_Course_Report.xlsx")
    else:
        st.warning("This feature will run when 'Sec Name' is available in the dataset.")

elif choice == "FTE by Division":
    st.header("FTE by Division")

    if 'Sec Divisions' in dean_df.columns:
        division_input = st.selectbox(
            "Select Division", dean_df['Sec Divisions'].dropna().unique())

        run = st.button("Run Report")

        if run:
            raw_df, orig_total, gen_total = wf.fte_by_div_raw(dean_df, fte_tier, division_input)

            report_df = wf.format_fte_output(raw_df, orig_total, gen_total)

            # Format Dataframe
            format_df = report_df[~report_df['Course Code'].isin(['Total', 'DIVISION TOTAL'])].copy()
            format_df = format_df.iloc[:, 2:]

            # add generated fte float at end(remove money sign, commas)
            format_df['Generated FTE Float'] = format_df['Generated FTE']\
            .str.replace('[\$,]', '', regex=True)\
            .astype(float)

            # sort by gen fte float
            format_df = format_df.sort_values(by='Generated FTE Float', ascending=False)

            # set index
            format_df.index = range(1, len(format_df) + 1)

            # make a copy for plot
            plot_df = format_df

            # drop gen fte float
            format_df = format_df.iloc[:, :-1]

            # Display Dataframe
            st.dataframe(format_df.head(10))

            # Create Plot
            fig, ax = plt.subplots()
            sns.barplot(data=plot_df.head(10), x='Sec Name', y='Generated FTE Float', ax=ax)
            ax.set_title(f"Top 10 Sections by Total FTE in {division_input}")
            ax.set_xlabel("Section Name")
            ax.set_ylabel("Generated FTE ($)")
            ax.tick_params(axis='x', rotation=45)
            
            # Display Plot
            st.pyplot(fig)

            # Save Plot as a png
            img_bytes = io.BytesIO()
            fig.savefig(img_bytes, format='png', bbox_inches='tight')
            img_bytes.seek(0)

            # Save button
            save_report(report_df, f"{division_input}_FTE_Report.xlsx", image = img_bytes)

            st.info(f"Total FTE: {orig_total:.3f}")
            st.info(f"Generated FTE: ${gen_total:,.2f}")
    else:
        st.info("Division data not available.")

elif choice == "FTE per Instructor":
    st.header("FTE per Instructor")
    if 'Sec All Faculty Last Names' in dean_df.columns:
        faculty_list = sorted(dean_df["Sec Faculty Info"].dropna().unique())
        instructor = st.selectbox("Select Instructor", faculty_list)

        run = st.button("Run Report")
        if run:
            report_df, orig_fte, gen_fte = wf.generate_faculty_fte_report(
                dean_df, fte_tier, instructor)

            report_df = report_df.fillna("")

            if report_df is not None:
                # Format dataframe
                report_df.index = range(1, len(report_df) + 1)

                # Display dataframe
                st.dataframe(report_df)

                # Format dataframe for plot
                report_df = report_df.sort_values(by='Total FTE',
                                                  ascending=False)

                # Display Plot
                st.bar_chart(report_df
                             [report_df['Sec Name'] != 'TOTAL']
                             .set_index('Sec Name')['Total FTE'])

                # Save Report
                filename = opfour.clean_instructor_name(instructor)
                save_report(report_df, filename)

                st.info(f"Total FTE: {orig_fte:.3f}")
                st.info(f"Generated FTE: ${gen_fte:,.2f}")

            else:
                st.warning("Select an Instructor.")
    else:
        st.warning("Instructor name column missing.")

elif choice == "FTE per Course":
    st.header("FTE per Course")
    if 'Sec Name' in unique_df.columns:
        course_name = st.text_input("Enter Course Name (Sec Name)")
        run = st.button("Run Report")
        if run:
            df_result, original_fte, generated_fte = wf.calculate_fte_by_course(dean_df, fte_tier, course_name)

            if df_result is not None:
                # Format dataframe
                df_result.index = range(1, len(df_result) + 1)

                # Display dataframe
                st.dataframe(df_result.head(10))

                # Format dataframe
                df_result = df_result.sort_values(by='Total FTE',
                                                  ascending=False)

                # Plot the dataframe
                st.bar_chart(df_result
                             [df_result['Sec Name'] != 'COURSE TOTAL']
                             .set_index('Sec Name')['Total FTE'])

                # Save Report button
                save_report(df_result, f"{course_name}_FTE_Report.xlsx")

                # Display Info to user
                st.info(f"Original Total FTE: {original_fte:.3f}")
                st.info(f"Generated FTE: ${generated_fte:,.2f}")
            else:
                st.warning("Course not found.")
    else:
        st.warning("This feature will run when 'Sec Name' is present in the FTE dataset.")

elif choice == "Home":

    st.title("📘 What is Generated FTE?")
    st.markdown("""
    **FTE (Full-Time Equivalent)** credits are produced by student enrollment in specific courses or programs.

    Because not all students are full-time, this allows the school to add up all the time students spend in class to calculate how many full-time students that would equal.
    """)

    st.markdown("---")
    st.header("🛠 Description of Each Option")
    st.markdown("""
    - **By Division**: Generate a report showing FTE by department.
    - **By Course Section**: View enrollment percentage and FTE for a specific class section.
    - **Division Summary**: Summary of FTE for all courses within a division.
    - **Instructor View**: FTE generated by each course an instructor teaches.
    - **By Course**: FTE report for all sections under a single course.
    """)
    st.markdown("---")
    st.header("📁 Upload Enrollment File")
    uploaded_file = st.file_uploader("Choose a file", type=["xlsx", "csv"])
    if uploaded_file is not None:
        st.success(f"Uploaded file: {uploaded_file.name}")
