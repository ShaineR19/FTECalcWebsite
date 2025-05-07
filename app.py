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
import time
import io
import streamlit as st
import pandas as pd
import web_functions as wf
import options4 as opfour
import seaborn as sns
import matplotlib.pyplot as plt
import xlsxwriter

def save_faculty_excel(data, instructor_name):
    output = io.BytesIO()

    # Clean numeric columns
    numeric_columns = ["Capacity", "FTE Count", "Total FTE", "Generated FTE"]
    for col in numeric_columns:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors='coerce')

    headers = [
        "Instructor", "Course Code", "Sec Name", "X Sec Delivery Method",
        "Meeting Times", "Capacity", "FTE Count", "Total FTE",
        "Sec Divisions", "Generated FTE"
    ]

    with xlsxwriter.Workbook(output, {'nan_inf_to_errors': True}) as workbook:
        worksheet = workbook.add_worksheet("Faculty Report")

        # === Formats ===
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        number_format = workbook.add_format({"num_format": "#,##0.00"})
        total_format = workbook.add_format({"bold": True, "bg_color": "#E0E0E0", "border": 1})
        total_money = workbook.add_format({"bold": True, "bg_color": "#E0E0E0", "num_format": "$#,##0.00", "border": 1})

        # === Write headers in row 0 ===
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header, header_format)

        current_row = 1
        grand_total_fte = 0
        grand_total_gen_fte = 0

        grouped = data[data["Course Code"] != "TOTAL"].groupby("Course Code")

        for course_code, group in grouped:
            course_total_fte = 0
            course_total_gen_fte = 0

            for i, (_, row) in enumerate(group.iterrows()):
                worksheet.write(current_row, 0, instructor_name if current_row == 1 else "")
                worksheet.write(current_row, 1, course_code)
                worksheet.write(current_row, 2, row.get("Sec Name", ""))
                worksheet.write(current_row, 3, row.get("X Sec Delivery Method", ""))
                worksheet.write(current_row, 4, row.get("Meeting Times", ""))
                worksheet.write(current_row, 5, row.get("Capacity", ""))
                worksheet.write(current_row, 6, row.get("FTE Count", ""))
                if pd.notna(row.get("Total FTE")):
                    worksheet.write_number(current_row, 7, row["Total FTE"], number_format)
                else:
                    worksheet.write(current_row, 7, "", number_format)
                worksheet.write(current_row, 8, row.get("Sec Divisions", ""))
                if pd.notna(row.get("Generated FTE")):
                    worksheet.write_number(current_row, 9, row["Generated FTE"], money_format)
                else:
                    worksheet.write(current_row, 9, "", money_format)

                course_total_fte += row["Total FTE"] if pd.notna(row.get("Total FTE")) else 0
                course_total_gen_fte += row["Generated FTE"] if pd.notna(row.get("Generated FTE")) else 0
                current_row += 1

            # Subtotal row
            # Subtotal row (no shading)
            worksheet.write(current_row, 1, "Total")
            if pd.notna(course_total_fte):
                worksheet.write_number(current_row, 7, course_total_fte)
            if pd.notna(course_total_gen_fte):
                worksheet.write_number(current_row, 9, course_total_gen_fte, money_format)
            current_row += 1

            grand_total_fte += course_total_fte if pd.notna(course_total_fte) else 0
            grand_total_gen_fte += course_total_gen_fte if pd.notna(course_total_gen_fte) else 0

        # === Grand total row ===
        worksheet.write(current_row, 0, "Total", total_format)
        for col in range(1, 10):
            if col == 7:
                worksheet.write_number(current_row, col, grand_total_fte, total_money)
            elif col == 9:
                worksheet.write_number(current_row, col, grand_total_gen_fte, total_money)
            else:
                worksheet.write(current_row, col, "", total_money)

        # === Column widths ===
        column_widths = [15, 12, 20, 20, 35, 10, 10, 12, 12, 15]
        for i, width in enumerate(column_widths):
            worksheet.set_column(i, i, width)

    output.seek(0)
    return output


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

# --- Initialize session state ---
if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None

# === Upload Page ===
if not st.session_state.file_uploaded:
    st.title("📁 Upload Course Data File")
    st.markdown("""
    Please upload the **deanDailyCsar.csv** or **deanDailyCsar.xlsx** file to generate faculty FTE reports.
    
    This application will:
    1. Read your uploaded data file
    2. Merge it with the reference data in unique_deansDailyCsar_FTE.xlsx
    3. Calculate FTE values for various reports
    """)
    
    uploaded_file = st.file_uploader("Upload your deanDailyCsar file:", type=["csv", "xlsx"])
    
    if uploaded_file is not None:
        st.success(f"Uploaded file: {uploaded_file.name}")
        
        # Test file load before proceeding - just check if we can read it
        try:
            if uploaded_file.name.endswith('.csv'):
                test_df = pd.read_csv(uploaded_file)
                
            else:
                test_df = pd.read_excel(uploaded_file)

                
            if "Sec Name" not in test_df.columns:
                st.error("The uploaded file appears to be missing required columns (Sec Name). Please check the file format.")
            else:
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("✅ Confirm Upload and Continue"):
                        st.session_state.uploaded_file = uploaded_file
                        st.session_state.file_uploaded = True
                        st.success("File confirmed! Proceeding to merge with reference data...")
                        st.rerun()
                with col2:
                    if st.button("❌ Reset Upload"):
                        # This will clear the file uploader
                        st.experimental_set_query_params()
                        st.rerun()
        except Exception as e:
            st.error(f"Error reading file: {e}")
            st.info("Please ensure the file is in the correct CSV or Excel format.")
            if st.button("❌ Reset Upload"):
                st.experimental_set_query_params()
                st.rerun()
            
    st.stop()  # Stop execution until file is uploaded

# === Main App After Upload ===
uploaded_file = st.session_state.uploaded_file
uploaded_file.seek(0)
# Process the uploaded file and merge with reference data
try:
    # Read the uploaded file (deanDailyCsar)
    if uploaded_file.type == 'text/csv':
        file_in = pd.read_csv(uploaded_file)
    else:
        file_in = pd.read_excel(uploaded_file)
        
    # Read the reference files
    fte_file_in = pd.read_excel("unique_deansDailyCsar_FTE.xlsx")
    fte_tier = pd.read_excel("FTE_Tier.xlsx")
    
    # Extract Course Code if not already present
    if "Course Code" not in file_in.columns:
        file_in["Course Code"] = file_in["Sec Name"].str.extract(r"([A-Z]{3}-\d{3})")
    
    if "Course Code" not in fte_file_in.columns:
        fte_file_in["Course Code"] = fte_file_in["Sec Name"].str.extract(r"([A-Z]{3}-\d{3})")
    
    # Merge the uploaded file with the reference FTE data
    dean_df = pd.merge(
        file_in,
        fte_file_in[["Course Code", "Contact Hours"]],
        how='left',
        on='Course Code'
    )
    
    # Process numeric columns
    dean_df["Contact Hours"] = pd.to_numeric(dean_df["Contact Hours"], errors='coerce')
    dean_df["FTE Count"] = pd.to_numeric(dean_df["FTE Count"], errors='coerce')
    
    # Calculate Total FTE
    dean_df["Total FTE"] = ((dean_df["Contact Hours"] * 16 * 
                           dean_df["FTE Count"]) / 512).round(3)
    
    # Sort the dataframe
    dean_df = dean_df.sort_values(["Sec Divisions", "Sec Name", "Sec Faculty Info"])
    
    # Clean column names
    dean_df.columns = dean_df.columns.str.strip()
    fte_file_in.columns = fte_file_in.columns.str.strip()
    
    # Set flag to show message
    if 'show_success' not in st.session_state:
        st.session_state.show_success = True

    if st.session_state.show_success:
        st.sidebar.success(f"✓ Data loaded successfully! ({len(dean_df)} rows)")
        time.sleep(2)  # Wait 2 seconds
        st.session_state.show_success = False
        st.rerun()
    
except Exception as e:
    st.error(f"Error processing files: {e}")
    st.info("Please ensure all required files are available and properly formatted.")
    
    # Add a button to reset and try again
    if st.button("Reset and Try Again"):
        st.session_state.file_uploaded = False
        st.session_state.uploaded_file = None
        st.rerun()
    st.stop()

st.sidebar.title("Navigation")

# Initialize session state for navigation
if 'nav_choice' not in st.session_state:
    st.session_state.nav_choice = "Home"

# Define buttons for each page
if st.sidebar.button("🏠 Home"):
    st.session_state.nav_choice = "Home"
if st.sidebar.button("📊 Sec Division Report"):
    st.session_state.nav_choice = "Sec Division Report"
if st.sidebar.button("📈 Course Enrollment %"):
    st.session_state.nav_choice = "Course Enrollment Percentage"
if st.sidebar.button("🏫 FTE by Division"):
    st.session_state.nav_choice = "FTE by Division"
if st.sidebar.button("👩‍🏫 FTE per Instructor"):
    st.session_state.nav_choice = "FTE per Instructor"
if st.sidebar.button("📚 FTE per Course"):
    st.session_state.nav_choice = "FTE per Course"

# Set the current choice
choice = st.session_state.nav_choice

# === Page content based on navigation choice ===
if choice == "Home":
    st.title("📘 Faculty FTE Report Generator")
    st.markdown("""
    Welcome to the Faculty FTE Report Generator. This tool helps analyze and visualize FTE:
    
    - **Section Division Reports**: View all courses within an academic division
    - **Course Enrollment Percentages**: Analyze enrollment rates across course sections
    - **FTE by Division**: Calculate FTE metrics for entire academic divisions
    - **FTE per Instructor**: Evaluate faculty teaching loads and generated FTE
    - **FTE per Course**: Compare section performance within specific courses test
    
    
    """)
    
    # st.success(f"Currently using data from: {uploaded_file.name}")
    
    # # Display dataset overview
    # st.subheader("Dataset Overview")
    # st.write(f"Total Rows: {len(dean_df)}")
    # if 'Sec Divisions' in dean_df.columns:
    #     divisions = dean_df['Sec Divisions'].dropna().unique()
    #     st.write(f"Divisions: {', '.join(divisions)}")
        
       
elif choice == "Sec Division Report":
    st.header("FTE by Division")

    if 'Sec Divisions' in dean_df.columns:
        all_divisions = sorted(dean_df['Sec Divisions'].dropna().unique())

        # Multiselect with a "Select All" option
        selected_divisions = st.multiselect("Select Division(s)", options=["Select All"] + all_divisions)

        # Additional input field for custom text entry (comma-separated)
        custom_input = st.text_input("Or enter division names separated by commas:")

        run = st.button("Run Report")

        if run:
            # Handle 'Select All'
            if "Select All" in selected_divisions:
                final_divisions = all_divisions
            else:
                final_divisions = selected_divisions

            # Parse manual input (e.g., "Math,Science,Humanities")
            if custom_input.strip():
                custom_list = [x.strip() for x in custom_input.split(",") if x.strip()]
                final_divisions.extend(custom_list)

            # Remove duplicates and validate
            final_divisions = list(set([div for div in final_divisions if div in all_divisions]))

            if not final_divisions:
                st.warning("No valid divisions selected.")
            else:
                for division_input in final_divisions:
                    st.subheader(f"Report for {division_input}")

                    raw_df, orig_total, gen_total = wf.fte_by_div_raw(dean_df, fte_tier, division_input)

                    if raw_df is not None:
                        report_df = wf.format_fte_output(raw_df, orig_total, gen_total)

                        # Clean and convert for plotting
                        plot_df = report_df[~report_df['Course Code'].isin(['Total', 'DIVISION TOTAL'])].copy()
                        plot_df = plot_df.iloc[:, 2:]
                        plot_df['Generated FTE'] = plot_df['Generated FTE'].str.replace('$', '').str.replace(',', '').astype(float)
                        plot_df = plot_df.sort_values(by='Generated FTE', ascending=False)
                        if len(plot_df) > 10:
                            plot_df = plot_df.head(10)
                        plot_df.index = range(1, len(plot_df) + 1)

                        # Plot
                        fig, ax = plt.subplots(figsize=(10, 6))
                        sns.barplot(data=plot_df, x='Sec Name', y='Generated FTE', ax=ax)
                        ax.set_title(f"Top 10 Sections by Generated FTE – {division_input}")
                        ax.set_xlabel("Section Name")
                        ax.set_ylabel("Generated FTE")
                        plt.xticks(rotation=45, ha='right')
                        st.pyplot(fig)

                        # Save plot to bytes
                        img_bytes = io.BytesIO()
                        fig.savefig(img_bytes, format='png', bbox_inches='tight')
                        img_bytes.seek(0)

                        # Save report
                        save_report(report_df, f"{division_input}_FTE_Report.xlsx", image=img_bytes)

                        st.info(f"Original FTE: {orig_total:.3f}")
                        st.info(f"Generated FTE: ${gen_total:,.2f}")
                    else:
                        st.warning(f"No data found for division {division_input}")
        
elif choice == "Course Enrollment Percentage":
    st.header("Course Enrollment Percentage")
    if 'Sec Name' in dean_df.columns and 'Course Code' in dean_df.columns:
        # Filter only valid course codes that aren't empty
        valid_courses = dean_df['Course Code'].dropna().unique()
        course = st.selectbox("Select Course", valid_courses)
        run = st.button("Run Report")
        if run:
            filtered = dean_df[dean_df['Course Code'] == course]
            filtered = filtered.drop_duplicates(subset="Sec Name")
            filtered["Enrollment Percentage"] = filtered.apply(
                lambda row: wf.calc_enrollment(row), axis=1)  # Use inline lambda function

            # Define calc_enrollment function locally to avoid dependency
            def calc_enrollment(row):
                try:
                    cap = float(row["Capacity"])
                    fte = float(row["FTE Count"])
                    if cap == 0:
                        return "0%"
                    percentage = (fte / cap) * 100
                    return f"{percentage:.2f}%"
                except (ValueError, TypeError, ZeroDivisionError):
                    return "N/A%"

            st.dataframe(filtered)
            save_report(filtered, f"{course}_Course_Report.xlsx")
    else:
        st.warning("This feature will run when 'Sec Name' and 'Course Code' are available in the dataset.")

elif choice == "FTE by Division":
    st.header("FTE by Division")

    if 'Sec Divisions' in dean_df.columns:
        division_input = st.selectbox(
            "Select Division", dean_df['Sec Divisions'].dropna().unique())

        run = st.button("Run Report")

        if run:
            raw_df, orig_total, gen_total = wf.fte_by_div_raw(dean_df, fte_tier, division_input)

            if raw_df is not None:
                report_df = wf.format_fte_output(raw_df, orig_total, gen_total)

                # Format DataFrame
                plot_df = report_df[~report_df['Course Code'].isin(['Total', 'DIVISION TOTAL'])].copy()
                plot_df = plot_df.iloc[:, 2:]

                # Convert Generated FTE from string to numeric
                plot_df['Generated FTE'] = plot_df['Generated FTE'].str.replace('$', '').str.replace(',', '').astype(float)

                # Sort and take top 10 if more than 10 rows
                plot_df = plot_df.sort_values(by='Generated FTE', ascending=False)
                if len(plot_df) > 10:
                    plot_df = plot_df.head(10)

                plot_df.index = range(1, len(plot_df) + 1)

                # Display DataFrame
                st.dataframe(report_df)

                # Plot top 10 with divisions (Sec Name) on X-axis
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(data=plot_df, x='Sec Name', y='Generated FTE', ax=ax)
                ax.set_title("Top 10 Sections by Generated FTE")
                ax.set_xlabel("Section Name")
                ax.set_ylabel("Generated FTE")
                plt.xticks(rotation=45, ha='right')
               
                
                # Display Plot
                st.pyplot(fig)

                # Save Plot as a png
                img_bytes = io.BytesIO()
                fig.savefig(img_bytes, format='png', bbox_inches='tight')
                img_bytes.seek(0)

                # Save button
                save_report(report_df, f"{division_input}_FTE_Report.xlsx", image=img_bytes)

                st.info(f"Total FTE: {orig_total:.3f}")
                st.info(f"Generated FTE: ${gen_total:,.2f}")
            else:
                st.warning(f"No data found for division {division_input}")
    else:
        st.info("Division data not available.")

elif choice == "FTE per Instructor":
    st.header("FTE per Instructor")
    if 'Sec Faculty Info' in dean_df.columns:
        faculty_list = sorted(dean_df["Sec Faculty Info"].dropna().unique())
        instructor = st.selectbox("Select Instructor", faculty_list)

        run = st.button("Run Report")
        if run:
            report_df, orig_fte, gen_fte = wf.generate_faculty_fte_report(dean_df, fte_tier, instructor)

            report_df = report_df.fillna("")
            report_df.index = range(1, len(report_df) + 1)

            st.dataframe(report_df)

            # Chart logic
            plot_df = report_df[report_df['Sec Name'] != 'TOTAL'].copy()
            plot_df['Total FTE'] = pd.to_numeric(plot_df['Total FTE'], errors='coerce')

            fig, ax = plt.subplots(figsize=(10, 6))
            sns.barplot(data=plot_df.sort_values(by='Total FTE', ascending=True).tail(10),
                        x='Total FTE', y='Sec Name', ax=ax)
            ax.set_title(f"Sections by Total FTE for {instructor}")
            ax.set_xlabel("Total FTE")
            ax.set_ylabel("Section")
            plt.tight_layout()
            st.pyplot(fig)

            # ✅ ADD DOWNLOAD BUTTON HERE
            excel_data = save_faculty_excel(report_df, instructor)
            st.download_button("📥 Download Instructor Report",
                               data=excel_data,
                               file_name=f"{instructor.replace(' ', '_')}_FTE_Report.xlsx")

            # Info messages
            st.info(f"Total FTE: {orig_fte:.3f}")
            st.info(f"Generated FTE: ${gen_fte:,.2f}")
    else:
        st.warning("Instructor name column missing.")

elif choice == "FTE per Course":
    st.header("FTE per Course")
    if 'Course Code' in dean_df.columns:
        course_list = sorted(dean_df['Course Code'].dropna().unique())
        course_name = st.selectbox("Select Course", course_list)
        
        run = st.button("Run Report")
        if run:
            df_result, original_fte, generated_fte = wf.calculate_fte_by_course(dean_df, fte_tier, course_name)

            if df_result is not None:
                df_result.index = range(1, len(df_result) + 1)
                st.dataframe(df_result)

                # Create plot data - need to convert Generated FTE from string to numeric
                plot_df = df_result[df_result['Sec Name'] != 'COURSE TOTAL'].copy()
                plot_df['Generated FTE'] = plot_df['Generated FTE'].str.replace('$', '').str.replace(',', '').astype(float)
                
                # Create Figure
                fig, ax = plt.subplots(figsize=(10, 6))
                plot_data = plot_df.sort_values(by='Total FTE', ascending=True)
                sns.barplot(data=plot_data, x='Total FTE', y='Sec Name', ax=ax)
                ax.set_title(f"Sections by Total FTE for Course {course_name}")
                ax.set_xlabel("Total FTE")
                ax.set_ylabel("Section Name")
                
                # Display Plot
                st.pyplot(fig)

                # Save Plot as a png
                img_bytes = io.BytesIO()
                fig.savefig(img_bytes, format='png', bbox_inches='tight')
                img_bytes.seek(0)
                
                save_report(df_result, f"{course_name}_FTE_Report.xlsx", image=img_bytes)

                st.info(f"Original Total FTE: {original_fte:.3f}")
                st.info(f"Generated FTE: ${generated_fte:,.2f}")
            else:
                st.warning(f"No data found for course {course_name}")
    else:
        st.warning("This feature will run when 'Course Code' is present in the dataset.")

# Add a reset button in the sidebar to return to upload page
if st.sidebar.button("Reset Application"):
    st.session_state.file_uploaded = False
    st.session_state.uploaded_file = None
    st.rerun()