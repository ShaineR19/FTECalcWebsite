"""
Streamlit App – Faculty FTE Report Generator test this

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
import functions as fn

def save_faculty_excel(data, instructor_name, chart_image=None):
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
                worksheet.write_number(current_row, col, grand_total_fte, total_format)
            elif col == 9:
                worksheet.write_number(current_row, col, grand_total_gen_fte, total_money)
            else:
                worksheet.write(current_row, col, "", total_format)

         # === Column widths ===
        column_widths = [15, 12, 20, 20, 35, 10, 10, 12, 12, 15]
        for i, width in enumerate(column_widths):
            worksheet.set_column(i, i, width)

        # === Optional Chart Sheet ===
        if chart_image:
            chart_sheet = workbook.add_worksheet("Generated FTE Chart")
            chart_sheet.insert_image("B2", "chart.png", {"image_data": chart_image})

    output.seek(0)
    return output


def save_report(df_full, default_filename="report.xlsx", image=None):
    """
    Prompts the user to name and download an Excel report.

    Parameters
    ----------
    df_full : pd.DataFrame
        The DataFrame to export.
    default_filename : str
        Default suggested filename (e.g., "report.xlsx").
    image : str or None
        Path to PNG image to embed in a secondary worksheet (optional).
    """

    filename = st.text_input("📄 Enter a filename for your report:", value=default_filename)

    # Ensure .xlsx extension
    if filename and not filename.endswith(".xlsx"):
        filename += ".xlsx"

    if filename:
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write DataFrame
            df_full.to_excel(writer, sheet_name='Full Report', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Full Report']

            # Auto-size columns
            for i, column in enumerate(df_full.columns):
                col_len = max(len(column), df_full[column].astype(str).map(len).max())
                worksheet.set_column(i, i, col_len + 2)

            # Optional chart/image
            if image:
                chart_sheet = workbook.add_worksheet("Graph Report")
                writer.sheets["Graph Report"] = chart_sheet
                chart_sheet.insert_image("A1", image)

        # Display download button
        st.download_button(
            label="💾 Download Excel Report",
            data=output.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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
                
            elif uploaded_file.name.endswith('.xlsx'):
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
    if uploaded_file.name.endswith('.csv'):
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
    - **FTE per Course**: Compare section performance within specific courses
    
    
    """)
    
    # st.success(f"Currently using data from: {uploaded_file.name}")
    
    # # Display dataset overview
    # st.subheader("Dataset Overview")
    # st.write(f"Total Rows: {len(dean_df)}")
    # if 'Sec Divisions' in dean_df.columns:
    #     divisions = dean_df['Sec Divisions'].dropna().unique()
    #     st.write(f"Divisions: {', '.join(divisions)}")
        
       
elif choice == "Sec Division Report":
    st.header("Sec Division Report")

    if 'Sec Divisions' in dean_df.columns:
        all_divisions = sorted(dean_df['Sec Divisions'].dropna().unique())

        # Clean selection controls
        select_all = st.checkbox("Select All Divisions")
        selected_divisions = st.multiselect("Select Division(s)", options=all_divisions)
        #custom_input = st.text_input("Or enter division names separated by commas:")

        # Gather final list
        if select_all:
            final_divisions = all_divisions
        else:
            final_divisions = selected_divisions.copy()
        #if custom_input.strip():
            #custom_list = [x.strip() for x in custom_input.split(",") if x.strip()]
            #final_divisions.extend(custom_list)

        # Validate and deduplicate
        final_divisions = list(set([div for div in final_divisions if div in all_divisions]))

        if final_divisions:
            st.success("Previewing selected division(s):")
            divisions_to_save = []

            for division in final_divisions:
                st.subheader(f"Division: {division}")
                df_div = dean_df[dean_df['Sec Divisions'] == division].copy()

                if "Course Code" in df_div.columns:
                    df_div.drop(columns=["Course Code"], inplace=True)

                st.dataframe(df_div)

                # Optional: Top Total FTE chart
                if 'Sec Name' in df_div.columns and 'Total FTE' in df_div.columns:
                    chart_data = (
                        df_div.groupby('Sec Name', as_index=False)['Total FTE']
                        .sum()
                        .sort_values(by='Total FTE', ascending=False)
                    )
                    if len(chart_data) > 10:
                        chart_data = chart_data.head(10)

                    import matplotlib.pyplot as plt
                    fig, ax = plt.subplots(figsize=(10, 5))
                    ax.bar(chart_data["Sec Name"], chart_data["Total FTE"])
                    ax.set_ylabel("Total FTE")
                    ax.set_xlabel("Section Name")
                    ax.set_title(f"Top {len(chart_data)} Sections by Total FTE - {division}")
                    ax.tick_params(axis='x', rotation=45)
                    st.pyplot(fig)

                    # Save plot to image
                    img_bytes = io.BytesIO()
                    fig.savefig(img_bytes, format='png', bbox_inches='tight')
                    img_bytes.seek(0)

                # Save checkbox
                save_this = st.checkbox(f"Save report for '{division}'?", key=f"save_{division}")
                if save_this:
                    divisions_to_save.append(division)

            # Save button
            save_clicked = st.button("Save Selected Reports", disabled=len(divisions_to_save) == 0)

            if save_clicked:
                for division in divisions_to_save:
                    df_div = dean_df[dean_df['Sec Divisions'] == division].copy()
                    if "Course Code" in df_div.columns:
                        df_div.drop(columns=["Course Code"], inplace=True)

                    filename = f"{division.lower().replace(' ', '_')}.xlsx"
                    df_div.to_excel(filename, index=False)

                    try:
                        fn.auto_format_excel(filename)
                        save_report(df_div, filename, image=img_bytes)
                        st.success(f"Saved and formatted file: {filename}")
                    except Exception as e:
                        st.warning(f"Saved file without formatting: {filename} – {e}")
        else:
            st.info("Please select or enter at least one valid division to preview and save.")
        
elif choice == "Course Enrollment Percentage":
    st.header("Course Enrollment Percentage")
    if 'Sec Name' in dean_df.columns and 'Course Code' in dean_df.columns:
        valid_courses = sorted(dean_df['Course Code'].dropna().unique())
        course = st.selectbox("Select Course", options=["--"] + list(valid_courses))

        run = st.button("Run Report")
        if course != "--" and run:
            filtered = dean_df[dean_df['Course Code'] == course].drop_duplicates(subset="Sec Name").copy()
            def calc_enrollment(row):
                try:
                    cap = float(row["Capacity"])
                    fte = float(row["FTE Count"])
                    if cap == 0:
                        return 0
                    return (fte / cap) * 100
                except (ValueError, TypeError, ZeroDivisionError):
                    return None
            filtered["Enrollment Percentage"] = filtered.apply(calc_enrollment, axis=1)

            # Format for display
            display_df = filtered.copy()
            display_df = display_df.sort_values(by='Enrollment Percentage', ascending=False)
            display_df["Enrollment Percentage"] = display_df["Enrollment Percentage"].apply(
                lambda x: f"{x:.2f}%" if isinstance(x, (float, int)) else "N/A"
            )

            st.dataframe(display_df)
            # Top 10 bar chart
            chart_data = filtered[["Sec Name", "Enrollment Percentage"]].dropna().sort_values(
                by="Enrollment Percentage", ascending=False)

            if not chart_data.empty:
                fig, ax = plt.subplots(figsize=(10, 5))
                ax.bar(chart_data["Sec Name"], chart_data["Enrollment Percentage"])
                ax.set_ylabel("Enrollment Percentage")
                ax.set_xlabel("Section Name")
                ax.set_title(f"Top {len(chart_data)} Sections by Enrollment %")
                ax.tick_params(axis='x', rotation=45)
                plt.tight_layout()  # Add this to fix layout issues
                st.pyplot(fig)

                # Save plot to image
                img_bytes = io.BytesIO()
                fig.savefig(img_bytes, format='png', bbox_inches='tight')
                img_bytes.seek(0)

                # Save Report
                filename = f"{course}_Course_Report.xlsx"
                save_report(display_df, filename, image=img_bytes)
            
            # Download section - Fixed
            #if st.button(f"📥 Download Report for {course}"):
                # Create a temporary file in a location that Streamlit can write to
                #safe_course = course.replace(" ", "_").lower()
                #filename = f"{safe_course}_enrollment_percentage.xlsx"
                #temp_path = os.path.join(tempfile.gettempdir(), filename)
                
                # Save the dataframe to Excel
                #filtered.to_excel()
                
                # Use your existing auto_format_excel function
                #try:
                    #auto_format_excel(filtered)
                    #st.success(f"File formatted successfully")
                #except Exception as e:
                    #st.warning(f"Could not format Excel file: {str(e)}")
                
                # Create download button with the formatted file
                #st.download_button(
                    #label=f"💾 Click here to download {filename}",
                    #data=filtered,
                    #file_name=filename,
                    #mime="application/vnd.ms-excel"
                #)
                
                #st.success(f"Report for {course} is ready for download!")
        
        elif run and course == "--":
            st.warning("Please select a valid course.")
    else:
        st.warning("This feature requires both 'Sec Name' and 'Course Code' columns in the dataset.")

elif choice == "FTE by Division":
    st.header("FTE by Division")

    if 'Sec Divisions' in dean_df.columns:
        all_divisions = sorted(dean_df['Sec Divisions'].dropna().unique())

        # Dropdown + manual entry
        division_select = st.selectbox("Select a Division", options=["--"] + list(all_divisions))
        division_input = st.text_input("Or enter one or more division names (comma-separated):").strip()

        run = st.button("Run Report")

        # Build list of divisions to process
        selected_divisions = []

        if division_input:
            manual_list = [x.strip() for x in division_input.split(",") if x.strip()]
            selected_divisions.extend(manual_list)

        if division_select != "--":
            selected_divisions.append(division_select)

        # Filter only valid divisions
        selected_divisions = list(set([div for div in selected_divisions if div in all_divisions]))

        if run and selected_divisions:
            for div in selected_divisions:
                st.subheader(f"Division: {div}")
                raw_df, orig_total, gen_total = wf.fte_by_div_raw(dean_df, fte_tier, div)

                if raw_df is not None:
                    report_df = wf.format_fte_output(raw_df, orig_total, gen_total)

                    # Format for both plot and dataframe
                    plot_df = report_df[~report_df['Course Code'].isin(['Total', 'DIVISION TOTAL'])].copy()
                    plot_df = plot_df.iloc[:, 2:]
                    plot_df['Generated FTE Float'] = plot_df['Generated FTE'].str.replace('$', '').str.replace(',', '').astype(float)
                    plot_df = plot_df.sort_values(by='Generated FTE Float', ascending=False)
                    plot_df.index = range(1, len(plot_df) + 1)

                    # Format for Dataframe
                    frame_df = plot_df.copy()
                    frame_df = frame_df.iloc[:, :-1]
                    
                    # Display Dataframe
                    st.dataframe(frame_df.head(10))

                    # Plot chart
                    fig, ax = plt.subplots(figsize=(10, 6))
                    sns.barplot(data=plot_df.head(10), x='Sec Name', y='Generated FTE Float', ax=ax)
                    ax.set_title("Top 10 Sections by Generated FTE")
                    ax.set_xlabel("Section Name")
                    ax.set_ylabel("Generated FTE ($)")
                    plt.xticks(rotation=45, ha='right')

                    # Display Plot
                    st.pyplot(fig)

                    # Save plot to image
                    img_bytes = io.BytesIO()
                    fig.savefig(img_bytes, format='png', bbox_inches='tight')
                    img_bytes.seek(0)

                    # Save Excel + image
                    save_report(report_df, f"{div}_fte.xlsx", image=img_bytes)
                    
                    # Summary stats
                    st.info(f"Total FTE: {orig_total:.3f}")
                    st.info(f"Generated FTE: ${gen_total:,.2f}")
                else:
                    st.warning(f"No data found for division: {div}")
        elif run and not selected_divisions:
            st.warning("Please select or enter at least one valid division.")
    else:
        st.info("Division data not available.")

elif choice == "FTE per Instructor":
    st.header("FTE per Instructor")
    if 'Sec Faculty Info' in dean_df.columns:
        faculty_list = sorted(dean_df["Sec Faculty Info"].dropna().unique())
        instructor = st.selectbox("Select Instructor", ["--"] + faculty_list)

        run = st.button("Run Report")
        if run and instructor != "--":

            report_df, orig_fte, gen_fte = wf.generate_faculty_fte_report(dean_df, fte_tier, instructor)

            report_df = report_df.fillna("")
            report_df.index = range(1, len(report_df) + 1)

            # Remove existing total row
            report_df = report_df[~report_df["Course Code"].astype(str).str.upper().eq("TOTAL")]

            # Calculate grand totals
            final_total_fte = report_df["Total FTE"].sum()
            final_total_gen = report_df["Generated FTE"].sum()

            # Build display rows
            display_rows = []
            first_row = True
            for course, group in report_df.groupby("Course Code"):
                show_course = True
                course_total_fte = group['Total FTE'].sum()
                course_gen_fte = group['Generated FTE'].sum()

                for idx, row in group.iterrows():
                    display_rows.append({
                        "Instructor": instructor if first_row else "",
                        "Course Code": course if show_course else "",
                        "Sec Name": row.get("Sec Name", ""),
                        "X Sec Delivery Method": row.get("X Sec Delivery Method", ""),
                        "Meeting Times": row.get("Meeting Times", ""),
                        "Capacity": row.get("Capacity", ""),
                        "FTE Count": row.get("FTE Count", ""),
                        "Total FTE": row.get("Total FTE", 0),
                        "Sec Divisions": row.get("Sec Divisions", ""),
                        "Generated FTE": row.get("Generated FTE", 0)
                    })
                    first_row = False
                    show_course = False

                display_rows.append({
                    "Instructor": "",
                    "Course Code": "",
                    "Sec Name": "SUBTOTAL",
                    "X Sec Delivery Method": "",
                    "Meeting Times": "",
                    "Capacity": "",
                    "FTE Count": "",
                    "Total FTE": course_total_fte,
                    "Sec Divisions": "",
                    "Generated FTE": course_gen_fte
                })

            # Append grand total row
            display_rows.append({
                "Instructor": "",
                "Course Code": "",
                "Sec Name": "TOTAL",
                "X Sec Delivery Method": "",
                "Meeting Times": "",
                "Capacity": "",
                "FTE Count": "",
                "Total FTE": final_total_fte,
                "Sec Divisions": "",
                "Generated FTE": final_total_gen
            })

            # Convert to DataFrame
            df_display = pd.DataFrame(display_rows)
            df_display.index = range(1, len(df_display) + 1)

            # Convert numeric columns
            for col in ['Capacity', 'FTE Count', 'Total FTE', 'Generated FTE']:
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)

            # Format for display
            df_display['Generated FTE'] = df_display['Generated FTE'].apply(lambda x: f"${x:,.2f}")
            df_display['Total FTE'] = df_display['Total FTE'].apply(lambda x: f"{x:.3f}")
            df_display['Capacity'] = df_display['Capacity'].astype(int)

            # Show table
            st.dataframe(df_display)

            # Plot
            plot_df = report_df.copy()
            plot_df['Sec Name'] = plot_df['Sec Name'].astype(str).str.strip()
            plot_df = plot_df[
                (plot_df['Sec Name'] != "") &
                (~plot_df['Sec Name'].str.upper().isin(['TOTAL', 'SUBTOTAL']))
            ]
            plot_df['Generated FTE'] = pd.to_numeric(plot_df['Generated FTE'], errors='coerce')

            fig, ax = plt.subplots(figsize=(10, 6))
            top_sections = plot_df.sort_values(by='Generated FTE', ascending=False).head(5)
            sns.barplot(data=top_sections, x='Sec Name', y='Generated FTE', ax=ax)
            ax.set_title(f"Top 5 Sections by Generated FTE for {instructor}")
            ax.set_ylabel("Generated FTE ($)")
            ax.set_xlabel("Section")
            ax.tick_params(axis='x', rotation=45)
            plt.tight_layout()
            st.pyplot(fig)

            img_buffer = io.BytesIO()
            fig.savefig(img_buffer, format='png', bbox_inches='tight')
            img_buffer.seek(0)

            # Download button
            excel_data = save_faculty_excel(report_df, instructor_name=instructor, chart_image=img_buffer)
            st.download_button("📥 Download Instructor Report",
                               data=excel_data,
                               file_name=opfour.clean_instructor_name(instructor))

            # Info messages
            st.info(f"Total FTE: {orig_fte:.3f}")
            st.info(f"Generated FTE: ${gen_fte:,.2f}")
    else:
        st.warning("Instructor name column missing.")

elif choice == "FTE per Course":
    st.header("FTE per Course")
    if 'Course Code' in dean_df.columns:
        course_list = sorted(dean_df['Course Code'].dropna().unique())
        course_name = st.selectbox("Select Course", ["--"] + course_list)
        
        run = st.button("Run Report")
        if run and course_name != "--":
            df_result, original_fte, generated_fte = wf.calculate_fte_by_course(dean_df, fte_tier, course_name)

            if df_result is not None:
                # Clean Totals from df_result
                plot_df = df_result[df_result['Sec Name'] != 'COURSE TOTAL'].copy()

                # Add a Numeric FTE for Sorting
                plot_df['Generated FTE Float'] = plot_df['Generated FTE'].str.replace('$', '').str.replace(',', '').astype(float)

                # Sort by Generated FTE Float
                plot_df = plot_df.sort_values(by='Generated FTE Float', ascending=False)

                # Reset Index
                plot_df.index = range(1, len(plot_df) + 1)

                # Create a Dataframe for Display and remove Generated FTE Float
                report_df = plot_df.copy()

                # Drop Generated FTE float
                report_df = report_df.iloc[:, :-1]

                # Display Dataframe top 10
                st.dataframe(report_df.head(10))

                # Create and flip the chart
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(data=plot_df.head(10), 
                x='Sec Name', 
                y='Generated FTE Float', 
                ax=ax, palette='Greens_r'
                )

                # Label and style
                ax.set_title(f"Top 10 Sections by Generated FTE for Course {course_name}", fontsize=16, weight='bold')
                ax.set_xlabel("Section Name", fontsize=14, weight='bold')
                ax.set_ylabel("Generated FTE ($)", fontsize=14, weight='bold')
                ax.tick_params(axis='y', labelsize =10)
                ax.tick_params(axis='x', rotation=45)
                # Data Labels
                for container in ax.containers:
                    ax.bar_label(container, fmt='${:,.2f}', padding=5, fontsize=10, color='black')

                # Show chart and save for Excel
                #sns.despine(ax=ax, left=True, bottom=True)
                plt.tight_layout()
                st.pyplot(fig)

                # Save Plot as a png
                img_bytes = io.BytesIO()
                fig.savefig(img_bytes, format='png', bbox_inches='tight')
                img_bytes.seek(0)
                
                save_report(df_result, f"{course_name}_FTE_Report.xlsx", image=img_bytes)

                st.info(f"Total FTE: {original_fte:.3f}")
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
