"""
web_functions.py

This module provides backend processing for the Faculty FTE Report Generator
Streamlit application. It contains functions for:

- Reading and merging enrollment and contact hour data
- Calculating original and generated FTE (Full-Time Equivalent) values
- Filtering and formatting reports by division, course, or instructor
- Computing enrollment percentages and course totals
- Exporting cleaned and structured data for reporting

Dependencies
------------
- pandas
- options4 (utility module)

Typical Use
-----------
This module is imported into the Streamlit `app.py` frontend as `wf`.

Example:
    import web_functions as wf
    df = wf.readfile()
    fte_df = wf.fte_by_div_raw(df, tier_df, 'ENG')
"""


import pandas as pd
import options4 as opfour


def readfile():
    """
    Reads, merges, and processes course and FTE data from CSV and Excel sources

    This function:
    - Reads 'deanDailyCsar.csv' and 'unique_deansDailyCsar_FTE.xlsx'
    - Extracts 'Course Code' if missing
    - Merges contact hours into the main dataset
    - Computes total FTE per section

    Returns
    -------
    pd.DataFrame
        Cleaned and sorted merged dataset ready for FTE analysis,
        or an empty list if files are missing.
    """

    try:
        # reads the deansDailyCsar.csv and unique_deansDailyCsar_FTE files in to a dataframe
        file_in = pd.read_csv('deanDailyCsar.csv')
        fte_file_in = pd.read_excel('unique_deansDailyCsar_FTE.xlsx')

        # merge prior dataframes
        # Extract Course Code from Sec Name if not already done
        if "Course Code" not in file_in.columns:
            file_in["Course Code"] = file_in["Sec Name"].str.extract(r"([A-Z]{3}-\d{3})")

        # Also create Course Code in credits_df
        if "Course Code" not in fte_file_in.columns:
            fte_file_in["Course Code"] = fte_file_in["Sec Name"].str.extract(r"([A-Z]{3}-\d{3})")

        # Merge only needed columns from credits_df
        merged_df = pd.merge(
            file_in,
            fte_file_in[["Course Code", "Contact Hours"]],
            how='left',
            on='Course Code'
        )

        merged_df["Contact Hours"] = pd.to_numeric(merged_df["Contact Hours"], errors='coerce')
        merged_df["FTE Count"] = pd.to_numeric(merged_df["FTE Count"], errors='coerce')

        # Calculate Total FTE
        merged_df["Total FTE"] = ((merged_df["Contact Hours"] * 16 *
                                   merged_df["FTE Count"]) / 512).round(3)

        # sorts the dataframe by sec divisions, sec name
        # and sec faculty info and assigns it to groups
        groups = merged_df.sort_values(["Sec Divisions", "Sec Name", "Sec Faculty Info"])

        return groups

    except FileNotFoundError:
        groups = []
        print("File Missing!")
        return groups


def calc_enrollment(row):
    """
    Calculates the enrollment percentage for a course section.

    Parameters
    ----------
    row : pd.Series
        A row from the DataFrame with 'Capacity' and 'FTE Count' fields.

    Returns
    -------
    str
        The enrollment percentage formatted as a string (e.g., "85.71%"),
        or "N/A%" if calculation is not possible.
    """

    try:
        cap = float(row["Capacity"])
        fte = float(row["FTE Count"])

        if cap == 0:
            return "0%"

        percentage = (fte / cap) * 100
        return f"{percentage:.2f}%"

    except (ValueError, TypeError, ZeroDivisionError):
        return "N/A%"


def fte_by_div_raw(file_in, fte_tier, div_code):
    """
    Computes raw and generated FTE totals for a given division.

    Parameters
    ----------
    file_in : pd.DataFrame
        Main dataset containing section-level details.
    fte_tier : pd.DataFrame
        FTE tier multipliers by prefix/course ID.
    div_code : str
        The division code to filter by (e.g., "ENG").

    Returns
    -------
    pd.DataFrame
        DataFrame of section-level and course-level FTE breakdowns.
    float
        Sum of original FTEs in the division.
    float
        Sum of generated FTEs using tier multipliers.
    """

    # Filter division
    div_code = div_code.upper()
    div_data = file_in[file_in['Sec Divisions'] == div_code].copy()

    if div_data.empty:
        return None, 0, 0

    # Create lookup for prefix/course ID → New Sector multiplier
    fte_lookup = {
        row['Prefix/Course ID']: row['New Sector']
        for _, row in fte_tier.iterrows()
        if pd.notna(row['Prefix/Course ID'])
    }

    # Extract course codes from Sec Name
    div_data['Course Code'] = div_data['Sec Name'].str.extract(r'([A-Z]+-\d+)')
    div_data = div_data.sort_values(['Course Code', 'Sec Name'])

    base_fte_value = 1926

    output_rows = []
    current_course = None
    course_total_fte = 0
    first_row = True

    grand_total_original_fte = 0
    grand_total_generated_fte = 0

    for _, row in div_data.iterrows():
        course = row['Course Code']
        sec = row['Sec Name'][:3] if pd.notna(row['Sec Name']) else ""

        new_sector_value = fte_lookup.get(sec, 0)
        total_fte = float(row['Total FTE']) if pd.notna(row['Total FTE']) else 0
        adjusted_fte = total_fte * (new_sector_value + base_fte_value)

        grand_total_original_fte += total_fte

        enrollment_per = ''
        if pd.notna(row['Capacity']) and pd.notna(row['FTE Count']) and float(row['Capacity']) > 0:
            enrollment_per = round((float(row['FTE Count']) / float(row['Capacity'])) * 100, 2)

        if course != current_course and current_course is not None:
            output_rows.append({
                'Division': '',
                'Course Code': 'Total',
                'Sec Name': '',
                'X Sec Delivery Method': '',
                'Meeting Times': '',
                'Capacity': '',
                'FTE Count': '',
                'Sec Faculty Info': '',
                'Total FTE': '',
                'Enrollment Per': '',
                'Generated FTE': course_total_fte
            })
            grand_total_generated_fte += course_total_fte
            course_total_fte = 0

        output_rows.append({
            'Division': div_code if first_row else '',
            'Course Code': course if course != current_course else '',
            'Sec Name': row['Sec Name'],
            'X Sec Delivery Method': row['X Sec Delivery Method'],
            'Meeting Times': row['Meeting Times'],
            'Capacity': row['Capacity'],
            'FTE Count': row['FTE Count'],
            'Sec Faculty Info': row['Sec Faculty Info'],
            'Total FTE': total_fte,
            'Enrollment Per': f"{enrollment_per}%" if enrollment_per != '' else '',
            'Generated FTE': f"{adjusted_fte:.2f}"
        })

        course_total_fte += adjusted_fte
        current_course = course
        first_row = False

    # Add final course total
    if current_course is not None:
        output_rows.append({
            'Division': '',
            'Course Code': 'Total',
            'Sec Name': '',
            'X Sec Delivery Method': '',
            'Meeting Times': '',
            'Capacity': '',
            'FTE Count': '',
            'Sec Faculty Info': '',
            'Total FTE': '',
            'Enrollment Per': '',
            'Generated FTE': course_total_fte
        })
        grand_total_generated_fte += course_total_fte

    output_df = pd.DataFrame(output_rows)
    return output_df, grand_total_original_fte, grand_total_generated_fte


def format_fte_output(raw_df, original_fte_total, generated_fte_total):
    """
    Formats the FTE output DataFrame for display, including currency formatting

    Parameters
    ----------
    raw_df : pd.DataFrame
        Unformatted FTE data generated by `fte_by_div_raw`.
    original_fte_total : float
        Sum of original (unadjusted) FTEs.
    generated_fte_total : float
        Sum of adjusted/generated FTEs.

    Returns
    -------
    pd.DataFrame
        Formatted DataFrame with human-readable FTE values and a
        division total row.
    """

    formatted_rows = []

    for _, row in raw_df.iterrows():
        formatted_row = row.copy()
        if isinstance(row['Generated FTE'], (float, int)):
            formatted_row['Generated FTE'] = "${:,.3f}".format(row['Generated FTE'])
        formatted_rows.append(formatted_row)

    df = pd.DataFrame(formatted_rows)
    df.loc[len(df.index)] = {
        'Division': '',
        'Course Code': 'DIVISION TOTAL',
        'Sec Name': '',
        'X Sec Delivery Method': '',
        'Meeting Times': '',
        'Capacity': '',
        'FTE Count': '',
        'Sec Faculty Info': '',
        'Total FTE': '',
        'Enrollment Per': '',
        'Generated FTE': "${:,.2f}".format(generated_fte_total)
    }

    return df


def calculate_fte_by_course(df, fte_tier, course_code, base_fte=1926):
    """
    Computes FTE statistics for a specific course across all its sections.

    Parameters
    ----------
    df : pd.DataFrame
        The merged dataset with section-level data.
    fte_tier : pd.DataFrame
        Tier multipliers for generating FTE values.
    course_code : str
        The course code (e.g., "ENG-111") to filter.
    base_fte : int, optional
        Base FTE value used in generation, default is 1926.

    Returns
    -------
    pd.DataFrame
        Formatted DataFrame containing FTE and enrollment details by section.
    float
        Total original FTE for the course.
    float
        Total generated FTE using sector multipliers.
    """

    course_code = course_code.upper()
    filtered = df[df['Course Code'] == course_code].copy()
    filtered = filtered.drop_duplicates(subset='Sec Name')

    if filtered.empty:
        return None, 0, 0

    # Load FTE lookup
    fte_lookup = {
        row['Prefix/Course ID']: row['New Sector']
        for _, row in fte_tier.iterrows()
        if pd.notna(row['Prefix/Course ID'])
    }

    output_rows = []
    total_original_fte = 0
    total_generated_fte = 0

    for _, row in filtered.iterrows():
        sec_prefix = row['Sec Name'][:3]
        new_sector = fte_lookup.get(sec_prefix, 0)
        total_fte = float(row['Total FTE']) if pd.notna(row['Total FTE']) else 0
        generated_fte = total_fte * (new_sector + base_fte)
        total_original_fte += total_fte
        total_generated_fte += generated_fte

        enrollment_per = ''
        if pd.notna(row['Capacity']) and pd.notna(row['FTE Count']) and float(row['Capacity']) > 0:
            enrollment_per = round((float(row['FTE Count']) / float(row['Capacity'])) * 100, 2)

        output_rows.append({
            'Sec Name': row['Sec Name'],
            'X Sec Delivery Method': row['X Sec Delivery Method'],
            'Sec Faculty Info': row['Sec Faculty Info'],
            'Meeting Times': row['Meeting Times'],
            'Capacity': row['Capacity'],
            'FTE Count': row['FTE Count'],
            'Total FTE': total_fte,
            'Enrollment Per': f"{enrollment_per}%" if enrollment_per else '',
            'Generated FTE': generated_fte
        })

    # Add summary row
    output_rows.append({
        'Sec Name': 'COURSE TOTAL',
        'X Sec Delivery Method': '',
        'Sec Faculty Info': '',
        'Meeting Times': '',
        'Capacity': '',
        'FTE Count': '',
        'Total FTE': total_original_fte,
        'Enrollment Per': '',
        'Generated FTE': total_generated_fte
    })

    df_out = pd.DataFrame(output_rows)
    df_out['Generated FTE'] = df_out['Generated FTE'].apply(lambda x: "${:,.2f}".format(x) if isinstance(x, (float, int)) else x)

    return df_out, total_original_fte, total_generated_fte


def generate_faculty_fte_report(dean_df, fte_tier, faculty_name):
    """
    Generates an FTE report for a specific faculty member.

    The function:
    - Filters the data for a faculty match
    - Removes duplicate sections
    - Calculates enrollment percentage
    - Computes generated FTE using multipliers

    Parameters
    ----------
    dean_df : pd.DataFrame
        The cleaned and merged dean dataset with all sections.
    fte_tier : pd.DataFrame
        DataFrame mapping course prefixes to FTE sector multipliers.
    faculty_name : str
        Exact match of the faculty member's name from 'Sec Faculty Info'.

    Returns
    -------
    pd.DataFrame
        Formatted faculty FTE report including a summary row.
    float
        Total original FTE assigned to the faculty.
    float
        Total generated FTE for the faculty using multipliers.
    """

    faculty_df = dean_df[dean_df["Sec Faculty Info"] == faculty_name].copy()
    faculty_df = opfour.remove_duplicate_sections(faculty_df)

    faculty_df["Enrollment Per"] = opfour.calculate_enrollment_percentage(
        faculty_df["FTE Count"], faculty_df["Capacity"])

    faculty_df = opfour.generate_fte(faculty_df, fte_tier)

    total_original = faculty_df["Total FTE"].sum()
    total_generated = faculty_df["Generated FTE"].sum()

    faculty_df["Generated FTE"] = faculty_df["Generated FTE"].apply(lambda x: f"${x:,.2f}")

    summary_row = pd.Series({
        "Sec Name": "TOTAL",
        "Total FTE": total_original,
        "Generated FTE": f"${total_generated:,.2f}"
    })

    return pd.concat([faculty_df, pd.DataFrame([summary_row])], ignore_index=True), total_original, total_generated
