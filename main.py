# -*- coding: utf-8 -*-
"""
Takes data from deanDailyCsar.xlsx and FTE_Tier.xlsx to determine classes FTE.
Allows the user to Get Course Enrollment, and get FTE by Division, Instructor,
and Course. 

GROUP A & B
Thuan Chau, Karen Brown, Harley Coughlin,Teresa Hearn, Shiane Ransford, Latoya Winston

04/28/2025

CSC-221-001

M7GroupAnBPro

"""
import functions as fn
import options4 


def main():
    '''
    Main function to handle the menu and the users options.

    Returns
    -------
    None.
    '''
    try:
        choice = 0

        while choice != '6':
            # reads the file into a data frame.
            file_in = fn.readfile()
            # displays menu
            fn.menu()
            # gets user input.
            choice = input("Which one will you choose? ")

            # evaluates the value of choice
            if choice == '1':

                print("option 1: Enter Sec Divison codes")
                # calls the function to write excel file by division.
                fn.sec_divisions(file_in)

            elif choice == '2':
                print("\nOption 2: Enrollment Percentage")
                # calls function to get course enrollment
                fn.option2_enrollment(file_in)

            elif choice == '3':
                print("\nOption 3: FTE by Division: ")
                # gets FTE by divisions
                fn.division_fte(file_in)

            elif choice == '4':
                print("\nOption 4: FTE by Instructor: ")
                # gets FTE by instructor
                options4.fte_per_faculty(file_in)

            elif choice == '5':
                print("\nOption 2: FTE by Course: ")
                # get FTE by course
                fn.fte_per_course(file_in)

            elif choice == '6':
                # exits program
                print("Exit Program: GoodBye!")
            else:
                print("\nPlease include an option between 1 and 6!")

    except FileNotFoundError as err:
        print(f"Error: File not found - {err}")
    except ValueError as err:
        print(f"Error: Invalid value - {err}")


if __name__ == '__main__':
    main()
