from Methods.methods import *


def main():
    filename = input("Please enter the .xlsx file path")
    data = read_file(filename)
    records = insert_data(data)
    number_of_days = calculate_date(records)
    print(number_of_days)

main()
