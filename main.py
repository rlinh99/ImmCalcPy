from Methods.methods import *
# from datetime import datetime

def main():
    workbook = get_file()
    given_date = get_given_date()
    # given_date = datetime.datetime(2018, 3, 1)
    data = read_workbook(workbook)

    records = insert_data(data)
    number_of_days = calculate_date(records, given_date)
    print("This person has been in Canada for {0} days".format(number_of_days))


main()
