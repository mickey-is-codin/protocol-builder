import xlsxwriter

def main():

    print_welcome()
    protocol_name = get_name()

    workbook  = xlsxwriter.Workbook(f"{protocol_name}.xlsx")
    worksheet = workbook.add_worksheet()

    conditions = get_conditions()
    create_header(worksheet, conditions)

    get_configs(conditions)

    workbook.close()

def print_welcome():
    welcome_message = "Welcome to ProtocolBuilder"
    print(f"\n\n{'=' * (len(welcome_message)+4)}")
    print(f"| {welcome_message} |")
    print(f"{'=' * (len(welcome_message)+4)}\n\n")

def get_name():
    print("Please enter the name of your experimental protocol: ")
    return input(" => ")

def get_conditions():
    print("\nPlease enter experimental conditions/variables separated by commas:")
    print("For example: Bed Elevation, Patient Orientation")
    conditions = input(" => ").split(",")
    return [condition.strip() for condition in conditions]

def create_header(worksheet, conditions):

    alphabet = [chr(i) for i in range(ord('A'), ord('Z')+1)]
    header_cells = [alphabet[i] + str(i+1) for i in range(len(conditions))]

    for ix, condition in enumerate(conditions):
        worksheet.write(f"{alphabet[ix]}1", condition)

def get_configs(conditions):

    protocol = {condition: "" for condition in conditions}

    print()
    for condition in conditions:
        print(f"Please enter configurations of {condition} separated by commas:")
        print("For example: 0 degrees, 30 degrees, 45 degrees")
        config = input(" => ").split(",")
        protocol[condition] = [x.strip() for x in config]

    print(protocol)

if __name__ == "__main__":
    main()
