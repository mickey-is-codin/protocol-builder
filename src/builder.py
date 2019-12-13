import xlsxwriter

def main():

    welcome_message = "Welcome to ProtocolBuilder"
    print(f"\n\n{'=' * (len(welcome_message)+4)}")
    print(f"| {welcome_message} |")
    print(f"{'=' * (len(welcome_message)+4)}\n\n")

    print("Please enter the name of your experimental protocol: ")
    protocol_name = input(" => ")

    workbook  = xlsxwriter.Workbook(f"{protocol_name}.xlsx")
    worksheet = workbook.add_worksheet()

    print("\nPlease experimental conditions/variables separated by commas:")
    print("For example: Bed Elevation, Patient Orientation")
    conditions = input(" => ").split(",")

    alphabet = [chr(i) for i in range(ord('A'), ord('Z')+1)]
    header_cells = [alphabet[i] + str(i+1) for i in range(len(conditions))]

    for ix, condition in enumerate(conditions):
        worksheet.write(f"{alphabet[ix]}1", condition)

    workbook.close()

if __name__ == "__main__":
    main()
