import matplotlib.pyplot as plt
import csv
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference

def askUser():
    total = 0

    # This loop asks the user for five numbers and adds each one to the total.
    for count in range(5):
        number = float(input("Please enter a number: "))
        total += number

    print("The sum of the 5 numbers entered is:", total)

def askIncome():
    print("\nEnter the information for five people.")

    # This loop asks for five names and incomes and adds them to the CSV file.
    for count in range(5):
        name = input("Please enter a name: ")
        income = int(input("Please enter their annual income: "))

        with open(r"C:\FinalExam\final.csv", "a", newline="") as file:
            file.write(f"\n{name},{income}")

    print("All five entries have been saved.")

def excelPie():
    names = []
    incomes = []

    with open(r"C:\FinalExam\final.csv", "r", newline="") as file:
        reader = csv.reader(file)

        for row in reader:
            if len(row) == 2:
                names.append(row[0])
                incomes.append(int(row[1]))

    # Create a new Excel workbook.
    workbook = Workbook()

    # Select the worksheet for the data.
    sheet = workbook.active
    sheet.title = "Income Data"

    # Add the names and incomes to Excel.
    for name, income in zip(names, incomes):
        sheet.append([name, income])

    # Create a pie chart for the income data.
    pie = PieChart()

    # Set the chart title using my student ID and today's date.
    pie.title = "coryou8501 " + datetime.now().strftime("%B %d, %Y")

    # Select the income values for the pie chart.
    data = Reference(sheet, min_col=2, min_row=1,
                     max_row=len(incomes))

    # Select the names for the chart labels.
    labels = Reference(sheet, min_col=1, min_row=1,
                       max_row=len(names))

    # Add income data to the pie chart.
    pie.add_data(data, titles_from_data=False)

    # Label the chart with each person's name.
    pie.set_categories(labels)

    # Place the chart below the spreadsheet data.
    sheet.add_chart(pie, "A11")

    # Save the Excel file with the chart.
    workbook.save(r"C:\FinalExam\final.xlsx")

    print("Excel pie chart saved as final.xlsx")

def verticalBar():
    names = []
    incomes = []

    with open(r"C:\FinalExam\final.csv", "r", newline="") as file:
        reader = csv.reader(file)

        for row in reader:
            if len(row) == 2:
                names.append(row[0])
                incomes.append(int(row[1]))

    plt.figure(figsize=(10, 6))
    plt.bar(names, incomes, color="green", label="Income")

    plt.title("coryou8501 " + datetime.now().strftime("%B %d, %Y"))
    plt.xlabel("Name")
    plt.ylabel("Income")
    plt.legend()
    plt.tight_layout()

    plt.show()

askUser()
askIncome()
excelPie()
verticalBar()
