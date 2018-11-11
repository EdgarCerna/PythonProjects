import math
import xlsxwriter

# Prompt User
observedQ = math.sqrt(float(input("Enter OBSERVED Percent of Recessive Gene:  ")) / 100.0)
observedP = 1 - observedQ

expectedQ = math.sqrt(float(input("Enter EXPECTED Percent of Recessive Gene:  ")) / 100.0)
expectedP = 1- expectedQ

numPeople = int(input("Enter Number of People:  "))

# Chi Squared Analysis
firstO = (observedP * observedP) * numPeople
firstE = (expectedP * expectedP) * numPeople
first = ((firstO - firstE) * (firstO - firstE)) / firstE

secondO = (2 * observedP * observedQ) * numPeople
secondE = (2 * expectedP * expectedQ) * numPeople
second = ((secondO - secondE) * (secondO - secondE)) / secondE

thirdO = (observedQ * observedQ) * numPeople
thirdE = (expectedQ * expectedQ) * numPeople
third = ((thirdO - thirdE) * (thirdO - thirdE)) / thirdE

result = first + second + third

# Create a workbook and add a worksheet
workbook = xlsxwriter.Workbook("Genetics01.xlsx")
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet
worksheet.write_row(
    "A1:D1",
    ['', 'Number (O)', 'Number (E)', '(O-E)^2 / E']
)
worksheet.write_row(
    "A2:D2",
    ['p^2', firstO, firstE, first]
)
worksheet.write_row(
    "A3:D3",
    ['2pq', secondO, secondE, second]
)
worksheet.write_row(
    "A4:D4",
    ['q^2', thirdO, thirdE, third]
)
worksheet.write_row(
    "A5:D5",
    ['', '', 'EX^2', result]
)

workbook.close()