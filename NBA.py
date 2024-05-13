import pandas as pd

customer_names = ["Ammar", "Jana", "Ali", "Sami", "Tarek", "Khollod", "Khaled", "Yehia", "Magdy", "Hoda", "Omar"]

customer_points = {name.lower(): 0 for name in customer_names}

while True:
    customer_name = input("Enter agent name (or 'X' to exit): ").strip()
    
    if customer_name == 'X':
        break
    
    customer_name_lower = customer_name.lower()
    
    if customer_name_lower in customer_points:
        customer_points[customer_name_lower] += 1
    else:
        print("Invalid agent name. Please enter a valid name.")

data = {'Customer': customer_names, 'Points': [customer_points[name.lower()] for name in customer_names]}


total_points = sum(data['Points'])


df = pd.DataFrame(data)

excel_writer = pd.ExcelWriter('NBA_not_full_name.xlsx', engine='xlsxwriter')

df.to_excel(excel_writer, sheet_name='Points', index=False, header=False)

workbook = excel_writer.book
worksheet = excel_writer.sheets['Points']

border_format = workbook.add_format({'border': 1})

worksheet.conditional_format('A1:B12', {'type': 'no_blanks', 'format': border_format})

format = workbook.add_format({'bold': True, 'align': 'center'})

worksheet.set_column('A:B', 15)
worksheet.write('A12', 'Total', format)
worksheet.write('B12', total_points, format)  

excel_writer.close()

print("Data saved to NBA_not_full_name.xlsx")
