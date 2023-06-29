import re

def parse_expenses(data):
    expenses = {"Fixed expenses": {}, "Variable expenses": {}, "Savings and debt repayment": {}}
    categories = list(expenses.keys())
    current_category = None
    current_item = None
    sub_item = False

    for line in data.split('\n'):
        if any(category.lower() in line.lower() for category in categories):
            current_category = next((category for category in categories if category.lower() in line.lower()), None)
            current_item = None
            sub_item = False
        elif current_category:
            match = re.search(r'- \[x\]\s*(.*):\s*RM\s*([\d,NA/]+)', line)
            if match:
                item = match.group(1)
                amount = match.group(2).replace(',', '').replace('N/A', '0').replace(' ', '')
                if amount.isdigit():
                    if current_item and current_item != item:
                        sub_item = False
                    if sub_item:
                        expenses[current_category][current_item] += int(amount)
                    else:
                        current_item = item
                        if item not in expenses[current_category]:
                            expenses[current_category][item] = 0
                        expenses[current_category][current_item] += int(amount)
                        sub_item = True

    return expenses

def calculate_totals(expenses):
    totals = {}
    for category in expenses:
        totals[category] = sum(expenses[category].values())
    totals["Total Expenses"] = sum(totals.values())
    return totals

data = """
1. **Actual Expenses (Cash Flow)**
    - Fixed expenses:
        - [x]  Rent: RM 0 per month
        - [ ]  Utilities: RM 300 per month
        - Insurance: RM X per month
            - [x]  SUAMI: RM 150
            - [x]  ISTERI: RM 171
            - [x]  ANAK: RM 183
        
        Total fixed expenses: RM 921 per month
        
    
    - Variable expenses:
        
        (* Total boleh check dekat app Monefy nanti end of the month)
        
        - [x]  Groceries: RM 500 per month
        - [x]  Mobile Prepaid: RM 50 per month
        - [x]  Transportation: RM 150 per month
        - [x]  Coway: RM N/A per month
        - Belanja ibu bapa: RM 600 per month
            - [x]  SUAMI: RM 0
            - [ ]  ISTERI: RM 200
        - [x]  Belanja anak anak: RM N/A per month
        - [ ]  Sedekah: DEPENDS
        - [ ]  Barang-barang rumah: RM 0*
        - [ ]  Pakaian baru: RM 0*
        - [ ]  Supplements/Medicines: RM 0*
        - [ ]  Entertainment: RM 0*
        - Others: RM 0 per month (Yg unexpected expenses kene taking note in advance supaya easy nak tracking)
            - [x]  Buai automatic: RM 300
            - [x]  Yuran ukm sharifah: RM 1500
        
        Total variable expenses: RM 1,300 per month
        
    
    - Savings and debt repayment:
        - Tabung untuk recovery: RM 1, 300 per month
            - [x]  SUAMI: RM 700
            - [x]  ISTERI: RM 500
        - [x]  Down payment rumah: RM N/A per month
        - [x]  Jalan jalan @ traveling: RM N/A per month
        - [x]  Tabung untuk emergency: RM N/A per month
        - [x]  Car loan repayment: RM 571 per month
        - [ ]  JPA loan repayment: DEPENDS
        - Retirement savings: RM 900 per month (at least must be 10%)
        //Dalam TH+ASB
            - [x]  SUAMI: RM 500
            - [x]  ISTERI: RM 500
        
        Total savings and debt repayment: RM 2,471  per month
""" 

expenses = parse_expenses(data)
totals = calculate_totals(expenses)

for category in totals:
    print(f"{category}: RM {totals[category]}")
