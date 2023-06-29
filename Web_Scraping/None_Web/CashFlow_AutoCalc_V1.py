def calculate_expenses(filename):
    expenses = {"Fixed Expenses": {}, "Variable Expenses": {}, "Savings and Debt Repayment": {}}
    with open(filename, 'r') as file:
        for line in file:
            category, item, amount = line.strip().split(',')
            expenses[category][item] = int(amount)
    return expenses

def calculate_totals(expenses):
    totals = {}
    for category in expenses:
        totals[category] = sum(expenses[category].values())
    totals["Total Expenses"] = sum(totals.values())
    return totals

expenses = calculate_expenses('expenses.txt')
totals = calculate_totals(expenses)

for category in totals:
    print(f"{category}: RM {totals[category]}")
