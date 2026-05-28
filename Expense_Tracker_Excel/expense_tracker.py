import json
from datetime import datetime
from openpyxl import Workbook
import matplotlib.pyplot as plt

FILE_NAME = "expenses.json"

# Default categories
categories = [
    "Groceries",
    "Transportation",
    "Entertainment",
    "Utilities"
]

# Budget dictionary
budgets = {}


# ---------------- LOAD DATA ---------------- #

def load_expenses():

    try:
        with open(FILE_NAME, "r") as file:
            return json.load(file)

    except FileNotFoundError:
        return []

    except json.JSONDecodeError:
        return []


# ---------------- SAVE DATA ---------------- #

def save_expenses(expenses):

    with open(FILE_NAME, "w") as file:
        json.dump(expenses, file, indent=4)


# ---------------- ADD EXPENSE ---------------- #

def add_expense(expenses):

    try:

        amount = float(input("Enter amount: ₹"))

        description = input(
            "Enter description: "
        )

        date = input(
            "Enter date (YYYY-MM-DD): "
        )

        print("\nCategories:")

        for i, cat in enumerate(categories, start=1):
            print(f"{i}. {cat}")

        print(f"{len(categories)+1}. Add New Category")

        choice = int(
            input("Choose category: ")
        )

        if choice == len(categories) + 1:

            new_category = input(
                "Enter new category: "
            )

            categories.append(new_category)

            category = new_category

        else:
            category = categories[choice - 1]

        expense = {
            "amount": amount,
            "description": description,
            "date": date,
            "category": category
        }

        expenses.append(expense)

        save_expenses(expenses)

        print("\nExpense added successfully!")

    except ValueError:
        print("Invalid input.")


# ---------------- VIEW EXPENSES ---------------- #

def view_expenses(expenses):

    if not expenses:
        print("\nNo expenses found.")
        return

    print("\nExpense Records")
    print("-" * 60)

    for index, exp in enumerate(expenses):

        print(
            f"{index+1}. "
            f"{exp['date']} | "
            f"{exp['description']} | "
            f"₹{exp['amount']} | "
            f"{exp['category']}"
        )


# ---------------- EDIT EXPENSE ---------------- #

def edit_expense(expenses):

    view_expenses(expenses)

    try:

        index = int(
            input("\nEnter expense number to edit: ")
        ) - 1

        if index < 0 or index >= len(expenses):
            print("Invalid number.")
            return

        expenses[index]["amount"] = float(
            input("New amount: ₹")
        )

        expenses[index]["description"] = input(
            "New description: "
        )

        expenses[index]["date"] = input(
            "New date (YYYY-MM-DD): "
        )

        expenses[index]["category"] = input(
            "New category: "
        )

        save_expenses(expenses)

        print("Expense updated successfully!")

    except ValueError:
        print("Invalid input.")


# ---------------- DELETE EXPENSE ---------------- #

def delete_expense(expenses):

    view_expenses(expenses)

    try:

        index = int(
            input("\nEnter expense number to delete: ")
        ) - 1

        if index < 0 or index >= len(expenses):
            print("Invalid number.")
            return

        deleted = expenses.pop(index)

        save_expenses(expenses)

        print(
            f"Deleted: {deleted['description']}"
        )

    except ValueError:
        print("Invalid input.")


# ---------------- SEARCH EXPENSE ---------------- #

def search_expense(expenses):

    keyword = input(
        "Enter keyword: "
    ).lower()

    found = False

    for exp in expenses:

        if keyword in exp["description"].lower():

            print(exp)

            found = True

    if not found:
        print("No matching expense found.")


# ---------------- SUMMARY ---------------- #

def expense_summary(expenses):

    if not expenses:
        print("No expenses found.")
        return

    total = 0

    category_totals = {}

    for exp in expenses:

        total += exp["amount"]

        category = exp["category"]

        if category in category_totals:

            category_totals[category] += exp["amount"]

        else:
            category_totals[category] = exp["amount"]

    print("\nExpense Summary")
    print("-" * 40)

    print(f"Total Spending: ₹{total:.2f}")

    print("\nCategory Wise Spending:")

    for cat, amount in category_totals.items():

        print(f"{cat}: ₹{amount:.2f}")

        if cat in budgets:

            if amount > budgets[cat]:

                print(
                    f"WARNING! "
                    f"{cat} budget exceeded!"
                )

    average = total / len(expenses)

    print(
        f"\nAverage Expense: ₹{average:.2f}"
    )

    highest = max(
        expenses,
        key=lambda x: x["amount"]
    )

    lowest = min(
        expenses,
        key=lambda x: x["amount"]
    )

    print("\nHighest Expense:")

    print(highest)

    print("\nLowest Expense:")

    print(lowest)

    return category_totals


# ---------------- EXPORT TO EXCEL ---------------- #

def export_to_excel(expenses):

    if not expenses:
        print("No expenses to export.")
        return

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Expenses"

    headers = [
        "Date",
        "Description",
        "Amount",
        "Category"
    ]

    sheet.append(headers)

    for exp in expenses:

        sheet.append([
            exp["date"],
            exp["description"],
            exp["amount"],
            exp["category"]
        ])

    workbook.save("expenses.xlsx")

    print(
        "Expenses exported to expenses.xlsx"
    )


# ---------------- CHART ---------------- #

def show_chart(category_totals):

    labels = list(category_totals.keys())

    amounts = list(category_totals.values())

    plt.pie(
        amounts,
        labels=labels,
        autopct="%1.1f%%"
    )

    plt.title("Expense Distribution")

    plt.show()


# ---------------- SET BUDGET ---------------- #

def set_budget():

    category = input(
        "Enter category: "
    )

    try:

        amount = float(
            input("Enter budget amount: ₹")
        )

        budgets[category] = amount

        print("Budget saved.")

    except ValueError:
        print("Invalid amount.")


# ---------------- MAIN PROGRAM ---------------- #

def main():

    expenses = load_expenses()

    while True:

        print("\n===== Expense Tracker =====")

        print("1. Add Expense")
        print("2. View Expenses")
        print("3. Edit Expense")
        print("4. Delete Expense")
        print("5. Search Expense")
        print("6. Expense Summary")
        print("7. Export To Excel")
        print("8. Show Expense Chart")
        print("9. Set Budget")
        print("10. Exit")

        choice = input("\nEnter choice: ")

        if choice == "1":

            add_expense(expenses)

        elif choice == "2":

            view_expenses(expenses)

        elif choice == "3":

            edit_expense(expenses)

        elif choice == "4":

            delete_expense(expenses)

        elif choice == "5":

            search_expense(expenses)

        elif choice == "6":

            category_totals = expense_summary(expenses)

        elif choice == "7":

            export_to_excel(expenses)

        elif choice == "8":

            category_totals = expense_summary(expenses)

            show_chart(category_totals)

        elif choice == "9":

            set_budget()

        elif choice == "10":

            print(
                "Thank you for using Expense Tracker!"
            )

            break

        else:
            print("Invalid choice.")


# ---------------- START PROGRAM ---------------- #

if __name__ == "__main__":
    main()