import matplotlib.pyplot as plt
import pandas as pd
from tabulate import tabulate


class MoneyManager:

    def __init__(self, income, sheet):
        self.sheet = sheet
        self.income = sheet.cell(row=2, column=5).value
        if not isinstance(self.income, (int, float)):
            raise ValueError("Income must be a number")
        self.expenses = []
        self.saving = 0

    def add_expenses(self, expense,sheet):
        self.expenses.append(expense)

        last_row = self.sheet.max_row
        new_row = last_row + 1

        try:
            self.sheet.cell(row=new_row, column=1).value = expense["Category"]
            self.sheet.cell(row=new_row, column=2).value = expense["Item"]
            self.sheet.cell(row=new_row, column=3).value = expense["Amount"]
            previous_cell = self.sheet.cell(row=(new_row - 1), column=4)

            if previous_cell.row == 2:
                self.sheet.cell(row=new_row, column=4).value = self.income - expense["Amount"]
            else:
                self.sheet.cell(row=new_row, column=4).value = (
                        previous_cell.value - expense["Amount"]
                )

            self.sheet.parent.save("Expenses.xlsx")
            print("\nExpense is added successfully!\n")

        except Exception as e:
            print(f"An error occurred while adding the expense: {e}")

    def reports(self,sheet):
        print("\n1) Items' Report\n2) Categories' Report\n3) Categories' Percentages\n4) Categories Costs\n")

        while True:
            choice = input("What Do You Want To Know (Press Q to EXIT): ")

            if choice.lower() == 'q':
                break

            categories_dict = {}
            for row in self.sheet.iter_rows(min_row=3, max_col=4, values_only=True):
                if row[0] is not None and row[2] is not None:
                    price_category = row[0]
                    price = row[2]
                    if price_category not in categories_dict:
                        categories_dict[price_category] = price
                    else:
                        categories_dict[price_category] += price

            if choice == '1':  # Finding Most Expensive & Cheapest Item
                cells_list = list(self.sheet.iter_rows(min_row=3, min_col=2, max_col=4, values_only=True))
                prices_dict = {cell[0]: cell[2] for cell in cells_list if cell[0] is not None and cell[2] is not None}

                if prices_dict:
                    most_expensive_item_price = max(prices_dict.values())
                    cheapest_item_price = min(prices_dict.values())

                    most_expensive_item_name = [key for key, value in prices_dict.items() if value == most_expensive_item_price]
                    cheapest_item_name = [key for key, value in prices_dict.items() if value == cheapest_item_price]

                    print(f"\nYour Most Expensive Item is {most_expensive_item_name} Costs: [$ {most_expensive_item_price}] ")
                    print(f"Your Cheapest Item is {cheapest_item_name} Costs: [$ {cheapest_item_price}]\n ")

                else:
                    print("No items found.")

            elif choice == '2':  # Finding Categories Most & Least Spent on
                if categories_dict:
                    most_category_spent_cost = max(categories_dict.values())
                    least_category_spent_cost = min(categories_dict.values())
                    most_categories_spent_on = [category for category, cost in categories_dict.items() if cost == most_category_spent_cost]
                    least_categories_spent_on = [category for category, cost in categories_dict.items() if cost == least_category_spent_cost]

                    most_category_spent_on_percentage = ((most_category_spent_cost / self.income) * 100)
                    least_category_spent_on_percentage = ((least_category_spent_cost / self.income) * 100)

                    print(f"\nThe Most Category You Spent Your Money on Is: {most_categories_spent_on}"
                          f" Costs: $ {most_category_spent_cost} Which Is {most_category_spent_on_percentage:.2f}% of Your Income")

                    print(f"The Least Category You Spent Your Money on Is: {least_categories_spent_on}"
                          f" Costs: $ {least_category_spent_cost} Which Is {least_category_spent_on_percentage:.2f}% of Your Income\n")

                else:
                    print("No categories found.")

            elif choice == '3':  # Categories' Percentages Charts
                if categories_dict:
                    category_percentage = {category: (amount / self.income) * 100 for category, amount in categories_dict.items()}

                    plt.pie(category_percentage.values(), labels=category_percentage.keys(), autopct="%1.1f%%")
                    plt.title("Categories Percentage")
                    plt.show()
                else:
                    print("No categories found.")

            elif choice == '4':  # Categories' Costs
                if categories_dict:
                    df = pd.DataFrame.from_dict(categories_dict, orient='index', columns=['Cost'])
                    df.reset_index(inplace=True)
                    df.rename(columns={'index': 'Category'}, inplace=True)

                    table_string = tabulate(df, headers='keys', tablefmt="grid", showindex=False)
                    print(table_string)
                    print()
                else:
                    print("No categories found.")

            else:
                print("Please Try a Valid Option")

            print("1) Items' Report\n2) Categories' Report\n3) Categories' Percentages\n4) Categories Costs\n")

    def calculate_savings_goal(self, price, duration):
        try:
            saving = (price / duration)
            monthly_income = float(input("What is Your Monthly Income: "))
            percentage = (saving / monthly_income) * 100

            if saving >= monthly_income:
                print("\nSorry, Your Income is not enough to get it by the months you want!!\n")

            elif saving >= (monthly_income * 0.5):
                print("\nBe Careful, You will live by less than the half of your income for each month!!")
                print(f"Anyways, You need to save {saving:.2f} each month which is {percentage:.2f}% of your Income\n")

            else:
                print(f"\nYou need to save {saving:.2f} each month\n")

        except ValueError:
            print("Invalid input. Please enter a numeric value for your monthly income.")
        except ZeroDivisionError:
            print("Duration must be greater than zero.")
        except Exception as e:
            print(f"An error occurred: {e}")
