import pandas as pd
import openpyxl
from datetime import datetime, date

# ask user which customer they would like to calculate for
customer = input("Which customer would you like to calculate for? ")
# customer = "Howard"

# read the customer.xlsx file
customer_df = pd.read_excel("customers.xlsx")

# find the customer in the customer.xlsx file
customer_df = customer_df.loc[customer_df["Name"] == customer]

# if the customer is not found, exit the program
if customer_df.empty:
    print("Customer not found")
    exit()

# get the earliest transaction date and convert it to a date object
earliest_transaction_date = customer_df["EarliestTransactionDate"].values[0]
earliest_transaction_date = datetime.strptime(earliest_transaction_date, "%Y‐%m‐%d").date()

# get the latest transaction date and convert it to a date object
latest_transaction_date = customer_df["LatestTransactionDate"].values[0]
latest_transaction_date = datetime.strptime(latest_transaction_date, "%Y‐%m‐%d").date()

total_transation_amount = customer_df["TotalTransactionAmount"].values[0]

total_days = (latest_transaction_date - earliest_transaction_date).days + 1 # add 1 to be inclusive

# get transaction per day amount
transaction_per_day = (total_transation_amount / total_days)

# import periods from periods.xlsx, array of objects with the following properties PeriodTitle, StartDate, EndDate
periods = pd.read_excel("periods.xlsx").to_dict("records")
for period in periods:
    period["StartDate"] = datetime.strptime(period["StartDate"], "%Y‐%m‐%d").date()
    period["EndDate"] = datetime.strptime(period["EndDate"], "%Y‐%m‐%d").date()
    period["length"] = (period["EndDate"] - period["StartDate"]).days

period_amounts = []

# find the first period that has an end date after the earliest transaction date
in_period = False
done = False

# count tranactions from before the first period
if earliest_transaction_date < periods[0]["StartDate"]:
    # if the latest transaction date is before the first period
    if latest_transaction_date < periods[0]["StartDate"]:
        amount = transaction_per_day * abs((latest_transaction_date - earliest_transaction_date).days + 1)
        period_amounts.append({"PeriodTitle": "Before " + periods[0]["Title"], "PeriodStartDate": earliest_transaction_date, "PeriodEndDate": latest_transaction_date, "AllocatedAmount": amount.round(0)})
        done = True
    else:
        days = (periods[0]["StartDate"] - earliest_transaction_date).days
        amount = transaction_per_day * abs((periods[0]["StartDate"] - earliest_transaction_date).days)
        period_amounts.append({"PeriodTitle": "Before " + periods[0]["Title"], "PeriodStartDate": earliest_transaction_date, "PeriodEndDate": periods[0]["StartDate"], "AllocatedAmount": amount.round(0)})

for period in periods:
    if done:
        break
    if period["EndDate"] >= earliest_transaction_date:
        first_period = period
        in_period = True
    # remove the period from the periods array
    if in_period:
        # if period starts after the latest transaction date - should break out of the loop
        if period["StartDate"] > latest_transaction_date:
            done = True
            break
        # if period starts before the earliest transaction date
        if period["StartDate"] < earliest_transaction_date:
            # calculate the amount for the period
            amount = transaction_per_day * ((period["EndDate"] - earliest_transaction_date).days + 1)
            # round amount to nearest integer
            period_amounts.append({"PeriodTitle": period["Title"], "PeriodStartDate": earliest_transaction_date, "PeriodEndDate": period["EndDate"], "AllocatedAmount": amount.round(0)})
        # if period ends after the latest transaction date
        elif period["EndDate"] > latest_transaction_date:
            # calculate the amount for the period
            amount = transaction_per_day * ((latest_transaction_date - period["StartDate"]).days + 1)
            period_amounts.append({"PeriodTitle": period["Title"], "PeriodStartDate": period["StartDate"], "PeriodEndDate": latest_transaction_date, "AllocatedAmount": amount.round(0)})
        # smack in the middle
        else:
            amount = transaction_per_day * (period["length"] + 1)
            period_amounts.append({"PeriodTitle": period["Title"], "PeriodStartDate": period["StartDate"], "PeriodEndDate": period["EndDate"], "AllocatedAmount": amount.round(0)})

# if the latest transaction date is after the last period
if not done:
    amount = transaction_per_day * abs((periods[len(periods) - 1]["EndDate"] - latest_transaction_date).days)
    period_amounts.append({"PeriodTitle": "After " + periods[len(periods) - 1]["Title"], "PeriodStartDate": periods[len(periods) - 1]["EndDate"], "PeriodEndDate": latest_transaction_date, "AllocatedAmount": amount.round(0)})

# display the results
print("CustomerName PeriodTitle PeriodStartDate PeriodEndDate AllocatedAmount")
for period_amount in period_amounts:
    print(customer, period_amount["PeriodTitle"], period_amount["PeriodStartDate"], period_amount["PeriodEndDate"], "$" + str(period_amount["AllocatedAmount"]))

# read totals
total = 0
for period_amount in period_amounts:
    total += period_amount["AllocatedAmount"]

print("Total: $" + str(total))

# write the results to an excel file
df = pd.DataFrame(period_amounts)
df.to_excel('{customer}_results.xlsx'.format(customer=customer), index=False)