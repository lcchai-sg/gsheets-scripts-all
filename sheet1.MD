# Google-Sheets
Google sheet scripts

Balance sheet has the following categories:
Date, Check #, Payee, Category, Debit, Debit Subtotals, Credit, Credit Subtotals, Balance

The SortByDate and SortByCategory functions are similar, but sort by column 1 and 4 respectively. Because it would also then sort the row headers, they both delete any row that has the string "Date" in the first column (instead of a date), then reinserts it and sets font to bold.

SumTotals can only be run if the balance sheet is sorted by category. It first checks if the category already exists. If so, the amount is added to the already existing category, writes the subtotal in the subtotals column, and deletes the subtotal from the row above. Otherwise, it creates the new category, and sets the subtotal to that amount, and writes it into the respective subtotals column.

The onOpen function is the default Google function that creates the script menu for ease of access.
