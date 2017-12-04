# OpenOffice
The facility for a spreadsheet

Frequently, we use a spreadsheet to note various transactions. The examples of sequences of transactions are collection of incomes, shoppings, donations and many more. There are three fields mostly present in a transaction: description, amount and date.
We organize the transactions in form of data range in a spreadsheet. Each row contains one transaction.
The facility allows appending new transaction at the bottom of the data range taking as a pattern previous transactions. We use an active cell to mark the pattern.
The macro copies the row with the selected transaction to the first empty line at the bottom of the a range. The newly created transaction is automatically updated with a current date. For this reason, the first field in the new row formated as a date type is being searched and substituted by the current date.
Finally, the position of a current cell is set to the same column as selected at the beginning but in newly created row.
The width of copied part is limited by empty cells (columns).

The presented macro has been developed in Open Office Basic.

antekg0@poczta.onet.pl
