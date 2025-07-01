# EAPB-State-formatting
Pay Balance data formatting for adjustment review - State taxation

Note: the provided data is a modified version of a confidential file. All data is fictional and for demonstration purposes only

Payroll balancing reviews consisted of pulling employee data from a database system which returned a .txt file
Reporting included all potential state taxes for each applicable pay period. Prior to this VBA macro, all actions were manual each time this report needed reviewed (anywhere from 2 a week to 15 or more).

This macro takes the raw .txt file that is opened into excel and it does the following:
Deliminates the data, Formats headers, hides commonly unused data columns, formats column data, inserts a calculation column for Unemploment exemptions; alerts whether there are OR metro taxes that may need secondary reporting pulled, locates and hides all columns with "0" balances, colors headers with matching grouped columns, renames the tab to "State" which is the applicable taxation data.

