# Pay Balance Data Formatting for Adjustment Review - State Taxation

> **Note:** The provided data is a modified version of a confidential file. All data is fictional and for demonstration purposes only.

| Aspect                 | Description                                                                                                                                                    |
|------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Project Overview**    | This VBA macro was one of my first projects, created using a combination of recorded steps and online resources. While it hasn’t been modified since creation and is not highly optimized, it effectively automates a previously manual process. |
| **Context**            | Payroll balancing reviews involved extracting employee data from a database, which produced a raw `.txt` file containing potential state taxes for each applicable pay period. Before this macro, all processing was manual—ranging from 2 to 15+ reports weekly. |
| **Macro Functionality** | The macro automates processing of the raw `.txt` file opened in Excel by: <ul><li>Delimiting the data</li><li>Formatting headers</li><li>Hiding unused data columns</li><li>Formatting column data</li><li>Inserting a calculation column for Unemployment exemptions</li><li>Alerting for OR metro taxes that may require additional reports</li><li>Hiding all columns with zero balances</li><li>Coloring headers for grouped columns</li><li>Renaming the worksheet tab to "State" to reflect applicable taxation data</li></ul> |

