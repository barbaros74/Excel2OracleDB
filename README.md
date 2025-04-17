# The App. for Data Transfer from Excel to Oracle DB         
A Desktop tool,which was developed by using Pandas and Tkinter libraries of Python3, imports the data from an Excel file into an Oracle DB table while creating the respective table within the DB if it doesn't exist.

## User’s Manual
First of all, a text file must have been created to be referenced during the DB connection with
-   **1st** column stands for the **row id** for the connection credentials
-   **2nd** is for **DB schema name**
-   **3rd** is for                            **DB connection password**
-   **4th** is for                                    **port**
-   **5th** is for                      **DB alias**

such as

![image](https://github.com/user-attachments/assets/1cfdfe78-5d11-484b-b60a-6f9584006eed)

An example implementation :

![image](https://github.com/user-attachments/assets/985ff905-e7d3-465f-a5a4-8c9bb5cab310)

The software inserts the datas, which are extracted from the file specified in the Data source file text field, into the table specified in the Table name text field. All existing tabs of the excel document will be scanned by toggling the Import Data button once, and the progress will be printed out to the darkgrey area.

> [!WARNING]
> Toggling the button without changing the current Connection, Data source file and Table name credentials will duplicate the data in the specified table.



