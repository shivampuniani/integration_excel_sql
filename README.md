# Python Database Integration Project

This project connects to an Excel file, retrieves data, and inserts it into a SQL Server database. It demonstrates how to use Python to work with SQL using `pyodbc`.

## Requirements

- Python 3.x
- `pyodbc` for database connection
- `pandas` for handling data in Excel format
- `configparser` for reading configuration

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/shivampuniani/integration_excel_sql

2. Install the required Python dependencies:
   ```bash
   pip install -r requirements.txt

3. Configure the database connection:

Update the connection details in the config.ini file with your SQL Server credentials. Here's the format:
   ```bash
   [database]
   server = SQLEXPRESS
   database = Test_Database
   username = sa
   password = 12345678

Update the connection strings (if required) in main.py for both SQL Server and Excel file handling.
SQL Server: Update the SERVER, DATABASE, UID, and PWD placeholders in the config.ini file.
Excel File: Ensure the path to your Excel file is correct in the main.py script.

4. Run the script:
   ```bash
   python main.py


EXCEL-SQL-Integration-project/  
│  
├── main.py               # Your main Python program  
├── requirements.txt      # Python dependencies  
├── README.md             # Project documentation  
├── config.ini            # config file to store and configure sql server and file data   
├── .gitignore            # Git ignore rules  
└── log_file.txt          # Log file (will be generated when running the program)  

