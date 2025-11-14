# Email automation using python

##Project Overview

This project automates the process of generating a daily report of the top 5 rented movies from the Sakila database and sending it via email. It connects to a MySQL database, fetches the required data, saves it into a professionally formatted Excel file, and emails the report automatically.

## Project Steps
### 1. Importing Libraries

- The project uses Python libraries for database connection, data manipulation, Excel formatting, and email automation.
- Libraries for MySQL connection allow Python to fetch data directly from the Sakila database.
- Pandas is used to handle and process data efficiently.
- Openpyxl is used to format the Excel file with titles, bold headers, and filters.
- SMTP and email libraries handle sending the Excel report via email.


### 2. Connect to MySQL Database

A connection is established to the Sakila database running locally. Database credentials, including host, username, password, and database name, are provided to access the required data.


### 3. Execute SQL Query

A query retrieves the top 5 most rented movies. Relevant tables are joined to calculate rental counts per movie, and the result is stored in a pandas DataFrame for further processing.


### 4. Save Data to Excel

The fetched data is saved to an Excel file. The process leaves space at the top for a report title.


### 5. Format Excel File

- The Excel file is enhanced with a professional layout:
- A bold, centered title is added at the top.
- Column headers are bolded and centered.
- Filters are applied to columns for easy sorting and analysis.


### 6. Send Email with Attachment

The formatted Excel report is sent automatically via email. The process uses Gmail SMTP and requires an app password if two-factor authentication is enabled. The report is attached to the email and sent to the intended recipient.


## Project Features

- Automated daily report generation.
- Professional Excel formatting: title, bold headers, and filters.
- Email automation with attachment.
- Fully Python-based and easily extendable.

## Future Improvements

- Include additional KPIs such as new customers, revenue, and total rentals.
- Combine all KPIs into a multi-sheet Excel file.
- Schedule the script using Windows Task Scheduler or cron for fully automated daily execution.
