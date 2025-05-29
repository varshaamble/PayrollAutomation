SalaryAutomation

# 💰 SalaryAutomation

A Python-based desktop application that automates the process of monthly salary calculation, slip generation, and employee salary record management.
----

## 📌 Features
- Add and update employee salary details
- Automatically calculate deductions (PF, ESIC, PT, etc.)
- Generate PDF salary slips
- Store salary records month-wise
- Export reports for management
- Easy-to-use GUI interface built with Tkinter
- MySQL database integration
----

*****PDF Data Format & Import Instructions*****
This application supports loading employee data directly from a standardized PDF file.

 Supported PDF Titles (Column Names)
           1. SR No
           2. Full Name of Employee
           3. UAN NO	
           4. PF NO/ MEMBER ID	
           5. ESIC NUMBER	
           6. Category
           7. Week Off
           8. Minimum Wages Per Day	
           9. Department	
           10. Location
           11. PanNo
           12. Aadhaar_No
           13. Join Date
           14. Date of Birth	
           15. MobileNo
           16. Site Expenses

           
            entered data can be in a Standard Format
*****@ How to Use*****
     Click "Use PDF Data" or "Load File" from the application menu.

     Select the PDF file containing the above titles.

     The application will extract data row-wise and insert it into the internal database in a standardized format.

     You can then edit, save, or generate salary slips for each employee.
-----

## 🛠️ Tech Stack
- Programming Language: Python
- GUI Framework: Tkinter
- Database: MySQL
- PDF Generation: FPDF
- Other Tools: XAMPP (local server)
----

## ⚙️ Installation

### Prerequisites
- Python 3.x installed
- MySQL Server (XAMPP or standalone)
- Required Python libraries
----


🗃️ Database Setup Instructions
  This application uses MySQL for storing employee and salary data. You can use XAMPP as your local server.

🛠️ Steps to Setup
Download and Install XAMPP
Visit: https://www.apachefriends.org/index.html

Open XAMPP Control Panel

Start Apache and MySQL services

Click Admin next to MySQL to open phpMyAdmin

Database Configuration in Code
The app uses the following credentials to connect to MySQL:

Database Configuration in Code
     host = 'localhost'
     user = 'root'
     password = ''
     database = 'allemployees'
     
Create the Database

In phpMyAdmin, click on New

Create a database named: allemployees

No need to create tables manually
→ The application will create all necessary tables automatically when run

 ### Clone the Repository
```bash
cd SalaryAutomation

**Install Dependencies
pip install fpdf mysql-connector-python

--Run the Application
python main.py

  
  
