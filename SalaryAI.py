import re
import pdfplumber
import mysql.connector
from mysql.connector import Error
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk, messagebox
from datetime import datetime
from PyPDF2 import PdfReader
import os
import calendar
from fpdf import FPDF
import pyautogui
import time
import webbrowser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
driver = None

#Dynamically creates table
def create_monthly_table(selected_month, selected_year):
    table_name = f"employees_{selected_month}_{selected_year}"
    try:
        connection = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',
            database='allemployees',
            port=3306
        )
        if connection.is_connected():
            cursor = connection.cursor()
            create_table_query = f'''
            CREATE TABLE IF NOT EXISTS {table_name} (
                sr_no INT,
                name VARCHAR(255),
                uan_no VARCHAR(50),
                pf_no VARCHAR(50),
                esic_no VARCHAR(50),
                category VARCHAR(255),
                weekly_off VARCHAR(50),
                minimum_wages_per_day DECIMAL(10, 2),
                department VARCHAR(255),
                location VARCHAR(255),
                PanNo VARCHAR(50),
                aadhaar_no VARCHAR(50),
                join_date CHAR(10),
                date_of_birth CHAR(10),
                mobile_no VARCHAR(15),
                site_expenses DECIMAL(10, 2),
                present_days INT,
                gross_salary DECIMAL(10, 2),
                basic_da DECIMAL(10, 2),
                conveyance DECIMAL(10, 2),
                hra DECIMAL(10, 2),
                bonus DECIMAL(10, 2),
                pf_employee DECIMAL(10, 2),
                pf_employer DECIMAL(10, 2),
                admin_charges DECIMAL(10, 2),
                esic_employee DECIMAL(10, 2),
                esic_employer DECIMAL(10, 2),
                advance DECIMAL(10, 2),
                professional_tax DECIMAL(10, 2),
                total_deduction DECIMAL(10, 2),
                net_pay DECIMAL(10, 2),
                week_off_count INT
            )
            '''
            cursor.execute(create_table_query)
            connection.commit()  # Commit the changes
            print(f"Table '{table_name}' created successfully.")
            return table_name  # Return the table name if successful
    except mysql.connector.Error as e:
        print(f"Error creating table: {e}")
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()
    
    return None  # Return None if table creation fails

def employee_exists(uan_no, table_name):
    try:
        connection = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',  
            database='allemployees',
            port=3306
        )
        if connection.is_connected():
            cursor = connection.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {table_name} WHERE uan_no = %s;", (uan_no,))
            count = cursor.fetchone()[0]
            return count > 0
    except Error as e:
        print(f"Error checking if employee exists: {e}")
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()
    return False

def calculate_present_days(week_off, total_days, selected_month, selected_year, join_date=None):
    month_mapping = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4,
        'May': 5, 'June': 6, 'July': 7, 'August': 8,
        'September': 9, 'October': 10, 'November': 11, 'December': 12
    }    
    if isinstance(selected_month, str):
        selected_month = month_mapping[selected_month.capitalize()]
    
    selected_year = int(selected_year)
    week_off_days = [day.strip().lower() for day in week_off.split()]
    _, last_day = calendar.monthrange(selected_year, selected_month)

    if join_date:
        try:
            join_date = datetime.strptime(join_date, "%d-%m-%Y")
            # If the employee joined after the selected month, return 0
            if join_date.year > selected_year or (join_date.year == selected_year and join_date.month > selected_month):
                return 0
            
            # If the employee joined in the selected month, calculate days from join date
            elif join_date.year == selected_year and join_date.month == selected_month:
                start_day = join_date.day 
            else:
                start_day = 1  
        except ValueError:
            start_day = 1  
    else:
        start_day = 1  #

    present_days = 0
    for day in range(start_day, last_day + 1):
        date = datetime(selected_year, selected_month, day)
        weekday_name = date.strftime("%A").lower()
        if weekday_name not in week_off_days:
            present_days += 1
    return present_days

def calculate_week_days(week_off, selected_month, selected_year, join_date):
    join_date = datetime.strptime(join_date, "%d-%m-%Y")    
    _, last_day = calendar.monthrange(selected_year, selected_month)
    if join_date.year == selected_year and join_date.month == selected_month:
        first_day = join_date.day
    else:
        first_day = 1
    week_off_days = [day.strip().lower() for day in week_off.split()]
    month_weekdays = []
    for day in range(first_day, last_day + 1):
        date = datetime(selected_year, selected_month, day)
        month_weekdays.append(date.strftime("%A").lower())  
    unique_week_off_days = set(week_off_days)
    week_off_count = sum(1 for day in month_weekdays if day in unique_week_off_days)
    return week_off_count

def get_total_days_in_month(selected_month, selected_year):
    month_mapping = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4,
        'May': 5, 'June': 6, 'July': 7, 'August': 8,
        'September': 9, 'October': 10, 'November': 11, 'December': 12
    }    
    if isinstance(selected_month, str):
        selected_month = month_mapping[selected_month.capitalize()]    
    selected_year = int(selected_year)    
    return calendar.monthrange(selected_year, selected_month)[1]

def generate_pdf(employee_data, table_name, selected_month, selected_year):
    
    if not employee_data:
        messagebox.showwarning("Warning", "No employee data available to generate PDF.")
        return
    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=15)

        pdf.set_left_margin(10)  
        pdf.set_right_margin(10)
        pdf.add_page()
        pdf.set_font("helvetica", size=5, style='B')

        label_x1 = 10
        value_x1 = 30 
        colon_x1 = 28
        start_y = 10                         
        pdf.set_xy(label_x1, start_y)  
        pdf.cell(0, 2, text="Company Name", border=0) 
        pdf.set_xy(colon_x1, start_y) 
        pdf.cell(0, 2, text=":", border=0) 
        pdf.set_xy(value_x1, start_y)  
        pdf.cell(0, 2, text=" ", border=0)  
        pdf.ln(3)
        start_y += 3
        pdf.set_xy(label_x1, start_y)  
        pdf.cell(0, 2, text="Address", border=0) 
        pdf.set_xy(colon_x1, start_y)  
        pdf.cell(0, 2, text=":", border=0) 
        pdf.set_xy(value_x1, start_y)  
        pdf.cell(0, 2, text="", border=0) 
        pdf.ln(3)
        start_y += 3
        pdf.set_xy(label_x1, start_y) 
        pdf.cell(0, 2, text="Email ID", border=0) 
        pdf.set_xy(colon_x1, start_y) 
        pdf.cell(0, 2, text=":", border=0) 
        pdf.set_xy(value_x1, start_y) 
        pdf.cell(0, 2, text="", border=0) 
        pdf.ln(3)
        start_y += 3
        pdf.set_xy(label_x1, start_y)  
        pdf.cell(0, 2, text="Mobile No", border=0) 
        pdf.set_xy(colon_x1, start_y)
        pdf.cell(0, 2, text=":", border=0)  
        pdf.set_xy(value_x1, start_y)  
        pdf.cell(0, 2, text="", border=0) 
        pdf.ln(3)
        start_y += 3
        month_name = calendar.month_name[int(selected_month)]
        pdf.set_xy(label_x1, start_y) 
        pdf.cell(0, 2, text="Month", border=0) 
        pdf.set_xy(colon_x1, start_y) 
        pdf.cell(0, 2, text=":", border=0)  
        pdf.set_xy(value_x1, start_y) 
        pdf.cell(0, 2, text=f"{month_name} {selected_year}", border=0) 
        pdf.ln(5)
        
        headers = ['SR No', 'Full Name of Employee', 'UAN No', 'PF No', 'ESIC No', 'Category', 'Week Off', 'Present Days', 'Min Wages/Day', 'Gross Salary', 'Basic DA', 'Conveyance', 'HRA', 'Bonus', 'Site Expenses', 'PF Employee', 'PF Employer', 'Admin Charges', 'ESIC Employee', 'ESIC Employer', 'Advance', 'Professional Tax', 'Total Deduction', 'Net Pay']
        column_widths = [5.5, 28, 11, 20, 10, 15, 10, 10.6, 13, 10.5, 8, 11, 6, 7, 12, 11, 11, 12, 13, 13, 8, 14, 13, 8]

        if len(headers) != len(column_widths):
            raise ValueError(f"Number of headers ({len(headers)}) does not match number of column widths ({len(column_widths)}).")
        
        pdf.set_font("helvetica", size=4.5, style='B') 
        for i, header in enumerate(headers):
            pdf.cell(column_widths[i], 5, header, border=1, align='C')
        pdf.ln()

        pdf.set_font("helvetica", size=4) 
        for index, employee in enumerate(employee_data):
            pdf.cell(column_widths[0], 5, str(index + 1), border=1, align='C')
            for i, header in enumerate(headers[1:]):
                pdf.cell(column_widths[i + 1], 5, str(employee.get(header, '')), border=1, align='C')
            pdf.ln()

        pdf_file_path = f"F:\\Software\\Salary Sheet\\{table_name}.pdf"
        pdf.output(pdf_file_path)
        print(f"PDF generated successfully: {pdf_file_path}")
        messagebox.showinfo("Success", f"PDF saved as {pdf_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error generating PDF: {e}")
        
def insert_data_into_monthly_db(employees, table_name, selected_month, selected_year):
    try:
        connection = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',  
            database='allemployees',
            port=3306
        )
        if connection.is_connected():
            cursor = connection.cursor()
            print("Connection established.")
            total_days = get_total_days_in_month(selected_month, selected_year)
            employee_data = []  # Initialize a list to collect employee data
            successfully_inserted = 0
            successfully_updated = 0
            for employee in employees:
                try:
                    present_days = calculate_present_days(employee['Weekly Off'], total_days, selected_month, selected_year, employee['Join Date'])    
                    week_off_count = calculate_week_days(employee['Weekly Off'], selected_month, selected_year, employee['Join Date'])
                    gross_salary = round(float(employee['Minimum Wages Per Day']) * present_days, 2)
                    gross_salarysite = round(float(employee['Site Expenses']) + gross_salary, 2)                  
                    basic_da = round(0.6 * gross_salary, 2) 
                    conveyance = round(0.1 * gross_salary, 2)  # 10% of gross salary
                    hra = round(0.15 * gross_salary, 2)  # 15% of gross salary                 
                    basic_da_for_pf = basic_da if basic_da <= 15000 else 15000
                    pf_employee = round(0.12 * basic_da_for_pf, 2)  # 12% of Basic DA                    
                    pf_employer = round(0.12 * basic_da_for_pf, 2)  # 12% of Basic DA
                    admin_charges = round(0.01 * basic_da_for_pf, 2)  # 1% of Basic DA
                    esic_employee = round(0.0075 * gross_salary, 2)  # 0.75% of gross salary
                    esic_employer = round(0.0325 * gross_salary, 2)  # 3.25% of gross salary
                    advance = 0 # default 0
                    professional_tax = 200  # Fixed amount
                    bonus = gross_salary - (basic_da + conveyance + hra + pf_employee + esic_employee + professional_tax)
                    bonus = round(bonus,2)
                    total_deduction = round(pf_employee + esic_employee + professional_tax - advance, 2)
                    #net_pay = round(gross_salary - total_deduction, 2)
                    net_pay = round(gross_salarysite - total_deduction, 2)
                    print(f"Processing employee {employee['UAN No']}: {employee['Full Name of Employee']}")  # Debug: Print the employee being processed
                    if employee_exists(employee['UAN No'], table_name):
                        update_query = f'''
                            UPDATE {table_name}
                            SET 
                                sr_no = %s, name =%s, pf_no =%s, esic_no =%s, category =%s, weekly_off =%s, minimum_wages_per_day =%s, 
                                department =%s, location =%s, PanNo =%s, aadhaar_no =%s, join_date =%s, date_of_birth =%s, mobile_no =%s, site_expenses =%s,
                                present_days =%s, gross_salary =%s, basic_da =%s, conveyance =%s, hra =%s, bonus =%s, 
                                pf_employee =%s, pf_employer =%s, admin_charges =%s, esic_employee =%s, esic_employer =%s, 
                                advance =%s, professional_tax =%s, total_deduction =%s, net_pay =%s, week_off_count =%s
                            WHERE uan_no = %s
                        '''                        
                        cursor.execute(update_query, (
                            employee['SR No'],
                            employee['Full Name of Employee'],
                            employee['PF No'],
                            employee['ESIC No'],
                            employee['Category'],
                            employee['Weekly Off'],
                            employee['Minimum Wages Per Day'],
                            employee['Department'],
                            employee['Location'],
                            employee['PanNo'],
                            employee['aadhaar_no'],
                            employee['Join Date'],
                            employee['Date of Birth'],
                            employee['Mobile_No'],
                            employee['Site Expenses'],
                            present_days,
                            gross_salary,
                            basic_da,
                            conveyance,
                            hra,
                            bonus,
                            pf_employee,
                            pf_employer,
                            admin_charges,
                            esic_employee,
                            esic_employer,
                            advance,
                            professional_tax,
                            total_deduction,
                            net_pay,
                            week_off_count,
                            employee['UAN No'] 
                        ))
                        employee_data.append({
                            'SR No': employee['SR No'],
                            'Full Name of Employee': employee['Full Name of Employee'],
                            'UAN No': employee['UAN No'],
                            'PF No': employee['PF No'],
                            'ESIC No': employee['ESIC No'],
                            'Category': employee['Category'],
                            'Week Off': employee['Weekly Off'],
                            'Present Days': present_days,
                            'Min Wages/Day': employee['Minimum Wages Per Day'],
                            'Gross Salary': gross_salary,
                            'Basic DA': basic_da,
                            'Conveyance': conveyance,
                            'HRA': hra,
                            'Bonus': bonus,
                            'Site Expenses': employee['Site Expenses'],
                            'PF Employee': pf_employee,
                            'PF Employer': pf_employer,
                            'Admin Charges': admin_charges,
                            'ESIC Employee': esic_employee,
                            'ESIC Employer': esic_employer,
                            'Advance': advance,
                            'Professional Tax': professional_tax,
                            'Total Deduction': total_deduction,
                            'Net Pay': net_pay
                        })
                        successfully_updated += 1 
                        print(f"Updated employee {employee['UAN No']} successfully")                                    
                    else:
                        insert_query = f'''
                            INSERT INTO {table_name} (sr_no, name, uan_no, pf_no, esic_no, category, weekly_off, minimum_wages_per_day, 
                                department, location, PanNo, aadhaar_no, join_date, date_of_birth, mobile_no, site_expenses, 
                                present_days, gross_salary, basic_da, conveyance, hra, bonus, 
                                pf_employee, pf_employer, admin_charges, esic_employee, esic_employer, 
                                advance, professional_tax, total_deduction, net_pay, week_off_count)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        '''
                        cursor.execute(insert_query, (
                            employee['SR No'],
                            employee['Full Name of Employee'],
                            employee['UAN No'],
                            employee['PF No'],
                            employee['ESIC No'],
                            employee['Category'],
                            employee['Weekly Off'],
                            employee['Minimum Wages Per Day'],
                            employee['Department'],
                            employee['Location'],
                            employee['PanNo'],
                            employee['aadhaar_no'],
                            employee['Join Date'],
                            employee['Date of Birth'],
                            employee['Mobile_No'],
                            employee['Site Expenses'],
                            present_days,
                            gross_salary,
                            basic_da,
                            conveyance,
                            hra,
                            bonus,
                            pf_employee,
                            pf_employer,
                            admin_charges,
                            esic_employee,
                            esic_employer,
                            advance,
                            professional_tax,
                            total_deduction,
                            net_pay,
                            week_off_count
                        ))
                        employee_data.append({
                            'SR No': employee['SR No'],
                            'Full Name of Employee': employee['Full Name of Employee'],
                            'UAN No': employee['UAN No'],
                            'PF No': employee['PF No'],
                            'ESIC No': employee['ESIC No'],
                            'Category': employee['Category'],
                            'Week Off': employee['Weekly Off'],
                            'Present Days': present_days,
                            'Min Wages/Day': employee['Minimum Wages Per Day'],
                            'Gross Salary': gross_salary,
                            'Basic DA': basic_da,
                            'Conveyance': conveyance,
                            'HRA': hra,
                            'Bonus': bonus,
                            'Site Expenses': employee['Site Expenses'],
                            'PF Employee': pf_employee,
                            'PF Employer': pf_employer,
                            'Admin Charges': admin_charges,
                            'ESIC Employee': esic_employee,
                            'ESIC Employer': esic_employer,
                            'Advance': advance,
                            'Professional Tax': professional_tax,
                            'Total Deduction': total_deduction,
                            'Net Pay': net_pay
                        })
                        connection.commit() 
                        successfully_inserted += 1
                        print(f"Inserted employee {employee['SR No']} successfully.")
                except Error as insert_error:
                    print(f"Failed to insert employee {employee['Full Name of Employee']}: {insert_error}")
            connection.commit() 
            if successfully_inserted > 0 or successfully_updated > 0:
                generate_pdf(employee_data, table_name, selected_month, selected_year)
            messagebox.showinfo("Success", f"Updated {successfully_updated} employees and inserted {successfully_inserted} new employees into the monthly table '{table_name}'.")
    except Error as e:
        print(f"Database connection error: {e}")
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

def extract_data_from_pdf(pdf_path):
    """Extracts employee data from a PDF file."""
    extracted_data = []
    try:
        if not os.path.exists(pdf_path):
            print(f"PDF file not found at: {pdf_path}")
            return extracted_data
        reader = PdfReader(pdf_path)
        pattern = r'(\d+)\s+([\w\s]+)\s+(\d{12})\s+([A-Z0-9]+)\s+(\d{10})\s+([\w\s/]+)\s+(\w+)\s+(\d+(\.\d+)?)\s+([\w\s]+)\s+([\w\s]+)\s+([\w\s]+)\s+(\d{12})\s+([\d-]+)\s+([\d-]+)\s+([\+\d\s\(\)]+)\s+(\d+(\.\d+)?)'  # Minimum Wages now accepts decimals
        for page in reader.pages:
            text = page.extract_text() 
            if text:
                print("Extracted text from page:", text)  
                lines = text.splitlines()

                for line in lines:                  
                    print("Processing line:",line.strip())  
                    
                    match = re.match(pattern, line.strip() ) 
                    if match:
                        sr_no = match.group(1)
                        full_name_of_employee = match.group(2).strip()
                        uan_no = match.group(3)
                        pf_no = match.group(4)
                        esic_no = match.group(5)
                        category = match.group(6).strip()
                        Week_off = match.group(7).strip()
                        minimum_wages_per_day = match.group(8).strip()
                        department = match.group(10).strip()
                        location = match.group(11).strip()
                        pan_no = match.group(12).strip()
                        aadhaar_no = match.group(13).strip()
                        join_date = match.group(14)
                        date_of_birth = match.group(15)
                        mobile_no = match.group(16).strip()
                        site_expenses = match.group(17).strip() 
                        extracted_data.append({
                            "SR No": sr_no,
                            "Full Name of Employee": full_name_of_employee,
                            "UAN No": uan_no,
                            "PF No": pf_no,
                            "ESIC No": esic_no,
                            "Category": category,
                            "Weekly Off": Week_off,
                            "Minimum Wages Per Day": minimum_wages_per_day,  
                            "Department": department,
                            "Location": location,
                            "PanNo": pan_no,
                            "aadhaar_no": aadhaar_no,
                            "Join Date": join_date,
                            "Date of Birth": date_of_birth,
                            "Mobile_No": mobile_no,  
                            "Site Expenses": site_expenses
                        })
                    else:
                        print(f"No match found for line: {line.strip()}") 
        if extracted_data:
            print(f"Extracted Employees: {len(extracted_data)}") 
        else:
            print("No employees extracted.")

        return extracted_data
    except Exception as e:
        print(f"Failed to extract data: {e}")
        return extracted_data
pdf_path = 'path/to/your/file.pdf'
employees = extract_data_from_pdf(pdf_path)
print("Extracted Employees:", employees)

def load_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        employees = extract_data_from_pdf(pdf_path)
        print("Extracted Employees:", employees) 
        if employees:
            selected_month = month_var.get()  
            selected_year = int(year_var.get()) 
            month_names = list(calendar.month_name)[1:]  
            month_map = {name: i+1 for i, name in enumerate(month_names)} 
            selected_month_num = month_map[selected_month]  
            table_name = create_monthly_table(selected_month, selected_year)  
            if table_name:  
                insert_data_into_monthly_db(employees, table_name, selected_month_num, selected_year) 
            else:
                messagebox.showerror("Error", "Table creation failed. No data inserted.")
        else:
            messagebox.showwarning("No Data", "No employees extracted from the PDF.")
    else:
        messagebox.showwarning("File Not Selected", "Please select a PDF file.")

def generate_salary_slips(selected_month, selected_year, base_save_directory):
    month_folder_name = f"{selected_month} {selected_year}"  
    save_directory = os.path.join(base_save_directory, month_folder_name)    
    os.makedirs(save_directory, exist_ok=True)        
    table_name = f"employees_{selected_month}_{selected_year}"
    try:
        connection = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',  # Your MySQL password
            database='allemployees',
            port=3306
        )
        if connection.is_connected():
            cursor = connection.cursor()
            query = f"SELECT * FROM {table_name}"
            cursor.execute(query)
            employees = cursor.fetchall()            
            for employee in employees:
                (sr_no, name, uan_no, pf_no, esic_no, category, weekly_off, minimum_wages_per_day, 
                 department, location, pan_no, aadhaar_no, join_date, 
                 date_of_birth, Mobile_No, site_expenses, present_days, gross_salary, basic_da, conveyance, hra, bonus, 
                 pf_employee, pf_employer, admin_charges, esic_employee, esic_employer, advance, professional_tax, total_deduction, net_pay, week_off_count) = employee
                
                sanitized_name = re.sub(r'[<>:"/\\|?*]', '', name)  
                sanitized_mobile = re.sub(r'[<>:"/\\|?*]', '', Mobile_No)  
                slip_filename = os.path.join(save_directory, f"{sanitized_name}_{sanitized_mobile}_salaryslip_{selected_month}_{selected_year}.pdf")
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("helvetica", size=12)
                pdf.set_font("helvetica", style='B', size=18)
                pdf.cell(0, 10, text="Company Name", ln=True, align='C')
                pdf.set_font("helvetica", size= 12 )
                pdf.cell(0, 10, text="Address", ln=False, align='C')
                pdf.ln(8)                  
                pdf.set_font('helvetica', 'B', 10)  
                pdf.cell(0, 10, text=f"PAYSLIP FOR THE MONTH OF {month_folder_name}", ln=True, align='C')  
                pdf.ln(12)
                pdf.rect(5, 5, 200, 133)                  
                pdf.line(5, 35, 205, 35)
                label_x1 = 5
                value_x1 = 42  
                colon_x1 = 40
                start_y = 38  
                label_x2 = 110
                value_x2 = 147  
                colon_x2 = 145
                pdf.set_font("helvetica", size=10)                
                pdf.set_xy(label_x1, start_y)  
                pdf.cell(80, 5, text="Employee Name", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x1, start_y)  
                pdf.cell(30, 5, text=f"{name}", border=0) 
                pdf.set_xy(label_x2, start_y) 
                pdf.cell(80, 5, text="Employee Code", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x2, start_y)  
                pdf.cell(100, 5, text=f"{uan_no}", border=0) 
                pdf.ln(5)  
                start_y += 6
                pdf.set_xy(label_x1, start_y) 
                pdf.cell(80, 5, text="Designation", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0)  
                pdf.set_xy(value_x1, start_y) 
                pdf.cell(30, 5, text=f"{category}", border=0)  
                pdf.set_xy(label_x2, start_y)  
                pdf.cell(80, 5, text="Location", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0)  
                pdf.set_xy(value_x2, start_y) 
                pdf.cell(100, 5, text=f"{location}", border=0)  
                pdf.ln(5)   
                start_y += 6
                pdf.set_xy(label_x1, start_y) 
                pdf.cell(80, 5, text="Department", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x1, start_y) 
                pdf.cell(30, 5, text=f"{department}", border=0) 
                pdf.set_xy(label_x2, start_y)  
                pdf.cell(80, 5, text="Join Date", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x2, start_y)  
                pdf.cell(100, 5, text=f"{join_date}", border=0)  
                pdf.ln(5)   
                start_y += 6
                pdf.set_xy(label_x1, start_y) 
                pdf.cell(80, 5, text="PAN No", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x1, start_y) 
                pdf.cell(30, 5, text=f"{pan_no}", border=0) 
                pdf.set_xy(label_x2, start_y)  
                pdf.cell(80, 5, text="DOB", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x2, start_y)  
                pdf.cell(100, 5, text=f"{date_of_birth}", border=0)  
                pdf.ln(5) 
                start_y += 6
                pdf.set_xy(label_x1, start_y) 
                pdf.cell(80, 5, text="Aadhaar No", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x1, start_y) 
                pdf.cell(30, 5, text=f"{aadhaar_no}", border=0) 
                pdf.set_xy(label_x2, start_y)  
                pdf.cell(80, 5, text="Esic No", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x2, start_y)  
                pdf.cell(100, 5, text=f"{esic_no}", border=0)  
                pdf.ln(5)   
                start_y += 6
                pdf.set_xy(label_x1, start_y) 
                pdf.cell(80, 5, text="UAN", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x1, start_y) 
                pdf.cell(30, 5, text=f"{uan_no}", border=0)                    
                pdf.set_xy(label_x2, start_y)  
                pdf.cell(80, 5, text="Weekly Offs", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x2, start_y)  
                pdf.cell(100, 5, text=f"{week_off_count}", border=0)  
                pdf.ln(5)   
                start_y += 6
                pdf.set_xy(label_x1, start_y) 
                pdf.cell(80, 5, text="EPFO", border=0) 
                pdf.set_xy(colon_x1, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x1, start_y) 
                pdf.cell(30, 5, text=f"{pf_no}", border=0)                     
                pdf.set_xy(label_x2, start_y)  
                pdf.cell(80, 5, text="No. Of Working Days", border=0)
                pdf.set_xy(colon_x2, start_y)  
                pdf.cell(5, 5, text=":", border=0) 
                pdf.set_xy(value_x2, start_y)  
                pdf.cell(100, 5, text=f"{present_days}", border=0)                                       
                pdf.ln(8)   
                pdf.set_font("helvetica", style='B', size=12)
                header_y = 85  
                line_y_offset = 2 
                line_spacing = 0  
                pdf.line(5, header_y - line_y_offset, 205, header_y - line_y_offset) 
                pdf.cell(60, 10, text="Earnings", border=0, align='C', ln=False)
                pdf.cell(64, 10, text="Contributions", border=0, align='C', ln=False)
                pdf.cell(70, 10, text="Deductions", border=0, align='C', ln=False)
                pdf.ln(5)  
                pdf.line(5, header_y + 5, 205, header_y + 5)  
                pdf.ln(3)
                start_y = pdf.get_y()
                cell_height = 8 
                earnings = [
                    ("Basic + DA", basic_da),
                    ("Conveyance", conveyance),
                    ("HRA", hra),
                    ("Bonus", bonus),
                    ("Site Expenses", site_expenses)
                ]
               
                Contributions = [
                    ("PF Employer", pf_employer),
                    ("ESIC Employer", esic_employer),
                    ("Admin Charges", admin_charges),
                ]
                deductions = [
                    ("PF (Employee)", pf_employee),
                    ("ESIC (Employee)", esic_employee),
                    ("Advance", advance),
                    ("Professional Tax", professional_tax)
                ]
                pdf.set_font("helvetica", size=10)
                start_x = 5
                start_y = pdf.get_y()  
                cell_height = 8
                pdf.set_xy(start_x, start_y)
                line_start_x = start_x + 68  
                line_start_y = 83  
                line_length = 55 
                line_start_x2 = start_x + 128  
                line_start_y2 = 83  
                line_length = 55  
                pdf.line(line_start_x, line_start_y, line_start_x, line_start_y + line_length)
                pdf.line(line_start_x2, line_start_y2, line_start_x2, line_start_y2 + line_length)
                
                for i in range(max(len(earnings), len(Contributions), len(deductions))):
                    if i < len(earnings):                    
                        pdf.set_x(5)  
                        pdf.cell(40, cell_height, text=earnings[i][0], border=0, ln=False)  
                        pdf.set_x(40) 
                        pdf.cell(5, cell_height, text=":", border=0, ln=False) 
                        pdf.set_x(42)  
                        pdf.cell(40, cell_height, text=str(earnings[i][1]), border=0, ln=False)
                    else:
                        pdf.cell(100, cell_height, text="", border=0, ln=False)  
                        
                    if i < len(Contributions):
                        pdf.set_x(75)  
                        pdf.cell(40, cell_height, text=Contributions[i][0], border=0, ln=False)  
                        pdf.set_x(110) 
                        pdf.cell(5, cell_height, text=":", border=0, ln=False) 
                        pdf.set_x(112)  
                        pdf.cell(40, cell_height, text=str(Contributions[i][1]), border=0, ln=False)  
                    else:
                        pdf.cell(100, cell_height, text="", border=0, ln=False)

                    if i < len(deductions):
                        pdf.set_x(135)  
                        pdf.cell(40, cell_height, text=deductions[i][0], border=0, ln=False)  
                        pdf.set_x(175) 
                        pdf.cell(5, cell_height, text=":", border=0, ln=False) 
                        pdf.set_x(177)  
                        pdf.cell(40, cell_height, text=str(deductions[i][1]), border=0, ln=True)  
                    else:
                        pdf.cell(100, cell_height, text="", border=0, ln=True)

                current_y = 130
                pdf.line(5, current_y, 205, current_y)  
                pdf.set_xy(5, current_y)
                cell_height = 8  
                pdf.set_font('helvetica', 'B', 11)  
                pdf.set_x(5) 
                gross_salarys = gross_salary + site_expenses
                pdf.cell(40, cell_height, text="Gross Amount", border=0, ln=False)  
                pdf.set_x(40) 
                pdf.cell(5, cell_height, text=":", border=0, ln=False) 
                pdf.set_x(42)  
                pdf.cell(40, cell_height, text=str(gross_salarys), border=0, ln=False)  

                total_contribution = round(pf_employer + esic_employer +admin_charges, 2)
                pdf.set_x(75)  
                pdf.cell(40, cell_height, text="Total Contribution", border=0, ln=False)  
                pdf.set_x(110) 
                pdf.cell(5, cell_height, text=":", border=0, ln=False) 
                pdf.set_x(112)  
                pdf.cell(40, cell_height, text=str(total_contribution), border=0, ln=False)  
                pdf.set_x(135)  
                pdf.cell(40, cell_height, text="Total Deduction", border=0, ln=False)  
                pdf.set_x(175) 
                pdf.cell(5, cell_height, text=":", border=0, ln=False) 
                pdf.set_x(177)  
                pdf.cell(40, cell_height, text=str(-total_deduction), border=0, ln=False)               
                current_y += cell_height          
                pdf.ln(10)    
                current_y = pdf.get_y()
                pdf.set_font('helvetica', 'B', 12)  
                pdf.set_x(5) 
                pdf.cell(38, 8, text="Net Pay", border=1, ln=False)               
                pdf.set_x(43)  
                pdf.cell(62, 8, text=str(net_pay), border=1, align='C', ln=False)  
                pdf.ln(5)  
                line_y_position = current_y + 8 + 5  
                pdf.line(5, line_y_position, 205, line_y_position)  
                pdf.set_y(line_y_position) 
                note_x_position = 15  
                pdf.set_font('helvetica', size=10)  
                pdf.set_x(note_x_position)  
                pdf.cell(100, 10, text="Note: This is a computer-generated statement, it does not need any signature", border=0, align='C', ln=False)
                pdf.output(slip_filename)     
            messagebox.showinfo("Success", "Salary slips generated successfully!")
    except mysql.connector.Error as e:
        print(f"Error accessing data: {e}")
    finally:
        if 'cursor' in locals() and cursor is not None:
            cursor.close()
        if 'connection' in locals() and connection.is_connected():
            connection.close()

def load_salary_slips(selected_month, selected_year):
    base_save_directory = r'F:\Software\Salary Slips'  
    generate_salary_slips(selected_month, selected_year, base_save_directory)

def on_create_table():
    selected_month = month_var.get()  
    selected_year = year_var.get()  

    if selected_month and selected_year:  
        print(f"Creating table for month: {selected_month}, year: {selected_year}")
        create_monthly_table(selected_month, selected_year)
    else:
        messagebox.showwarning("Warning", "Please select both a month and a year.")

## start automatically sending via whatsapp
def send_pdf_via_whatsapp(contact_number, file_path, prefix='+91'):
    global driver
    try:
        search_box = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='3']"))
        )
        search_box.click()
        search_box.clear()
        full_contact_number = f"{prefix} {contact_number}"
        search_box.send_keys(full_contact_number)
        print(f"Searching for contact {full_contact_number}...")
        time.sleep(3) 
        contacts = WebDriverWait(driver, 40).until(
            EC.presence_of_all_elements_located((By.XPATH, "//span[@title]"))
        )
        if contacts:
            first_contact = contacts[0]
            first_contact.click()
            print(f"Clicked on contact: {first_contact.get_attribute('title')}")
            attachment_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@title='Attach']"))
            )
            attachment_button.click()
            time.sleep(1)
            file_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
            file_input.send_keys(os.path.abspath(file_path))
            print(f"Uploading file '{file_path}'...")
            time.sleep(5)
            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
            )
            send_button.click()
            print("PDF sent successfully!")
            time.sleep(2)
            print(f"No contacts found for {full_contact_number}.")
    except Exception as e:
        print(f"Error during sending PDF to +{contact_number}: {e}")

def send_pdfs_in_folder(folder_path):
    global driver
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
    for pdf_file in pdf_files:
        file_path = os.path.join(folder_path, pdf_file)
        try:
            contact_number = pdf_file.split('_')[1].replace('.pdf', '')
        except IndexError:
            print(f"Invalid file name format: {pdf_file}. Skipping...")
            continue

        print(f"Sending file: {file_path} to contact: {contact_number}")
        send_pdf_via_whatsapp(contact_number, file_path)

def initialize_driver():
    global driver
    driver_path = "C:\\chromedriver-win64\\chromedriver.exe" 
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://web.whatsapp.com/")
    print("Waiting for QR code scan...")
    WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='chats-filled']")))
    print("QR code scanned successfully!")
    
def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        initialize_driver()
        send_pdfs_in_folder(folder_path)
        driver.quit()  

def setup_gui():
    global month_var, year_var 
    root = tk.Tk()
    root.title("Salary Generator")
    root.geometry("600x600")
    root.config(bg="white")      
    subTitle=tk.Label(root,text="Please Select Month and Year",font=("times new roman",20,"bold"),bg="White",fg="black")
    subTitle.pack(pady=(50,10))    
    style = ttk.Style()
    style.configure("TCombobox", font=("times new roman", 14), background="white", foreground="black")
    style.configure("TCombobox.Listbox", height=5)  
    selection_frame = tk.Frame(root, bg="white")
    selection_frame.pack(pady=(20,155))
    month_names = list(calendar.month_name)[1:] 
    current_month = datetime.now().month
    month_var = tk.StringVar(value=month_names[current_month - 1])  
    month_label = tk.Label(selection_frame, text="Select Month:", bg="white", font=("times new roman", 16))
    month_label.grid(row=0, column=0, padx=5)
    month_menu = ttk.Combobox(
        selection_frame,
        textvariable=month_var,
        values=month_names,
        width=15,
        state='readonly',
        style="TCombobox"
    )
    month_menu.grid(row=0, column=1, padx=5, pady=5)
    month_menu['font'] = ("times new roman", 16)
    current_year = datetime.now().year
    year_var = tk.StringVar(value=str(current_year))  
    years = [str(year) for year in range(2024, current_year + 1)]  
    year_label = tk.Label(selection_frame, text="Select Year:", bg="white", font=("times new roman", 16))
    year_label.grid(row=0, column=2, padx=5)  
    year_menu = ttk.Combobox(
        selection_frame,
        textvariable=year_var,
        values=years,
        width=10,
        state='readonly',
        style="TCombobox"
    )
    year_menu.grid(row=0, column=3, padx=5)  
    year_menu['font'] = ("times new roman", 16)
    load_button = tk.Button(root, text="Load PDF", command=load_pdf, bg="black",fg="White", width=20, height=1, font=("times new roman",18))
    load_button.pack(pady=(10,5))    
    load_salary_button = tk.Button(root, text="Generate Salary Slip", command=lambda: load_salary_slips(month_var.get(), year_var.get()), bg="blue", fg="White", width=20, height=1, font=("times new roman",18))
    load_salary_button.pack(pady=(40,5))
    select_folder_button = tk.Button(root, text="Send Salary slip", command=select_folder, bg="Green" , fg="White",width=20, height=1, font=("times new roman",18 ))
    select_folder_button.pack(pady=(40,5))
    root.mainloop()

if __name__ == "__main__":
    setup_gui()


