import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from getpass import getpass
import schedule
import time
import logging
import win32com.client
import os
import sys
from datetime import datetime
import pandas as pd

# Configure logging with both file and console output
log_filename = f"email_automation_{datetime.now().strftime('%Y%m%d')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler(sys.stdout)
    ]
)

# Email credentials and recipients
email_address = None
email_password = None
recipients = ["example1@gmail.com", "example12@gmail.com"]

# File paths
excel_file = r"D:\Temp\Cleaner NE Report.xlsx"
script_dir = os.path.dirname(os.path.abspath(__file__))
temp_chart_path = os.path.join(script_dir, "temp_chart.png")

def cleanup_excel(excel=None):
    """Helper function to properly clean up Excel instances"""
    try:
        if excel:
            excel.DisplayAlerts = False
            excel.Quit()
            del excel
    except:
        pass

def get_table_html(ws):
    """Extract table data and convert to HTML"""
    try:
        logging.info("Detecting table range in Sources sheet...")
        
        # Find the last row with data in column A
        last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
        # Find the last column with data in row 1
        last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column
        
        # Get the range including headers
        table_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
        logging.info(f"Detected range: {table_range.Address}")
        
        # Convert range to list of lists
        data = []
        for row in table_range.Rows:
            row_data = []
            for cell in row.Cells:
                row_data.append(cell.Text)
            data.append(row_data)
        
        # Convert to pandas DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])  # First row as headers
        
        # Style the HTML table
        table_html = df.to_html(index=False, classes='styled-table')
        
        # Add CSS styling
        table_style = """
        <style>
        .styled-table {
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 0.9em;
            font-family: Arial, sans-serif;
            min-width: 400px;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
        }
        .styled-table thead tr {
            background-color: #009879;
            color: #ffffff;
            text-align: left;
        }
        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
            border: 1px solid #dddddd;
        }
        .styled-table tbody tr {
            border-bottom: 1px solid #dddddd;
        }
        .styled-table tbody tr:nth-of-type(even) {
            background-color: #f3f3f3;
        }
        .styled-table tbody tr:last-of-type {
            border-bottom: 2px solid #009879;
        }
        </style>
        """
        
        return table_style + table_html
            
    except Exception as e:
        logging.error(f"Error creating table HTML: {str(e)}")
        return ""

def refresh_excel_and_save_chart():
    excel = None
    wb = None
    table_html = ""
    try:
        logging.info("Initializing Excel application...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        logging.info(f"Opening workbook: {excel_file}")
        wb = excel.Workbooks.Open(excel_file)
        
        # Refresh all data connections
        logging.info("Refreshing all data connections...")
        for connection in wb.Connections:
            try:
                connection.Refresh()
                logging.info(f"Refreshed connection: {connection.Name}")
            except Exception as e:
                logging.warning(f"Failed to refresh connection {connection.Name}: {str(e)}")
        
        # Add delay to ensure refresh completes
        logging.info("Waiting for refresh to complete...")
        time.sleep(10)

        # Calculate the entire workbook
        logging.info("Calculating workbook...")
        wb.Application.Calculate()
        
        # First get the table from Sources sheet
        logging.info("Accessing Sources sheet...")
        ws_sources = wb.Worksheets("Sources")
        table_html = get_table_html(ws_sources)
        
        # Then get the chart from Sheet2
        logging.info("Accessing Sheet2 for chart...")
        ws_chart = wb.Worksheets("Sheet2")
        ws_chart.Activate()
        
        # Export chart
        logging.info("Checking for charts...")
        chart_objects = ws_chart.ChartObjects()
        
        if chart_objects.Count > 0:
            chart = chart_objects(1)
            logging.info(f"Found chart: {chart.Name}")
            
            os.makedirs(os.path.dirname(temp_chart_path), exist_ok=True)
            
            if os.path.exists(temp_chart_path):
                os.remove(temp_chart_path)
            
            logging.info(f"Exporting chart to {temp_chart_path}")
            chart.Chart.Export(Filename=temp_chart_path, FilterName="PNG")
            
            time.sleep(2)
            if os.path.exists(temp_chart_path):
                file_size = os.path.getsize(temp_chart_path)
                logging.info(f"Chart exported successfully. File size: {file_size} bytes")
                if file_size == 0:
                    logging.error("Exported chart file is empty")
                    return False, ""
            else:
                logging.error("Chart file was not created")
                return False, ""
        else:
            logging.error("No charts found in Sheet2")
            return False, ""
        
        # Save and close properly
        wb.Save()
        wb.Close()
        return True, table_html

    except Exception as e:
        logging.error(f"Error in refresh_excel_and_save_chart: {str(e)}")
        return False, ""
    
    finally:
        cleanup_excel(excel)

def send_reminder():
    logging.info("Starting send_reminder process...")

    try:
        # Refresh Excel and save chart
        success, table_html = refresh_excel_and_save_chart()
        if not success:
            logging.error("Failed to prepare chart and table for email")
            return

        # Email details
        subject = f"Network Elements Distribution Report Testing - {datetime.now().strftime('%Y-%m-%d')}"
        body = """
        Hi,

        Please find below the latest Network Elements distribution report.
        This report is automatically generated from the latest data in Power Query.

        Regards,
        Automated Reporting System
        """

        # SMTP server details for Outlook
        smtp_server = "smtp.office365.com"
        smtp_port = 587

        # Create the email
        msg = MIMEMultipart('related')
        msg["From"] = email_address
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject

        # Create HTML version of the message
        html = f"""
        <html>
            <body>
                <p>{body}</p>
                <img src="cid:chart">
                <br>
                {table_html}
            </body>
        </html>
        """
        
        msg.attach(MIMEText(html, 'html'))

        # Attach the chart
        if os.path.exists(temp_chart_path):
            with open(temp_chart_path, 'rb') as f:
                img_data = f.read()
                if len(img_data) > 0:
                    img = MIMEImage(img_data)
                    img.add_header('Content-ID', '<chart>')
                    msg.attach(img)
                    logging.info("Chart attached to email successfully")
                else:
                    logging.error("Chart file is empty")
                    return
        else:
            logging.error("Chart file not found for email attachment")
            return

        # Send the email
        logging.info("Attempting to send email...")
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_address, email_password)
            server.sendmail(email_address, recipients, msg.as_string())
            logging.info(f"Email successfully sent to {', '.join(recipients)}")

    except Exception as e:
        logging.error(f"Error in send_reminder: {str(e)}")
    
    finally:
        # Clean up temporary file
        try:
            if os.path.exists(temp_chart_path):
                os.remove(temp_chart_path)
                logging.info("Temporary chart file removed")
        except Exception as e:
            logging.error(f"Failed to remove temporary file: {str(e)}")

def schedule_emails():
    schedule.every().thursday.at("18:00").do(send_reminder)
    logging.info("Email schedule initialized for every Thursday at 18:00")

    while True:
        try:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
        except Exception as e:
            logging.error(f"Error in scheduler: {str(e)}")
            time.sleep(60)  # Wait a minute before retrying

def test_email_setup():
    """Test email setup and connectivity"""
    try:
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        
        logging.info("Testing SMTP connection...")
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_address, email_password)
            logging.info("SMTP authentication successful")
        return True
    except Exception as e:
        logging.error(f"SMTP test failed: {str(e)}")
        return False

if __name__ == "__main__":
    try:
        logging.info("=== Script Started ===")
        
        # Get credentials
        email_address = input("Enter your Outlook email address: ")
        email_password = getpass("Enter your email password (hidden): ")
        
        # Test email setup
        if not test_email_setup():
            logging.error("Email setup test failed. Exiting...")
            sys.exit(1)
        
        # Send initial test email
        logging.info("Sending test email...")
        send_reminder()
        
        # Start scheduler
        logging.info("Starting scheduler...")
        schedule_emails()
        
    except KeyboardInterrupt:
        logging.info("Script terminated by user")
    except Exception as e:
        logging.error(f"Unexpected error: {str(e)}")
    finally:
        logging.info("=== Script Ended ===")
