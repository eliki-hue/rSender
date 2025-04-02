import os
import csv
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from tkinter import Tk, filedialog, messagebox, simpledialog  # Added simpledialog import

class ReportSender:
    def __init__(self):
        self.setup_logging()
        
    def setup_logging(self):
        logging.basicConfig(
            filename='report_sender.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
    def load_students(self, csv_path):
        """Improved CSV loading with error handling"""
        students = []
        try:
            with open(csv_path, mode='r', encoding='utf-8-sig') as file:
                # Detect CSV format
                sample = file.read(1024)
                file.seek(0)
                dialect = csv.Sniffer().sniff(sample)
                reader = csv.DictReader(file, dialect=dialect)
                
                # Verify required columns
                required = {'student_name', 'parent_email'}
                if not required.issubset(reader.fieldnames):
                    missing = required - set(reader.fieldnames)
                    raise ValueError(f"Missing columns: {missing}")
                
                # Process rows
                for row in reader:
                    # Clean and validate data
                    row = {k: v.strip() if v else '' for k, v in row.items()}
                    if not row['student_name'] or not row['parent_email']:
                        continue
                    students.append(row)
                    
            logging.info(f"Loaded {len(students)} students from {csv_path}")
            return students
            
        except Exception as e:
            logging.error(f"Failed to load CSV: {str(e)}")
            messagebox.showerror("CSV Error", 
                f"Could not read student data:\n\n{str(e)}\n\n"
                "Please check:\n"
                "- File is a valid CSV\n"
                "- Contains 'student_name' and 'parent_email' columns\n"
                "- No empty rows in data")
            return None

    def find_pdf(self, pdf_dir, student_name):
        """Flexible PDF finding with multiple patterns"""
        patterns = [
            f"{student_name}.pdf",
            f"{student_name.replace(' ', '_')}.pdf",
            f"{student_name.replace(' ', '')}.pdf",
            f"Report_{student_name}.pdf",
            f"{student_name.split()[-1]}_{student_name.split()[0]}.pdf",  # Doe_John.pdf
            f"{student_name[0]}{student_name.split()[-1]}.pdf"  # JDoe.pdf
        ]
        
        for pattern in patterns:
            pdf_path = os.path.join(pdf_dir, pattern)
            if os.path.exists(pdf_path):
                return pdf_path
        return None

    def send_reports(self):
        """Main execution flow"""
        try:
            root = Tk()
            root.withdraw()
            
            # 1. Select CSV file
            csv_path = filedialog.askopenfilename(
                title="Select Student Data CSV",
                filetypes=[("CSV Files", "*.csv")])
            if not csv_path:
                return
                
            # 2. Select PDF directory
            pdf_dir = filedialog.askdirectory(title="Select PDF Reports Folder")
            if not pdf_dir:
                return
                
            # 3. Load student data
            students = self.load_students(csv_path)
            if not students:
                return
                
            # 4. Verify PDFs exist
            pdf_matches = []
            missing = []
            
            for student in students:
                pdf_path = self.find_pdf(pdf_dir, student['student_name'])
                if pdf_path:
                    pdf_matches.append((student, pdf_path))
                else:
                    missing.append(student['student_name'])
            
            # 5. Show preview
            if not pdf_matches:
                messagebox.showerror("No Matches", 
                    "Could not match any PDFs to students!\n\n"
                    f"Missing PDFs for: {', '.join(missing[:5])}..."
                    f"\n\nTotal missing: {len(missing)}")
                return
                
            # 6. Get email credentials
            email = simpledialog.askstring("Email", "Your email address:")
            if not email:
                return
                
            password = simpledialog.askstring("Password", "Your email password:", show='*')
            if not password:
                return
                
            smtp_server = simpledialog.askstring("SMTP", "SMTP server:", 
                initialvalue="smtp.gmail.com")
            if not smtp_server:
                return
                
            smtp_port = simpledialog.askinteger("Port", "SMTP port:", 
                initialvalue=587)
            if not smtp_port:
                return
                
            # 7. Send emails
            sent_count = 0
            try:
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(email, password)
                
                for student, pdf_path in pdf_matches:
                    try:
                        msg = MIMEMultipart()
                        msg['From'] = email
                        msg['To'] = student['parent_email']
                        msg['Subject'] = f"Report Card: {student['student_name']}"
                        
                        body = f"""Dear Parent/Guardian,
                        
Please find attached the report card for {student['student_name']}.

Best regards,
[Your School]"""
                        
                        msg.attach(MIMEText(body, 'plain'))
                        
                        with open(pdf_path, "rb") as f:
                            attach = MIMEApplication(f.read(), _subtype="pdf")
                            attach.add_header('Content-Disposition', 'attachment', 
                                            filename=os.path.basename(pdf_path))
                            msg.attach(attach)
                        
                        server.sendmail(email, student['parent_email'], msg.as_string())
                        sent_count += 1
                        logging.info(f"Sent to {student['parent_email']}")
                        
                    except Exception as e:
                        logging.error(f"Failed to send to {student['parent_email']}: {str(e)}")
                
                server.quit()
                messagebox.showinfo("Complete", 
                    f"Successfully sent {sent_count} of {len(students)} reports\n\n"
                    f"Failed to send: {len(students) - sent_count}")
                    
            except Exception as e:
                logging.error(f"SMTP Error: {str(e)}")
                messagebox.showerror("Email Error", 
                    f"Failed to send emails:\n\n{str(e)}")
                
        except Exception as e:
            logging.error(f"Application error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")

if __name__ == "__main__":
    app = ReportSender()
    app.send_reports()