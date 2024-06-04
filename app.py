import time
import json
import os
import tkinter as tk
from ttkbootstrap import Style
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait


class LMSAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("LMS Assitant - BUKC")
        self.style = Style(theme="darkly")

        self.main_font = ("Helvetica", 12)
        self.pad = {"padx": 10, "pady": 10}
        self.pad = {"padx": 10, "pady": 10}

        # Main window
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        header_label_main = ttk.Label(self.main_frame, text="LMS Genie(Bukc) - Your Personal Assistant",
                                      style="primary.Inverse.TLabel", font=("Roboto", 16))
        header_label_main.pack(**self.pad)

        u1_frame = ttk.LabelFrame(self.main_frame, text="User Credentials", padding="10")
        u1_frame.pack(fill=tk.X, **self.pad)

        self.enrollment_label = ttk.Label(u1_frame, text="Enrollment ID:", font=self.main_font)
        self.enrollment_label.grid(row=0, column=0, **self.pad, sticky=tk.W)
        self.enrollment_entry = ttk.Entry(u1_frame, font=self.main_font)
        self.enrollment_entry.grid(row=0, column=1, **self.pad)

        self.password_label = ttk.Label(u1_frame, text="Password:", font=self.main_font)
        self.password_label.grid(row=1, column=0, **self.pad, sticky=tk.W)
        self.password_entry = ttk.Entry(u1_frame, show="*", font=self.main_font)
        self.password_entry.grid(row=1, column=1, **self.pad)

        self.institute_label = ttk.Label(u1_frame, text="Institute:", font=self.main_font)
        self.institute_label.grid(row=2, column=0, **self.pad, sticky=tk.W)
        self.var_inst = tk.StringVar()
        self.inst_dropdown = ttk.Combobox(u1_frame, textvariable=self.var_inst, state="readonly",
                                          font=self.main_font)
        self.inst_dropdown['values'] = [
            "Select",
            "Finishing School",
            "Health Sciences Campus (Karachi)",
            "IPP (Karachi)",
            "Islamabad E-8 Campus",
            "Islamabad H-11 Campus",
            "Karachi Campus",
            "Lahore Campus",
            "NCMPR",
            "PN Nursing College",
            "PN School Of Logistics"
        ]
        self.inst_dropdown.grid(row=2, column=1, **self.pad)
        self.inst_dropdown.current(0)

        # Frame for course selection
        courses_frame = ttk.LabelFrame(self.main_frame, text="Course Selection", padding="10")
        courses_frame.pack(fill=tk.X, **self.pad)

        self.course_label = ttk.Label(courses_frame, text="Select Course:", font=self.main_font)
        self.course_label.grid(row=0, column=0, **self.pad, sticky=tk.W)
        self.course_var = tk.StringVar()
        self.course_dropdown = ttk.Combobox(courses_frame, textvariable=self.course_var, state="readonly",
                                            font=self.main_font)
        self.course_dropdown.grid(row=0, column=1, **self.pad)

        self.load_courses_button = ttk.Button(courses_frame, text="Load Courses", command=self.load_courses)
        self.load_courses_button.grid(row=1, column=0, **self.pad)
        self.refresh_courses_button = ttk.Button(courses_frame, text="Refresh Courses", command=self.refresh_courses)
        self.refresh_courses_button.grid(row=1, column=1, **self.pad)

        action_frame = ttk.LabelFrame(self.main_frame, text="Actions", padding="10")
        action_frame.pack(fill=tk.X, **self.pad)

        self.download_button = ttk.Button(action_frame, text="Download Lectures", command=self.download_lectures)
        self.download_button.grid(row=0, column=0, **self.pad)

        self.attendance_button = ttk.Button(action_frame, text="View Attendance", command=self.view_attendance)
        self.attendance_button.grid(row=0, column=1, **self.pad)

        self.deadlines_button = ttk.Button(action_frame, text="Check All Deadlines", command=self.check_deadlines)
        self.deadlines_button.grid(row=1, column=0, **self.pad)

        self.reset_button = ttk.Button(action_frame, text="Reset App", command=self.reset_app)
        self.reset_button.grid(row=1, column=1, **self.pad)

        self.status_label = ttk.Label(self.main_frame, text="Status: Ready", font=self.main_font)
        self.status_label.pack(**self.pad)

        self.data_file = "user_data.json"
        self.load_user_data()

    def save_user_data(self):
        std_data = {
            "enrollment": self.enrollment_entry.get(),
            "password": self.password_entry.get(),
            "institute": self.var_inst.get(),
            "courses": self.course_dropdown['values']
        }
        with open(self.data_file, "w") as file:
            json.dump(std_data, file)

    def load_user_data(self):
        if os.path.exists(self.data_file):
            with open(self.data_file, "r") as file:
                std_data = json.load(file)
                self.enrollment_entry.insert(0, std_data.get("enrollment", ""))
                self.password_entry.insert(0, std_data.get("password", ""))
                self.var_inst.set(std_data.get("institute", "Select"))
                self.course_dropdown['values'] = std_data.get("courses", [])

    def login_to_cms(self, driver, enrollment, password, institute):
        driver.get("https://cms.bahria.edu.pk/Logins/Student/Login.aspx")
        enrollment_input = driver.find_element(By.ID, "BodyPH_tbEnrollment")
        enrollment_input.send_keys(enrollment)
        password_input = driver.find_element(By.ID, "BodyPH_tbPassword")
        password_input.send_keys(password)
        institute_select = driver.find_element(By.ID, "BodyPH_ddlInstituteID")
        institute_select.send_keys(institute)
        driver.find_element(By.ID, "BodyPH_btnLogin").click()

    def load_courses(self):
        if self.course_dropdown['values']:
            messagebox.showinfo("Courses", "Courses are already loaded.")
            return

        enrollment = self.enrollment_entry.get()
        password = self.password_entry.get()
        institute = self.var_inst.get()

        if not all([enrollment, password, institute]) or institute == "Select":
            messagebox.showerror("Error", "Please fill in all fields and select an institute.")
            return

        self.set_status("Starting Chrome browser...")
        chrome_options = Options()
        chrome_options.add_argument("start-maximized")
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_options)

        try:
            self.login_to_cms(driver, enrollment, password, institute)
            self.set_status("Navigating to Assignments page...")
            driver.get("https://cms.bahria.edu.pk/Sys/Common/GoToLMS.aspx")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get("https://lms.bahria.edu.pk/Student/Assignments.php")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "courseId"))
            )

            course_dropdown = driver.find_element(By.ID, "courseId")
            options = course_dropdown.find_elements(By.TAG_NAME, "option")
            courses = [option.text.strip() for option in options if option.text.strip() != "Select Course"]
            self.course_dropdown['values'] = courses
            self.set_status("Courses updated.")
            self.save_user_data()
        except Exception as e:
            self.set_status("Error occurred while loading courses.")
            messagebox.showerror("Error", str(e))
        finally:
            driver.quit()

    def refresh_courses(self):
        self.load_courses()

    def download_lectures(self):
        enrollment = self.enrollment_entry.get()
        password = self.password_entry.get()
        institute = self.var_inst.get()
        selected_course = self.course_var.get()

        if not all([enrollment, password, institute, selected_course]):
            messagebox.showerror("Error", "Please fill in all fields.")
            return

        self.set_status("Starting Chrome browser...")
        chrome_opt = Options()
        chrome_opt.add_argument("start-maximized")
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_opt)

        try:
            self.login_to_cms(driver, enrollment, password, institute)
            driver.get("https://cms.bahria.edu.pk/Sys/Common/GoToLMS.aspx")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get("https://lms.bahria.edu.pk/Student/LectureNotes.php")
            select_course_dropdown = driver.find_element(By.ID, "courseId")
            select_course_dropdown.send_keys(selected_course)
            time.sleep(5)  # Wait for lectures to load
            download_elements = driver.find_elements(By.XPATH, "//td[contains(., 'Download Lecture')]/a")
            for element in download_elements:
                download_link = element.get_attribute("href")
                if download_link:
                    driver.execute_script("window.open('" + download_link + "', '_blank');")
                    time.sleep(2)
            self.set_status("Lectures downloaded.")
            messagebox.showinfo("Success", "Lectures downloaded successfully.")
        except Exception as e:
            self.set_status("Error occurred.")
            messagebox.showerror("Error", str(e))
        finally:
            driver.quit()

    def view_attendance(self):
        enrollment = self.enrollment_entry.get()
        password = self.password_entry.get()
        institute = self.var_inst.get()

        if not all([enrollment, password, institute]) or institute == "Select":
            messagebox.showerror("Error", "Please fill in all fields and select an institute.")
            return

        self.set_status("Starting Chrome browser...")
        chrome_options = Options()
        chrome_options.add_argument("start-maximized")
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_options)

        try:
            self.login_to_cms(driver, enrollment, password, institute)
            # attendance_link = WebDriverWait(driver, 10).until(
            #     EC.element_to_be_clickable(
            #         (By.CSS_SELECTOR,
            #          'a.list-group-item[href="https://cms.bahria.edu.pk/Sys/Student/ClassAttendance/StudentWiseAttendance.aspx"]'))
            # )
            # attendance_link.click()
            #
            driver.get("https://cms.bahria.edu.pk/Sys/Student/ClassAttendance/StudentWiseAttendance.aspx")

            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(10)
            self.set_status("Attendance viewed.")
            messagebox.showinfo("Success", "Attendance viewed successfully.")
        except Exception as e:
            self.set_status("Error occurred.")
            messagebox.showerror("Error", str(e))
        finally:
            driver.quit()

    def check_deadlines(self):
        enrollment = self.enrollment_entry.get()
        password = self.password_entry.get()
        institute = self.var_inst.get()

        if not all([enrollment, password, institute]) or institute == "Select":
            messagebox.showerror("Error", "Please fill in all fields and select an institute.")
            return

        self.set_status("Starting Chrome browser...")
        chrome_opt = Options()
        chrome_opt.add_argument("start-maximized")
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_opt)

        try:
            self.login_to_cms(driver, enrollment, password, institute)
            driver.get("https://cms.bahria.edu.pk/Sys/Common/GoToLMS.aspx")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get("https://lms.bahria.edu.pk/Student/Assignments.php")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "courseId"))
            )

            for crs in self.course_dropdown['values']:
                select_course_dropdown = driver.find_element(By.ID, "courseId")
                select_course_dropdown.send_keys(crs)
                time.sleep(5)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                deadlines_elements = driver.find_elements(By.XPATH, "//td[contains(text(), 'Due Date')]")
                for deadline in deadlines_elements:
                    print(f"Course: {crs} - Deadline: {deadline.text}")

            self.set_status("Deadlines retrieved.")
            messagebox.showinfo("Success", "Deadlines retrieved successfully.")
        except Exception as e:
            self.set_status("Error occurred.")
            messagebox.showerror("Error", str(e))
        finally:
            driver.quit()

    def reset_app(self):
        self.enrollment_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)
        self.var_inst.set("Select")
        self.course_dropdown.set('')
        self.course_dropdown['values'] = []
        if os.path.exists(self.data_file):
            os.remove(self.data_file)
        self.set_status("App reset.")

    def set_status(self, message):
        self.status_label.config(text=f"Status: {message}")


if __name__ == "__main__":
    root = tk.Tk()
    app = LMSAssistant(root)
    root.mainloop()
