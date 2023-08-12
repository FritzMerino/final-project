import os
import csv
import subprocess
from PyQt6.QtWidgets import *
from gui import *


class Logic(QMainWindow, Ui_MainWindow):
    """
    Class that handles the logic of the student grade management application.
    Inherits from QMainWindow and the UI class Ui_MainWindow
    """
    def __init__(self) -> None:
        """
        Initialize the application and connect buttons to their respective functions.
        :return: None
        """
        super().__init__()
        self.setupUi(self)

        self.button_add.clicked.connect(self.add_student)
        self.button_rem.clicked.connect(self.rem_student)
        self.button_edit.clicked.connect(self.edit_student)
        self.button_clear.clicked.connect(self.clear_students)
        self.button_sum.clicked.connect(self.sum_students)
        self.button_open_csv.clicked.connect(self.open_csv)

    @staticmethod
    def show_empty_list_error() -> None:
        """
        Display a warning message when the student list is empty, indicating that students need to be added.
        :return: None
        """
        error_message = QMessageBox()
        error_message.setIcon(QMessageBox.Icon.Warning)
        error_message.setWindowTitle("Empty list")
        error_message.setText('The student list is empty. Please add students.')
        error_message.exec()

    def add_student(self) -> None:
        """
        Add a new student's name and score within the list
        :return: None
        """
        try:
            name, ok = QInputDialog.getText(self, 'New Student', 'Student Name')
            if ok and name:
                # Display a warning that an existing name is already in the list
                existing_names = [self.list_students.item(row).text().split(":")[0].strip() for row in
                                  range(self.list_students.count())]
                if name.capitalize() in existing_names:
                    error_message = QMessageBox()
                    error_message.setIcon(QMessageBox.Icon.Warning)
                    error_message.setWindowTitle("Duplicate Name")
                    error_message.setText('Student with this name already exists')
                    error_message.exec()
                else:
                    score_input, ok = QInputDialog.getInt(self, 'New Student', 'Student Score')
                    if ok:
                        if self.list_students.count() == 0:
                            self.list_students.addItem('NAMES \t\tSCORES')

                        item_text = f'{name.capitalize()}: \t\t{score_input}'
                        self.list_students.addItem(item_text)
        except Exception as e:
            print(f"An error occurred: {e}")

    def rem_student(self) -> None:
        """
        Remove the name and score of a student selected based on the current index within the list.
        :return: None
        """
        try:
            if self.list_students.count() == 0:
                self.show_empty_list_error()
            else:
                current_index = self.list_students.currentRow()
                item = self.list_students.item(current_index)
                if item is not None and current_index != 0:
                    name = item.text().split(':')[0].strip()
                    question = QMessageBox.question(
                        self, 'Remove student',
                        f'Do you want to remove {name}?',
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )

                    if question == QMessageBox.StandardButton.Yes:
                        self.list_students.takeItem(current_index)
        except Exception as e:
            print(f"An error occurred: {e}")

    def edit_student(self) -> None:
        """
        Edit the name and score of a student selected based on the current index within the list.
        :return: None
        """
        try:
            if self.list_students.count() == 0:
                self.show_empty_list_error()
            else:
                current_index = self.list_students.currentRow()
                item = self.list_students.item(current_index)
                if item is not None and current_index != 0:
                    existing_name, existing_score_str = item.text().split(':')
                    existing_name = existing_name.strip()
                    existing_score = existing_score_str.strip()

                    name, ok = QInputDialog.getText(self, 'Edit Student', 'Edit Student Name',
                                                    text=existing_name)

                    if ok:
                        # If the user wants to change the name, check for duplicate names
                        if name.capitalize() != existing_name:
                            existing_names = [self.list_students.item(row).text().split(':')[0].strip() for row in
                                              range(self.list_students.count())]
                            if name.capitalize() in existing_names:
                                error_message = QMessageBox()
                                error_message.setIcon(QMessageBox.Icon.Warning)
                                error_message.setWindowTitle("Duplicate Name")
                                error_message.setText('Student with this name already exists')
                                error_message.exec()
                                return  # Stop further execution if duplicate name is found

                        score_input, ok = QInputDialog.getInt(self, 'Edit Student',
                                                              'Edit Student Score',
                                                              value=int(existing_score))
                        if ok:
                            try:
                                item_text = f"{name.capitalize()}: \t\t{score_input}"
                                item.setText(item_text)
                            except ValueError:
                                return
        except Exception as e:
            print(f"An error occurred: {e}")

    def clear_students(self) -> None:
        """
        Clear the list of students and the summary.
        :return: None
        """
        try:
            if self.list_students.count() == 0:
                self.show_empty_list_error()
            else:
                question = QMessageBox.question(
                    self, 'Clear students',
                    f'Do you want to clear everything?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if question == QMessageBox.StandardButton.Yes:
                    self.list_students.clear()
                    self.list_summary.clear()
        except Exception as e:
            print(f'An error occurred: {e}')

    @staticmethod
    def grades_summary(grade: list[int]) -> list[str]:
        """
        Generate a list of grades (A, B, C, D, E) based on the given scores.
        :param grade: List of integer scores.
        :return: List of grade strings corresponding to the input scores.
        """
        grades = []
        best = max(grade)
        for student, score in enumerate(grade):
            if score >= best - 10:
                grades.append('A')
            elif score >= best - 20:
                grades.append('B')
            elif score >= best - 30:
                grades.append('C')
            elif score >= best - 40:
                grades.append('D')
            else:
                grades.append('E')

        return grades

    def sum_students(self) -> str:
        """
        Generates a summary of grades for every student's scores.
        :return: str: A summary containing information about the total number of students,
                 average grade, highest grade, and lowest grade, along with individual student data.
        """
        try:
            if self.list_students.count() == 0:
                self.show_empty_list_error()
            else:
                students_data = []
                for row in range(1, self.list_students.count()):
                    item = self.list_students.item(row)
                    if item is not None:
                        try:
                            student_data = item.text().split(":")
                            if len(student_data) >= 2:
                                name = student_data[0].strip()
                                score = int(student_data[1].strip())
                                student_dict = {"name": name, "grades": [score]}
                                students_data.append(student_dict)
                        except Exception as e:
                            print(f"Error processing item at row {row}: {e}")

                total_students = len(students_data)
                total_grades = sum(student['grades'][0] for student in students_data)
                average_grade = total_grades / total_students
                highest_grade = max(student['grades'][0] for student in students_data)
                lowest_grade = min(student['grades'][0] for student in students_data)

                summary_data = [
                    f"Total Students: {total_students}",
                    f"Average Grade: {average_grade:.2f}",
                    f"Highest Grade: {highest_grade}",
                    f"Lowest Grade: {lowest_grade}"
                ]

                # Calculate grades for each student
                grades = self.grades_summary([student['grades'][0] for student in students_data])

                # Add individual student data to the summary
                for student, grade in zip(students_data, grades):
                    name = student['name']
                    score = student['grades'][0]
                    summary_data.append(f"{name}'s score is {score} and grade is {grade}")

                # Clear and populate the list_summary widget
                self.list_summary.clear()
                self.list_summary.addItems(summary_data)

                try:
                    file_path = os.path.expanduser('~/Desktop/students.csv')
                    with open(file_path, 'w', newline='') as csv_file:
                        fieldnames = ['Students', 'Scores', 'Grades']
                        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                        writer.writeheader()

                        for student, grade in zip(students_data, grades):
                            name = student["name"]
                            score = student["grades"][0]
                            writer.writerow({'Students': name, 'Scores': score, 'Grades': grade})
                except Exception as e:
                    print(f'An unexpected CSV error occurred: {e}')
        except Exception as e:
            print(f"An error occurred: {e}")

    @staticmethod
    def open_csv() -> None:
        """
        Opens the CSV file containing student data directly. It's recommended to generate a summary of grades
        before using this to ensure the data is up-to-date.
        :return: None
        """
        file_path = os.path.expanduser('~/Desktop/students.csv')
        try:
            subprocess.Popen(['start', 'excel', file_path], shell=True)
        except Exception as e:
            print(f"Error opening CSV file: {e}")
