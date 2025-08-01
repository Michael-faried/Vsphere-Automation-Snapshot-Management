from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QComboBox, QMessageBox)
from PyQt6.QtGui import QFont
import ssl
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from pyvim import connect
from pyVmomi import vim
import pytz
import urllib.parse
from tkinter import messagebox
from tkinter import filedialog
import csv,time
from datetime import datetime, timedelta
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QThread, pyqtSignal
import ctypes  
import concurrent.futures


class VCenterSnapshotViewer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.vcenter_server = "" #### ADD your Vsphere IP here 
        self.vcenter_user = ""
        self.vcenter_password = ""
        self.matching_snapshots = []
        self.service_instance = None

        self.init_ui()


    def init_ui(self):

        # Apply system-independent styles
        self.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF; /* White background */
                color: #000000; /* Black text color */
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: #F1F1F1; /* Light background */
                color: #000000; /* Black text */
            }
            QPushButton {
                background-color: #1877FF; /* Blue background for buttons */
                color: white;
            }
            QPushButton:hover {
                background-color: #1565C0;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
        """)

        # Create a modern font for the snapshot text box
        app_font = QFont("Segoe UI", 12, QFont.Weight.DemiBold)  # Clean, modern look
        app_font.setStyleHint(QFont.StyleHint.SansSerif)  # Ensures the use of a sans-serif font
        app_font.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1)  # Slight letter spacing for clarity

                        
        # Set the application ID for the Windows taskbar (necessary for taskbar icon)
        if hasattr(ctypes, 'windll'):  # Only apply on Windows
            myappid = "com.company.vcenter.snapshotmanagement"  # Unique application ID
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        # Set the application icon for the taskbar
        self.setWindowTitle("vCenter Snapshot Management")
        self.setWindowIcon(QIcon(r"src\icon.ico"))


        screen_geometry = QApplication.primaryScreen().availableGeometry()
        window_width, window_height = 800, 600
        x = (screen_geometry.width() - window_width) // 2
        y = (screen_geometry.height() - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)

        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # Grid layout for input fields
        input_layout = QGridLayout()
        main_layout.addLayout(input_layout)
        
        # vCenter Server
        input_layout.addWidget(QLabel("vCenter Server:"), 0, 0)
        self.server_input = QLineEdit(self.vcenter_server)
        self.server_input.setReadOnly(True)  # input non-editable
        self.server_input.setStyleSheet(self.input_style2())
        input_layout.addWidget(self.server_input, 0, 1)



        # vCenter Username
        input_layout.addWidget(QLabel("vCenter Username:"), 1, 0)
        self.username_input = QLineEdit()
        self.username_input.setStyleSheet(self.input_style())
        self.username_input.setPlaceholderText("Enter vCenter Username")  
        input_layout.addWidget(self.username_input, 1, 1)

        # vCenter Password
        input_layout.addWidget(QLabel("vCenter Password:"), 2, 0)
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setStyleSheet(self.input_style())
        self.password_input.setPlaceholderText("Enter vCenter Password")  
        input_layout.addWidget(self.password_input, 2, 1)

        # Snapshot Age
        input_layout.addWidget(QLabel("Snapshot Age (days):"), 3, 0)
        self.age_input = QLineEdit()
        self.age_input.setStyleSheet(self.input_style())
        self.age_input.setPlaceholderText("Enter Snapshot Age in Days")  
        input_layout.addWidget(self.age_input, 3, 1)

        # Snapshot Name Dropdown
        input_layout.addWidget(QLabel("Snapshot Type:"), 4, 0)
        self.snapshot_name_dropdown = QComboBox()
        self.snapshot_name_dropdown.addItems([
            "-- Please Select Snapshots Type --",
            "VEEAM SNAPSHOTS",
            "VCD ADMIN","Cloud NOC USER",
            "All SNAPSHOTS"
        ])
        self.snapshot_name_dropdown.setStyleSheet(self.combo_style())
        input_layout.addWidget(self.snapshot_name_dropdown, 4, 1)

        # Buttons
        button_layout = QHBoxLayout()
        main_layout.addLayout(button_layout)

        self.retrieve_button = QPushButton("Retrieve Snapshots")
        self.retrieve_button.setStyleSheet(self.button_style())
        self.retrieve_button.clicked.connect(self.retrieve_snapshots)
        button_layout.addWidget(self.retrieve_button)

        self.download_button = QPushButton("Download CSV")
        self.download_button.setObjectName("downloadButton")  
        self.download_button.setEnabled(False)
        self.download_button.setStyleSheet(self.button_style())
        self.download_button.clicked.connect(self.download_csv)
        button_layout.addWidget(self.download_button)

        # New Delete Snapshots Button (Red)
        self.delete_button = QPushButton("Delete Snapshots")
        self.delete_button.setStyleSheet(self.delete_button_style())  
        self.delete_button.setEnabled(False) 
        self.delete_button.clicked.connect(self.delete_snapshots)
        button_layout.addWidget(self.delete_button)

        # Text Box for Snapshot Details
        self.snapshot_text = QTextEdit()
        self.snapshot_text.setReadOnly(True)
        self.snapshot_text.setStyleSheet(self.text_box_style())
        main_layout.addWidget(self.snapshot_text)
        
        # Apply the font to the snapshot text box only
        self.snapshot_text.setFont(app_font)

        # Display the "Ready" text inside the QTextEdit box initially
        self.snapshot_text.append("Ready")

        # Apply label styles
        self.apply_label_style()


    # this function to apply label style
    def apply_label_style(self):
        labels = self.findChildren(QLabel)  
        for label in labels:
            label.setStyleSheet(self.label_style())



    def label_style(self):
        return """
        QLabel {
            font-size: 16px;  /* Increase font size */
            font-weight: bold;  /* Make the text bold */
            color: #063970;  /* Set text color to #063970 */
        }
        """



    def input_style(self):
        return """
        QLineEdit {
            background-color: #F1F1F1;
            border: 2px solid #1877FF;
            border-radius: 10px;
            padding: 3px;
            font-size: 16px;
        }
        QLineEdit:focus {
            background-color:rgb(166, 222, 226);
        }
        """


        
    def input_style2(self):
        return """
        QLineEdit {
            background-color:rgb(180, 180, 180);
            border: 2px solid #1877FF;
            border-radius: 10px;
            padding: 3px;
            font-size: 16px;
        }
        QLineEdit:focus {
            background-color:rgb(180, 180, 180);
        }
        """

    def button_style(self):
        return """
        QPushButton {
            background-color: #1877FF;
            color: white;
            border-radius: 15px;
            font-size: 14px;
            font-weight: bold;
            padding: 10px 20px;
            border: none;
        }
        QPushButton:hover {
            background-color: #1565C0;
        }
        QPushButton:pressed {
            background-color: #0D47A1;
        }
        QPushButton#downloadButton {
            background-color: #66BB6A;  /* Light Green */
        }
        QPushButton#downloadButton:hover {
            background-color: #81C784;  /* Darker Green */
        }
        QPushButton#downloadButton:pressed {
            background-color: #388E3C;  /* Even Darker Green */
        }
        """

    def delete_button_style(self):
        return """
        QPushButton {
            background-color: #D32F2F;  /* Red color */
            color: white;
            border-radius: 15px;
            font-size: 14px;
            font-weight: bold;
            padding: 10px 20px;
            border: none;
        }
        QPushButton:hover {
            background-color: #C62828;
        }
        QPushButton:pressed {
            background-color: #B71C1C;
        }
        """
        
    def combo_style(self):
        return """
        QComboBox {
            background-color: #F1F1F1;
            border: 2px solid #1877FF;
            border-radius: 10px;
            font-size: 14px;
            padding: 5px;
        }
        QComboBox::down-arrow {
            border: none;
            width: 12px;
            height: 12px;
            background-color: #1877FF;  /* Blue color for the arrow */
            border-radius: 3px;  /* Rounded corners for a modern look */
        }
        QComboBox::down-arrow:hover {
            background-color: #1565C0;  /* Darker blue on hover */
        }
        QComboBox:editable {
            background-color: #E0E0E0;
        }
        """


    def text_box_style(self):
        return """
        QTextEdit {
            background-color: #F1F1F1;
            border: 2px solid #1877FF;
            border-radius: 10px;
            font-size: 14px;
            padding: 10px;
        }

        QTextEdit QScrollBar:vertical {
            border: none;
            background:rgb(127, 216, 252);
            width: 10px;
            border-radius: 5px;
            margin: 0px 0px 0px 0px;
        }

        QTextEdit QScrollBar::handle:vertical {
            background: #1877FF;
            border-radius: 5px;
            min-height: 30px;
        }

        QTextEdit QScrollBar::handle:vertical:hover {
            background: #1565C0;
        }

        QTextEdit QScrollBar::add-line:vertical, QTextEdit QScrollBar::sub-line:vertical {
            border: none;
            background: none;
            height: 0px;
        }

        QTextEdit QScrollBar::up-arrow:vertical, QTextEdit QScrollBar::down-arrow:vertical {
            border: none;
            background: none;
        }

        QTextEdit QScrollBar:horizontal {
            border: none;
            background:rgb(241, 241, 241);
            height: 10px;
            border-radius: 5px;
            margin: 0px 0px 0px 0px;
        }

        QTextEdit QScrollBar::handle:horizontal {
            background: #1877FF;
            border-radius: 5px;
            min-width: 30px;
        }

        QTextEdit QScrollBar::handle:horizontal:hover {
            background: #1565C0;
        }

        QTextEdit QScrollBar::add-line:horizontal, QTextEdit QScrollBar::sub-line:horizontal {
            border: none;
            background: none;
            width: 0px;
        }

        QTextEdit QScrollBar::left-arrow:horizontal, QTextEdit QScrollBar::right-arrow:horizontal {
            border: none;
            background: none;
        }
        """


    def connect_to_vcenter(self):
        try:
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE

            self.service_instance = connect.SmartConnect(
                host=self.server_input.text(),
                user=self.username_input.text(),
                pwd=self.password_input.text(),
                sslContext=context
            )
            return self.service_instance
        except Exception as e:
            QMessageBox.critical(self, "Connection Error", f"Failed to connect to vCenter: {str(e)}")
            return None

    def get_snapshot_age(self, snapshot_create_time):
        now = datetime.now(pytz.utc)
        return (now - snapshot_create_time).days

    def get_formatted_date(self, snapshot_create_time):
        return snapshot_create_time.strftime('%Y-%m-%d')
    

    def check_snapshots(self, vm, snapshot_name_filter, age_filter):
        vm_snapshots = []
        try:
            if vm.snapshot:
                for snapshot in vm.snapshot.rootSnapshotList:
                    decoded_name = urllib.parse.unquote(snapshot.name)
                    snapshot_age = self.get_snapshot_age(snapshot.createTime)

                    # Check if the snapshot meets both the name and age filters
                    if (
                        snapshot_name_filter == "All SNAPSHOTS" or
                        (snapshot_name_filter == "VEEAM SNAPSHOTS" and (
                            "Restore Point" in decoded_name or
                            "Veeam Replica Working Snapshot" in decoded_name or
                            "VEEAM BACKUP TEMPORARY SNAPSHOT" in decoded_name or
                            "VEEAM" in decoded_name
                        )) or
                        (snapshot_name_filter == "VCD ADMIN" and "user-VCD-snapshot" in decoded_name) or
                        (snapshot_name_filter == "Cloud NOC USER" and "VM Snapshot" in decoded_name)
                    ) and snapshot_age >= age_filter:
                        vm_snapshots.append({
                            "vm_name": vm.name,
                            "snapshot_name": decoded_name,
                            "create_time": snapshot.createTime,
                            "age_days": snapshot_age
                        })
        except Exception as e:
            print(f"Error while checking snapshots for VM {vm.name}: {str(e)}")

        return vm_snapshots



    def list_snapshots_for_all_vms(self, service_instance, snapshot_name_filter, age_filter):
        self.matching_snapshots.clear()

        try:
            content = service_instance.RetrieveContent()
            container = content.viewManager.CreateContainerView(content.rootFolder, [vim.VirtualMachine], True)
            vms = container.view

            # Display the status message inside the QTextEdit box
            self.snapshot_text.append(f"     Found <font color='#D2122E'>{len(vms)}</font> VMs in Vcenter. Checking Now For Snapshots... <font color='#D2122E'> Please Wait just seconds ..... </font><br>\n")

            QApplication.processEvents()

            with ThreadPoolExecutor(max_workers=30) as executor:
                future_to_vm = {executor.submit(self.check_snapshots, vm, snapshot_name_filter, age_filter): vm for vm in vms}

                for future in as_completed(future_to_vm):
                    vm_snapshots = future.result()
                    if vm_snapshots:
                        self.matching_snapshots.extend(vm_snapshots)

            container.Destroy()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error while listing snapshots: {str(e)}")

    def retrieve_snapshots(self):
        try:
            snapshot_age_filter = int(self.age_input.text())
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid age filter input. Please enter a valid number.")
            return

        snapshot_name_filter = self.snapshot_name_dropdown.currentText()

        service_instance = self.connect_to_vcenter()
        if not service_instance:
            return

        # Clear any previous status and start with "Ready"
        self.snapshot_text.clear()
        self.snapshot_text.append("<font color='green'> Connected To Vcenter </font>")

        self.snapshot_text.append(" <font color='blue'>  Retrieving snapshots </font> ... ")
        QApplication.processEvents()

        try:
            self.list_snapshots_for_all_vms(service_instance, snapshot_name_filter, snapshot_age_filter)

            if self.matching_snapshots:

                # Sort the snapshots by age in descending order
                self.matching_snapshots.sort(key=lambda x: x['age_days'], reverse=True)

                # Display the total number of snapshots found
                self.snapshot_text.append(f"Total Snapshots Found: <font color='#480ca8'>{len(self.matching_snapshots)}</font>")

                # Set font to monospaced for column alignment
                font = self.snapshot_text.font()
                font.setFamily("Courier New")
                self.snapshot_text.setFont(font)

                # Format and add header
                header = f" {'VM Name':<45}{'Snapshot Name':<45}{'Created At':<25}{'Age (Days)':<15}"

                self.snapshot_text.append(header)
                self.snapshot_text.append("=" * len(header))

                #  matching snapshots to the QTextEdit
                for snapshot in self.matching_snapshots:
                    line = f"{snapshot['vm_name']:<45}{snapshot['snapshot_name']:<45}{self.get_formatted_date(snapshot['create_time']):<25}{snapshot['age_days']:<15}"
                    self.snapshot_text.append(line)
                    QApplication.processEvents()

                # Enable the "Download CSV" & "Delete Button" button once data is ready
                self.download_button.setEnabled(True)
                self.delete_button.setEnabled(True)
            else:
                self.snapshot_text.append("No snapshots found with the specified criteria.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error while retrieving snapshots: {str(e)}")




    def download_csv(self):
        if not self.matching_snapshots:
            messagebox.showwarning("No Data", "No snapshots to download.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                with open(file_path, mode="w", newline="") as file:
                    writer = csv.DictWriter(file, fieldnames=["vm_name", "snapshot_name", "create_time", "age_days"])
                    writer.writeheader()
                    writer.writerows(self.matching_snapshots)
                messagebox.showinfo("Success", f"Snapshots data saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV: {str(e)}")

    

    def delete_snapshots(self):
        if not self.matching_snapshots:
            QMessageBox.warning(self, "No Data", "No snapshots to delete.")
            return

        reply = QMessageBox.question(
            self, "Delete Snapshots", "Are you sure you want to delete all reviewed snapshots?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.No:
            return

        batch_size = 5
        total_snapshots = len(self.matching_snapshots)
        total_batches = (total_snapshots + batch_size - 1) // batch_size
        estimated_time = round((total_batches * 70) / 60, 2)  # 20 seconds per batch

        self.snapshot_text.append(f"\n \n Starting snapshot deletion... Estimated time: ~{estimated_time} MINs ")

        # Disable the delete button to prevent duplicate actions .... Very Important 
        self.delete_button.setEnabled(False)

        # Create and start the worker thread
        self.worker = SnapshotDeletionWorker(self.service_instance, self.matching_snapshots, batch_size)
        self.worker.update_status.connect(self.snapshot_text.append)  # Update status in the GUI
        self.worker.completed.connect(self.on_deletion_completed)
        self.worker.start()



    def on_deletion_completed(self):
        self.snapshot_text.append("Snapshot deletion completed.")
        self.delete_button.setEnabled(True)  # Re-enable the delete button


class SnapshotDeletionWorker(QThread):
    # Signal to update the GUI with progress or completion
    update_status = pyqtSignal(str)
    completed = pyqtSignal()

class SnapshotDeletionWorker(QThread):
    # Signal to update the GUI with progress or completion
    update_status = pyqtSignal(str)
    completed = pyqtSignal()

    def __init__(self, service_instance, snapshots, batch_size):
        super().__init__()
        self.service_instance = service_instance
        self.snapshots = snapshots
        self.batch_size = batch_size


    def run(self):
        total_snapshots = len(self.snapshots)
        ongoing_futures = []  # Store ongoing futures to let them complete in the background
        executor = ThreadPoolExecutor(max_workers=self.batch_size)  # Reuse a single executor

        for i in range(0, total_snapshots, self.batch_size):
            batch = self.snapshots[i:i + self.batch_size]
            self.update_status.emit(f"<b> Processing batch {i // self.batch_size + 1} of {total_snapshots // self.batch_size + 1}...</b>")

            batch_start_time = datetime.now()

            # Submit batch tasks without waiting for completion
            for snapshot in batch:
                future = executor.submit(self.delete_vm_snapshots, snapshot['vm_name'])
                ongoing_futures.append(future)

            # Wait up to 1 minute for some tasks, but don't block the loop
            while datetime.now() - batch_start_time < timedelta(seconds=30):
                time.sleep(1)

            self.update_status.emit(f"<b> <font color='#D2122E'>Batch {i // self.batch_size + 1} started. Moving to the next batch...</font></b>")

        self.update_status.emit("<b>    <b> All batches started. Waiting for background tasks to complete...</b>  </b>")

        # Wait for any remaining tasks to complete after all batches have been submitted
# Process completed tasks as they finish
        for future in concurrent.futures.as_completed(ongoing_futures):
            try:
                vm_name = self.snapshots[ongoing_futures.index(future)]['vm_name']
                success = future.result()
                if success:
                    self.update_status.emit(f"<font color='#28b463'> Snapshots for VM {vm_name} deleted successfully.</font>")
                else:
                    self.update_status.emit(f"<font color='#E52020'> Failed to delete snapshots for VM {vm_name}.</font>")
            except Exception as e:
                self.update_status.emit(f"Error deleting snapshots: {str(e)}")


        self.update_status.emit("<b>All snapshots processed.</b>")
        self.completed.emit()


    def delete_vm_snapshots(self, vm_name):
        vm = self.get_vm_by_name(vm_name)
        if vm:
            try:
                if vm.snapshot:
                    task = vm.RemoveAllSnapshots_Task()
                    while task.info.state not in [vim.TaskInfo.State.success, vim.TaskInfo.State.error]:
                        time.sleep(1)  # Polling task status every second
                    return task.info.state == vim.TaskInfo.State.success
            except Exception as e:
                print(f"Error deleting snapshots for VM {vm_name}: {str(e)}")
        return False

    def get_vm_by_name(self, vm_name):
        try:
            content = self.service_instance.RetrieveContent()
            container = content.viewManager.CreateContainerView(content.rootFolder, [vim.VirtualMachine], True)
            for vm in container.view:
                if vm.name == vm_name:
                    container.Destroy()
                    return vm
            container.Destroy()
            return None
        except Exception as e:
            print(f"Error retrieving VM {vm_name}: {str(e)}")
            return None





if __name__ == "__main__":
    app = QApplication([])
    window = VCenterSnapshotViewer()
    window.show()
    app.exec()