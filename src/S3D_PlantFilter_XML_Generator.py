import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QFileDialog, QLabel, QMessageBox, QProgressBar, QTabWidget,
)
from PyQt6.QtCore import Qt
from datetime import datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom

class XmlGeneratorApp(QWidget):
    """
    A PyQt6 application to generate an XML filter file from an Excel input.
    """
    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.last_directory = "" # To remember the last opened directory
        self.excel_path_template = None
        self.init_ui()

    def init_ui(self):
        """Initializes the user interface components."""
        self.setWindowTitle('S3D Filter XML Generator')
        self.setGeometry(300, 300, 550, 300)

        main_layout = QVBoxLayout(self)
        tab_widget = QTabWidget()

        # --- Tab 1: 1. Filter (Detail) ---
        tab1 = QWidget()
        tab1_layout = QVBoxLayout(tab1)
        
        info_label_simple = QLabel('Select an Excel file with sheet "S3DFilter" and required columns: Name(Filter), FullPath(Filter), ObjectPath')        
        info_label_simple.setWordWrap(True)
        info_label_simple.setToolTip('Select an Excel file with sheet "S3DFilter" and required columns: Name(Filter), FullPath(Filter), ObjectPath')

        file_layout_simple = QHBoxLayout()
        file_layout_simple.addWidget(QLabel("Excel File:"))
        self.file_path_display = QLineEdit('No file selected.')
        self.file_path_display.setReadOnly(True)
        self.file_path_display.setStyleSheet('font-style: italic; color: grey;')
        file_layout_simple.addWidget(self.file_path_display, 1)
        self.load_button = QPushButton('Load')
        self.load_button.clicked.connect(self.open_file_dialog_simple)
        file_layout_simple.addWidget(self.load_button)
        self.run_button = QPushButton('Run')
        self.run_button.clicked.connect(self.generate_xml_simple)
        self.run_button.setEnabled(False)
        file_layout_simple.addWidget(self.run_button)

        tab1_layout.addLayout(file_layout_simple)
        tab1_layout.addStretch()

        # --- Tab 2: 2. Filter (Bulk) ---
        tab2 = QWidget()
        tab2_layout = QVBoxLayout(tab2)
        
        info_label_template = QLabel("Select an Excel file with sheets: '1.S3DFilterPath', '2.S3DFilterName', '3.FixedObjectPath'")        
        info_label_template.setWordWrap(True)
        info_label_template.setToolTip("Select an Excel file with sheets: '1.S3DFilterPath', '2.S3DFilterName', '3.FixedObjectPath'")

        file_layout_template = QHBoxLayout()
        file_layout_template.addWidget(QLabel("Excel File:"))
        self.file_path_display_template = QLineEdit('No file selected.')
        self.file_path_display_template.setReadOnly(True)
        self.file_path_display_template.setStyleSheet('font-style: italic; color: grey;')
        file_layout_template.addWidget(self.file_path_display_template, 1)
        self.load_button_template = QPushButton('Load')
        self.load_button_template.clicked.connect(self.open_file_dialog_template)
        file_layout_template.addWidget(self.load_button_template)
        self.run_button_template = QPushButton('Run')
        self.run_button_template.clicked.connect(self.generate_xml_from_template)
        self.run_button_template.setEnabled(False)
        file_layout_template.addWidget(self.run_button_template)

        tab2_layout.addLayout(file_layout_template)
        tab2_layout.addStretch()

        tab_widget.addTab(tab1, "1. Filter (Detail)")
        tab_widget.addTab(tab2, "2. Filter (Bulk)")

        main_layout.addWidget(tab_widget)

        # --- Global Status and Progress Bar ---
        self.status_label = QLabel("Ready.")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.status_label)
        main_layout.addWidget(self.progress_bar)

    def open_file_dialog_simple(self):
        """Opens a file dialog for the 1. Filter (Detail)."""
        self._open_file_dialog(
            'excel_path', 
            self.file_path_display, 
            self.run_button
        )

    def open_file_dialog_template(self):
        """Opens a file dialog for the 2. Filter (Bulk)."""
        self._open_file_dialog(
            'excel_path_template', 
            self.file_path_display_template, 
            self.run_button_template
        )

    def _open_file_dialog(self, path_attr, path_display_widget, run_button):
        """Opens a file dialog to select an Excel file."""
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", self.last_directory, "Excel Files (*.xlsx *.xls)")
        if file_name:
            setattr(self, path_attr, file_name)
            path_display_widget.setText(file_name)
            path_display_widget.setStyleSheet('font-style: normal; color: black;')
            run_button.setEnabled(True)
            self.status_label.setText('File loaded. Click "Run" to generate the XML.')
            # Remember the directory for next time
            self.last_directory = os.path.dirname(file_name)

    def generate_xml_simple(self):
        """
        Reads the Excel file, processes the data, and writes the XML file.
        """
        if not self.excel_path:
            QMessageBox.warning(self, 'Warning', 'Please select an Excel file first.')
            return
        
        self.status_label.setText("Generating XML for 'Filter (Detail)'...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        QApplication.processEvents()

        try:
            # --- Read data from all three sheets ---
            xls_simple = pd.ExcelFile(self.excel_path)
            if 'S3DFilter' not in xls_simple.sheet_names:
                QMessageBox.critical(self, 'Error', "Excel file must contain the sheets: 'S3DFilter'")
                return

            df_simple = pd.read_excel(xls_simple, sheet_name='S3DFilter')

            # --- Validate required columns ---
            required_columns = ['Name(Filter)', 'FullPath(Filter)', 'ObjectPath']
            if not all(col in df_simple.columns for col in required_columns):
            
                QMessageBox.critical(self, 'Error', f"Excel file must contain the columns: {', '.join(required_columns)}")
                return
            
            # --- Build XML Structure ---
            xml_root = ET.Element('xml')
            
            # Information Element
            info_elem = ET.SubElement(xml_root, 'Information')
            info_elem.set('FileType', 'FiltersAndStyleRules')
            info_elem.set('DateExported', datetime.now().strftime('%m_%d_%y_%H_%M_%S'))

            filters_elem = ET.SubElement(xml_root, 'Filters')
            plant_filters_elem = ET.SubElement(filters_elem, 'PlantFilters')

            # --- New logic: Combine FullPath(Filter) and Name(Filter) ---
            # Create a new column for the final full path.
            df_simple['FinalFullPath'] = df_simple.apply(lambda row: f"{row['FullPath(Filter)']}\\{row['Name(Filter)']}", axis=1)

            # Group data by Filter Name and FullPath
            grouped = df_simple.groupby(['Name(Filter)', 'FinalFullPath'])
            num_groups = len(grouped)
            self.progress_bar.setMaximum(num_groups)
            
            for i, ((name, full_path), group) in enumerate(grouped):
                # Create Filter element
                filter_elem = ET.SubElement(plant_filters_elem, 'Filter', {
                    'Name': str(name),
                    'FullPath': str(full_path),
                    'Category': '1',
                    'FilterType': 'Simple',
                    'Ignore': 'False'
                })

                # Add the static FilterDef for MFObjectType
                ET.SubElement(filter_elem, 'FilterDef', {
                    'Type': 'MFObjectType',
                    'GroupPath': r'Systems\PipelineSystems'
                })

                # Add FilterDef for each ObjectPath in the group
                for _, row in group.iterrows():
                    ET.SubElement(filter_elem, 'FilterDef', {
                        'Type': 'MFSystem',
                        'ObjectPath': str(row['ObjectPath']),
                        'IncludeNested': 'True'
                    })
                self.progress_bar.setValue(i + 1)
                QApplication.processEvents()
            
            # Add empty user and catalog filters
            ET.SubElement(filters_elem, 'UserFilters')
            ET.SubElement(filters_elem, 'CatalogFilters')

            # --- Write to file with pretty printing ---
            output_path, _ = QFileDialog.getSaveFileName(self, "Save XML File", "output_filters.xml", "XML Files (*.xml)")

            if output_path:
                # Use minidom for pretty printing the XML
                xml_string = ET.tostring(xml_root, 'utf-8')
                reparsed = minidom.parseString(xml_string)
                pretty_xml = reparsed.toprettyxml(indent="\t", encoding="utf-8")

                with open(output_path, "wb") as f:
                    f.write(pretty_xml)
                
                QMessageBox.information(self, 'Success', f'XML file successfully generated at:\n{output_path}')
                self.status_label.setText('Done. You can load another file or close the application.')

        except Exception as e:
            QMessageBox.critical(self, 'An Error Occurred', str(e))
            self.status_label.setText("An error occurred.")
        finally:
            self.progress_bar.setVisible(False)

    def generate_xml_from_template(self):
        """
        Reads an Excel file with a template structure and generates the XML.
        """
        if not self.excel_path_template:
            QMessageBox.warning(self, 'Warning', 'Please select an Excel file first.')
            return

        self.status_label.setText("Generating XML for 'Filter (Bulk)'...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        QApplication.processEvents()

        try:
            # --- Read data from all three sheets ---
            xls = pd.ExcelFile(self.excel_path_template)
            if '1.S3DFilterPath' not in xls.sheet_names or '2.S3DFilterName' not in xls.sheet_names or '3.FixedObjectPath' not in xls.sheet_names:
                QMessageBox.critical(self, 'Error', "Excel file must contain the sheets: '1.S3DFilterPath', '2.S3DFilterName', '3.FixedObjectPath'")
                return

            df_filter = pd.read_excel(xls, sheet_name='2.S3DFilterName')
            df_object_path = pd.read_excel(xls, sheet_name='3.FixedObjectPath')
            df_full_path = pd.read_excel(xls, sheet_name='1.S3DFilterPath')

            # --- Validate required columns ---
            required_filter_cols = ['Name(Filter)', 'SystemName(WBS)']
            if not all(col in df_filter.columns for col in required_filter_cols):
                QMessageBox.critical(self, 'Error', f"The '2.S3DFilterName' sheet must contain the columns: {', '.join(required_filter_cols)}")
                return
            if 'ObjectPath(Template)' not in df_object_path.columns:
                QMessageBox.critical(self, 'Error', "The 'FixedObjectPath' sheet must contain the column: 'ObjectPath(Template)'")
                return
            if 'FilterPath(Template)' not in df_full_path.columns:
                QMessageBox.critical(self, 'Error', "The '1.S3DFilterPath' sheet must contain the column: 'FilterPath(Template)'")
                return
            
            full_path_template = df_full_path['FilterPath(Template)'].iloc[0]
            object_path_templates = df_object_path['ObjectPath(Template)'].tolist()

            # --- Build XML Structure ---
            xml_root = ET.Element('xml')
            info_elem = ET.SubElement(xml_root, 'Information', {'FileType': 'FiltersAndStyleRules', 'DateExported': datetime.now().strftime('%m_%d_%y_%H_%M_%S')})
            filters_elem = ET.SubElement(xml_root, 'Filters')
            plant_filters_elem = ET.SubElement(filters_elem, 'PlantFilters')

            self.progress_bar.setMaximum(len(df_filter))

            # Get column indices to handle special characters in names
            name_col_idx = df_filter.columns.get_loc('Name(Filter)') + 1  # +1 because index=False for itertuples
            wbs_col_idx = df_filter.columns.get_loc('SystemName(WBS)') + 1

            for i, row in enumerate(df_filter.itertuples(index=True)):
                name = row[name_col_idx] 
                wbs_name = row[wbs_col_idx]

                # Use WBS name for path if available, otherwise fallback to filter name
                path_suffix = wbs_name if pd.notna(wbs_name) and str(wbs_name).strip() else name
                full_path = f"{full_path_template}\\{path_suffix}"
                
                filter_elem = ET.SubElement(plant_filters_elem, 'Filter', {
                    'Name': str(name),
                    'FullPath': str(full_path),
                    'Category': '1',
                    'FilterType': 'Simple',
                    'Ignore': 'False'
                })

                ET.SubElement(filter_elem, 'FilterDef', {'Type': 'MFObjectType', 'GroupPath': r'Systems\PipelineSystems'})

                for obj_path_template in object_path_templates:
                    object_path = f"{obj_path_template}\\{path_suffix}"
                    ET.SubElement(filter_elem, 'FilterDef', {'Type': 'MFSystem', 'ObjectPath': str(object_path), 'IncludeNested': 'True'})
                
                self.progress_bar.setValue(i + 1)
                QApplication.processEvents()

            ET.SubElement(filters_elem, 'UserFilters')
            ET.SubElement(filters_elem, 'CatalogFilters')

            # --- Write to file ---
            output_path, _ = QFileDialog.getSaveFileName(self, "Save XML File", "output_filters_template.xml", "XML Files (*.xml)")
            if output_path:
                xml_string = ET.tostring(xml_root, 'utf-8')
                reparsed = minidom.parseString(xml_string)
                pretty_xml = reparsed.toprettyxml(indent="\t", encoding="utf-8")

                with open(output_path, "wb") as f:
                    f.write(pretty_xml)
                
                QMessageBox.information(self, 'Success', f'XML file successfully generated at:\n{output_path}')
                self.status_label.setText('Done. You can load another file or close the application.')

        except Exception as e:
            QMessageBox.critical(self, 'An Error Occurred', str(e))
            self.status_label.setText("An error occurred.")
        finally:
            self.progress_bar.setVisible(False)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = XmlGeneratorApp()
    ex.show()
    sys.exit(app.exec())
