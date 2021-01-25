import io
import sys

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.label import Label
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import Tk
from json_to_excel import generate_excel


class ExcelExportApp(App):

    def build(self):
        self.title = " Mission Victory India: Editorial Export Utility"
        self.icon = "icon.png"
        self.input_file = None
        self.output_dir = None

        main_layout = BoxLayout(orientation="vertical")

        input_layout = BoxLayout(orientation="horizontal", size_hint=(1, 0.08))
        input_label = Label(text="Select Input JSON File:")
        input_layout.add_widget(input_label)

        self.input_text = TextInput(multiline=True, readonly=True, halign="left")
        input_layout.add_widget(self.input_text)

        input_button = Button(text="Upload JSON",
                              background_normal="C:\\Users\\user\\PycharmProjects\\MVIPostToExcel\\icons8-json-64.png"
                              )
        input_button.bind(on_press=self.select_input_file)
        input_layout.add_widget(input_button)
        main_layout.add_widget(input_layout)

        output_layout = BoxLayout(orientation="horizontal", size_hint=(1, 0.08))
        output_label = Label(text="Select Output Excel Directory:")
        output_layout.add_widget(output_label)

        self.output_text = TextInput(multiline=True, readonly=True, halign="left")
        output_layout.add_widget(self.output_text)

        output_button = Button(text="Select Output Directory")
        output_button.bind(on_press=self.select_output_dir)
        output_layout.add_widget(output_button)
        main_layout.add_widget(output_layout)

        export_button = Button(text="Export", pos_hint={"center_x": 0.5, "center_y": 0.5}, size_hint=(1, 0.08))
        export_button.bind(on_press=self.export)
        main_layout.add_widget(export_button)

        self.export_log = TextInput(multiline=True, readonly=True, halign="left")
        main_layout.add_widget(self.export_log)

        return main_layout

    def select_input_file(self, instance):
        Tk().withdraw()
        self.input_file = askopenfilename(title="Select JSON input file",
                                          filetypes=(("JSON files", "*.json"), ("all files", "*.*")))
        self.input_text.text = self.input_file

    def select_output_dir(self, instance):
        Tk().withdraw()
        self.output_dir = askdirectory()
        self.output_text.text = self.output_dir

    def export(self, instance):
        if self.export_log.text == "" or self.export_log.text is None:
            self.export_log.text = "Starting Export"
        else:
            self.export_log.text += "\nStarting Export"
        if self.input_file == "" or self.input_file is None:
            self.export_log.text += "\nNo Input File Selected; Aborting"
        else:
            if self.output_dir == "" or self.output_dir is None:
                self.export_log.text += "\nNo Output Directory Selected; Aborting"
            else:
                self.export_log.text += "\nInput File: " + self.input_file
                self.export_log.text += "\nOutput Directory: " + self.output_dir
                old_sys_out = sys.stdout
                new_sys_out = io.StringIO()
                sys.stdout = new_sys_out
                generate_excel(self.input_file, self.output_dir)
                self.export_log.text += "\n\n\n" + new_sys_out.getvalue() + "\n\n\n"
                sys.stdout = old_sys_out
                self.export_log.text += "\nExport Completed"


if __name__ == '__main__':
    export_app = ExcelExportApp()
    export_app.run()
