import customtkinter as ctk
import pandas as pd
import sys
import os

ctk_listbox = os.path.abspath(os.path.join(os.path.dirname(__file__), 'CTkListbox-main'))
sys.path.append(ctk_listbox)

from CTkListbox import ctk_listbox as ctkLB

class MergeXcel(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MergeXcel")
        self.geometry("500x600")
        self.after(201, self.iconbitmap("E:\\MergeXcel\\icon.ico"))

        #-------------------- HEADER --------------------#
        header_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        header_frame.place(relx=0.5, rely=0.2, relwidth=0.95, relheight=0.3, anchor=ctk.S)

        header_label = ctk.CTkLabel(header_frame, text="MergeXcel", font=("Lucida Calligraphy", 50))
        header_label.place(relx=0.5, rely=0.65, anchor=ctk.CENTER)
        
        #-------------------- BODY --------------------#
        body_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        body_frame.place(relx=0.5, rely=0.225, relwidth=0.95, relheight=0.76, anchor=ctk.N)

        self.filenames = []
        select_files_button = ctk.CTkButton(body_frame, text="SELECT👆", command=self.openFileNames)
        select_files_button.place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

        self.selected_files_listbox = ctkLB.CTkListbox(body_frame)
        self.selected_files_listbox.place(relx=0.5, rely=0.45, relwidth=0.6, relheight=0.55, anchor=ctk.CENTER)

        self.radio_var = ctk.IntVar(value=2)
        multi_sheets_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into multiple sheets", variable=self.radio_var, value=1, state=ctk.DISABLED)# disabled for now to avoid complications
        multi_sheets_radio_button.place(relx=0.3, rely=0.8, anchor=ctk.CENTER)
        single_sheet_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into single sheet", variable=self.radio_var, value=2)
        single_sheet_radio_button.place(relx=0.7, rely=0.8, anchor=ctk.CENTER)

        merge_button = ctk.CTkButton(body_frame, text="MERGE 📄➕📄", command=self.mergeFiles)
        merge_button.place(relx=0.5, rely=0.9, anchor=ctk.CENTER)

        self.mainloop()

    def openFileNames(self):
        new_filenames = list(ctk.filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")], multiple=True))

        if len(new_filenames) != 0:
            self.filenames = new_filenames
            for i in range(self.selected_files_listbox.size()):
                self.selected_files_listbox.delete(i)
                
            for i, file in enumerate(self.filenames):
                self.selected_files_listbox.insert(i, file)

    def mergeFiles(self):
        if self.radio_var.get() == 2 and len(self.filenames) != 0:
            merged_df = pd.read_excel(self.filenames[0])

            for file in self.filenames[1:]:
                df = pd.read_excel(file)
                
                merged_df = pd.merge(merged_df, df, how="outer")

            merged_df.to_excel('C:\\Users\\Lenovo\\Desktop\\merged_sheets.xlsx', index=False)

        else:
            #     writer = pd.ExcelWriter("C:\\Users\\Lenovo\\Desktop\\output.xlsx", engine="xlsxwriter")
            #     for file, i in enumerate(self.filenames):
            #         df = pd.read_excel(file)
            #         df.to_excel(writer, sheet_name="Sheet {i + 1}", index=False)

            #         writer.save()
            pass

if __name__ == "__main__":
    MergeXcel()
