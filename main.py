import customtkinter as ctk
import pandas as pd

class MergeXcel(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MergeXcel")
        self.geometry("500x600")
        self.after(201, self.iconbitmap(".git\\icon.ico")

        header_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        header_frame.place(relx=0.5, rely=0.2, relwidth=0.95, relheight=0.3, anchor=ctk.S)

        header_label = ctk.CTkLabel(header_frame, text="MergeXcel", font=("Lucida Calligraphy", 50))
        header_label.place(relx=0.5, rely=0.65, anchor=ctk.CENTER)

        body_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        body_frame.place(relx=0.5, rely=0.225, relwidth=0.95, relheight=0.76, anchor=ctk.N)

        self.filenames = []
        select_files_button = ctk.CTkButton(body_frame, text="SELECT👆", command=self.openFileNames)
        select_files_button.place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

        self.radio_var = ctk.IntVar(value=2)
        multi_sheets_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into multiple sheets", variable=self.radio_var, value=1, state=ctk.DISABLED)# disabled for now to avoid complications
        multi_sheets_radio_button.place(relx=0.3, rely=0.8, anchor=ctk.CENTER)
        single_sheet_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into single sheet", variable=self.radio_var, value=2)
        single_sheet_radio_button.place(relx=0.7, rely=0.8, anchor=ctk.CENTER)

        merge_button = ctk.CTkButton(body_frame, text="MERGE 📄➕📄", command=self.mergeFiles)
        merge_button.place(relx=0.5, rely=0.9, anchor=ctk.CENTER)

        self.mainloop()

    def openFileNames(self):
        self.filenames = list(ctk.filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")], multiple=True))

    def mergeFiles(self):
        if self.radio_var.get() == 2:
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
