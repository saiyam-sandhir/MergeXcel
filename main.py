import customtkinter as ctk
from PIL import Image
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

        logo_frame = ctk.CTkFrame(header_frame, width=400, height=100, border_width=0, fg_color="transparent")
        logo_frame.place(relx=0.5, rely=0.65, anchor=ctk.CENTER)

        logo_image = ctk.CTkImage(light_image=Image.open(".\\icon.png"),dark_image=Image.open(".\\icon.png"),size=(100, 100))
        logo_label = ctk.CTkLabel(logo_frame, image=logo_image, text="")
        logo_label.place(relx=0.15, rely=0.5, anchor=ctk.CENTER)

        header_label = ctk.CTkLabel(logo_frame, text="MergeXcel", font=("Lucida Calligraphy", 50))
        header_label.place(relx=0.265, rely=0.5, anchor=ctk.W)
        
        #-------------------- BODY --------------------#
        body_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        body_frame.place(relx=0.5, rely=0.225, relwidth=0.95, relheight=0.76, anchor=ctk.N)

        self.filenames = []
        self.select_files_button = ctk.CTkButton(body_frame, text="SELECT", font=("Arial", 20), command=self.openFileNames)
        self.select_files_button.place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

        self.select_files_progress_bar = ctk.CTkProgressBar(body_frame, height=10, border_width=2, corner_radius=5, orientation="indeterminate")
        self.select_files_progress_bar.place(relx=0.5, rely=0.1, relwidth=0.6, anchor=ctk.CENTER)
        self.select_files_progress_bar.place_forget()

        self.selected_files_listbox = ctkLB.CTkListbox(body_frame)
        self.selected_files_listbox.place(relx=0.5, rely=0.45, relwidth=0.6, relheight=0.55, anchor=ctk.CENTER)

        self.radio_var = ctk.IntVar(value=2)
        multi_sheets_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into multiple sheets", variable=self.radio_var, value=1, state=ctk.DISABLED)# disabled for now to avoid complications
        multi_sheets_radio_button.place(relx=0.3, rely=0.8, anchor=ctk.CENTER)
        single_sheet_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into single sheet", variable=self.radio_var, value=2)
        single_sheet_radio_button.place(relx=0.7, rely=0.8, anchor=ctk.CENTER)

        self.merge_button = ctk.CTkButton(body_frame, text="MERGE", font=("Arial", 20), command=self.mergeFiles)
        self.merge_button.place(relx=0.5, rely=0.9, anchor=ctk.CENTER)

        self.merge_files_progress_bar = ctk.CTkProgressBar(body_frame, height=10, border_width=2, corner_radius=5, orientation="indeterminate")
        self.merge_files_progress_bar.place(relx=0.5, rely=0.9, relwidth=0.6, anchor=ctk.CENTER)
        self.merge_files_progress_bar.place_forget()

        self.merge_message_label = ctk.CTkLabel(body_frame, text="", font=("Arial", 10), height=12)
        self.merge_message_label.place(relx=0.5, rely=0.96, anchor=ctk.CENTER)
        self.merge_message_label.place_forget()

        self.mainloop()
        
    def openFileNames(self):
        new_filenames = list(ctk.filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")], multiple=True))

        self.select_files_button.place_forget()
        self.select_files_progress_bar.place(relx=0.5, rely=0.1, relwidth=0.6, anchor=ctk.CENTER)
        self.select_files_progress_bar.start()
        
        if len(new_filenames) != 0:
            self.filenames = new_filenames
            for i in range(self.selected_files_listbox.size()):
                self.selected_files_listbox.delete(i)
                
            for i, file in enumerate(self.filenames):
                self.selected_files_listbox.insert(i, file)

        self.select_files_progress_bar.stop()
        self.select_files_progress_bar.place_forget()
        self.select_files_button.place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

    def mergeFiles(self):
        try:
            if self.radio_var.get() == 2 and len(self.filenames) != 0:
                self.merge_button.place_forget()
                self.merge_files_progress_bar.place(relx=0.5, rely=0.9, relwidth=0.6, anchor=ctk.CENTER)
                self.merge_files_progress_bar.start()

                merged_df = pd.read_excel(self.filenames[0])

                for file in self.filenames[1:]:
                    df = pd.read_excel(file)
                        
                    merged_df = pd.merge(merged_df, df, how="outer")

                self.merge_files_progress_bar.stop()
                self.merge_files_progress_bar.place_forget()
                self.merge_button.place(relx=0.5, rely=0.9, anchor=ctk.CENTER)

                save_loc = ctk.filedialog.asksaveasfile(filetypes=[("Excel file", "*.xlsx")], defaultextension=[("Excel file", ".xlsx")])
                merged_df.to_excel(save_loc.name, index=False)

            else:
                #   writer = pd.ExcelWriter("C:\\Users\\Lenovo\\Desktop\\output.xlsx", engine="xlsxwriter")
                #   for file, i in enumerate(self.filenames):
                #       df = pd.read_excel(file)
                #       df.to_excel(writer, sheet_name="Sheet {i + 1}", index=False)

                #       writer.save()
                pass

            self.merge_message_label.configure(text="Merge Successful", text_color="green", font=("Arial", 10))
            self.merge_message_label.place(relx=0.5, rely=0.96, anchor=ctk.CENTER)
            self.after(5000, self.merge_message_label.place_forget)

        except Exception:
            self.merge_message_label.configure(text=f"Error: {Exception}", text_color="red", font=("Arial", 10))
            self.merge_message_label.place(relx=0.5, rely=0.96, anchor=ctk.CENTER)
            self.after(5000, self.merge_message_label.place_forget)

if __name__ == "__main__":
    MergeXcel()
