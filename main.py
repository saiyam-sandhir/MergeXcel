import customtkinter as ctk

class MergeXcel(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MergeXcel")
        self.geometry("500x600")

        header_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        header_frame.place(relx=0.5, rely=0.2, relwidth=0.95, relheight=0.3, anchor=ctk.S)

        body_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        body_frame.place(relx=0.5, rely=0.225, relwidth=0.95, relheight=0.76, anchor=ctk.N)

        select_files_button = ctk.CTkButton(body_frame, text="SELECT👆")
        select_files_button.place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

        radio_var = ctk.IntVar(value=2)
        
        multi_sheets_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into multiple sheets", variable=radio_var, value=1)
        multi_sheets_radio_button.place(relx=0.3, rely=0.8, anchor=ctk.CENTER)

        single_sheet_radio_button = ctk.CTkRadioButton(body_frame, text="Merge into single sheet", variable=radio_var, value=2)
        single_sheet_radio_button.place(relx=0.7, rely=0.8, anchor=ctk.CENTER)

        merge_button = ctk.CTkButton(body_frame, text="MERGE 📄➕📄")
        merge_button.place(relx=0.5, rely=0.9, anchor=ctk.CENTER)

        self.mainloop()

if __name__ == "__main__":
    MergeXcel()
