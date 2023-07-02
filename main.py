import customtkinter as ctk

class MergeXcel(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MergeXcel")
        self.geometry("500x600")

        header_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        header_frame.place(relx=0.5, rely=0.2, relwidth=0.95, relheight=0.3, anchor=ctk.S)

        header_label = ctk.CTkLabel(header_frame, text="MergeXcel", font=("Lucida Calligraphy", 50))
        header_label.place(relx=0.5, rely=0.65, anchor=ctk.CENTER)

        body_frame = ctk.CTkFrame(self, corner_radius=20, border_width=5)
        body_frame.place(relx=0.5, rely=0.225, relwidth=0.95, relheight=0.76, anchor=ctk.N)

        self.mainloop()

if __name__ == "__main__":
    MergeXcel()
