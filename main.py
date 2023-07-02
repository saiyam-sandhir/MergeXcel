import customtkinter as ctk

class MergeXcel(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MergeXcel")
        self.geometry("500x600")

        self.mainloop()

if __name__ == "__main__":
    MergeXcel()
