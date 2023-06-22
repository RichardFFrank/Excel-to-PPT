from Transformations.Excel import parse_excel
from Transformations.PowerPoint import populateSlides
import tkinter as tk
from tkinter import filedialog, messagebox
import Assets.Constants as Constants


class GUI:
    def __init__(self) -> None:
        self.excelFilePath = None
        self.pptFilePath = None

        self.root = tk.Tk()
        self.root.title(Constants.TITLE_HEADLINE)

        self.label = tk.Label(self.root, text=Constants.INSTRUCTIONS, font=("Arial", 25))
        self.label.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        self.button = tk.Button(self.root, text=Constants.EXCEL_LABEL, command=lambda x="XLS" : self.selectFile(selection=x))
        self.button.grid(row=1, column=0, padx=10, pady=10)

        self.excelFileLabelText = tk.StringVar(value=Constants.FILE_NOT_SELECTED)
        self.excelFileLabel = tk.Label(self.root, textvariable=self.excelFileLabelText, font=("Arial", 15))
        self.excelFileLabel.grid(row=1, column=1, padx=10, pady=10)

        self.button = tk.Button(self.root, text=Constants.PPT_LABEL, command=lambda x="PPT" : self.selectFile(selection=x))
        self.button.grid(row=2, column=0, padx=10, pady=10)

        self.pptFileLabelText = tk.StringVar(value=Constants.FILE_NOT_SELECTED)
        self.pptFileLabel = tk.Label(self.root, textvariable=self.pptFileLabelText, font=("Arial", 15))
        self.pptFileLabel.grid(row=2, column=1, padx=10, pady=10)

        self.submitButton = tk.Button(self.root, text=Constants.SUBMIT_BUTTON, command=self.submit)
        self.submitButton.grid(row=3, column=0, padx=10, pady=10)

        self.outputFilePath = tk.StringVar()
        self.outputLabel = tk.Label(self.root, textvariable=self.outputFilePath, font=("Arial", 15))
        self.outputLabel.grid(row=3, column=1, padx=10, pady=10)

        self.root.mainloop()


    def selectFile(self, selection) -> str:
        excelFileTypes = [("Excel files", ".xlsx .xls")]
        pptFileTypes = [("PowerPoint files", ".ppt .pptx")]

        if selection == "PPT":
            self.pptFilePath = self.getFilePath(pptFileTypes)
            if self.pptFilePath:
                self.pptFileLabelText.set(f"Selected file: {self.pptFilePath}")

        elif selection == "XLS":
            self.excelFilePath = self.getFilePath(excelFileTypes)
            if self.excelFilePath:
               self.excelFileLabelText.set(f"Selected file: {self.excelFilePath}")

        else:
            print("need file type")
            return

    
    def getFilePath(self, ft: str) -> str:
        return filedialog.askopenfilename(
                title = "Open a file to upload",
                filetypes=ft
            )


    def submit(self) -> str:
        if not self.excelFilePath and not self.pptFilePath:
            messagebox.showerror(Constants.ERROR_TITLE, Constants.ERROR_TEXT)
        else:
            try: 
                contentDict = parse_excel(filePath=self.excelFilePath)
                self.outputFilePath.set(f"Completed slides saved as: {populateSlides(contentDict, filePath=self.pptFilePath)}")
                self.excelFileLabelText.set(Constants.EXCEL_LABEL)
                self.pptFileLabelText.set(Constants.PPT_LABEL)
            except FileNotFoundError as FNF:
                messagebox.showerror(message=FNF.with_traceback)
            except TimeoutError as TE:
                messagebox.showerror(message=TE.with_traceback)


    # def onClosing(self):
    #     if messagebox.askyesno(title="Quit?", message="Are you sure you want to exit the application?"):
    #         self.root.destroy()
    #         print("User closed the application.")
