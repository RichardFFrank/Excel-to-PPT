from Transformations.Excel import parse_excel
from Transformations.PowerPoint import populateSlides
import tkinter as tk
from tkinter import filedialog, messagebox


class GUI:
    def __init__(self) -> None:
        self.excelFilePath = None
        self.pptFilePath = None
        
        self.root = tk.Tk()
        self.root.title("Excel to PPT")

        self.label = tk.Label(self.root, text="Select the template file of your powerpoint and the excel file to parse using the buttons below.", font=("Arial", 25))
        self.label.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        self.button = tk.Button(self.root, text="Select Excel File", command=lambda x="XLS" : self.selectFile(selection=x))
        self.button.grid(row=1, column=0, padx=10, pady=10)

        self.excelFileLabelText = tk.StringVar(value="No excel file selected.")
        self.excelFileLabel = tk.Label(self.root, textvariable=self.excelFileLabelText, font=("Arial", 15))
        self.excelFileLabel.grid(row=1, column=1, padx=10, pady=10)

        self.button = tk.Button(self.root, text="Select PPT File", command=lambda x="PPT" : self.selectFile(selection=x))
        self.button.grid(row=2, column=0, padx=10, pady=10)

        self.pptFileLabelText = tk.StringVar(value="No powerpoint file selected.")
        self.pptFileLabel = tk.Label(self.root, textvariable=self.pptFileLabelText, font=("Arial", 15))
        self.pptFileLabel.grid(row=2, column=1, padx=10, pady=10)

        self.submitButton = tk.Button(self.root, text="Press to perform transformation.", command=self.submit)
        self.submitButton.grid(row=3, column=0, padx=10, pady=10)

        self.outputFilePath = tk.StringVar()
        self.outputLabel = tk.Label(self.root, textvariable=self.outputFilePath, font=("Arial", 15))
        self.outputLabel.grid(row=3, column=1, padx=10, pady=10)

        self.root.mainloop()


    def selectFile(self, selection) -> str:
        excelFileTypes = [("Excel files", ".xlsx .xls")]
        pptFIleTypes = [("PowerPoint files", ".ppt .pptx")]

        if selection == "PPT":
            self.pptFilePath = self.getFilePath(pptFIleTypes)
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
            messagebox.showerror("File Selection Incomplete", "Please select both files before submitting.")
        else:
            try: 
                contentDict = parse_excel(filePath=self.excelFilePath)
                self.outputFilePath.set(f"Completed slides saved as: {populateSlides(contentDict, filePath=self.pptFilePath)}")
            except FileNotFoundError:
                messagebox.showerror(message=FileNotFoundError.with_traceback)
            except TimeoutError:
                messagebox.showerror(TimeoutError.with_traceback)



    # def onClosing(self):
    #     if messagebox.askyesno(title="Quit?", message="Are you sure you want to exit the application?"):
    #         self.root.destroy()
    #         print("User closed the application.")
