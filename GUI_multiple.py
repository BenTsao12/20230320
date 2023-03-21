import tkinter as tk
from tkinter import filedialog
from FileAnalyzer import FileAnalyzer

class FileAnalyzerUI:
    def __init__(self,master):
        self.master = master
        self.file_paths = []
        master.title("File Analyzer _ ver.2")
       

        #creat a menu bar
        menubar = tk.Menu(master)
        master.config(menu=menubar)

        # Create a Frame to display the file path
        self.file_frame = tk.Frame(master)
        self.file_frame.pack(padx=10, pady=10)

        # Create a Label to display the file path
        self.file_label = tk.Label(self.file_frame, text="Select a file")
        self.file_label.pack(side="left")

        # Create a button to browse for a file
        self.browse_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.browse_button.pack(padx=10, pady=10)

        # Create a button to analyze the file
        self.analyze_button = tk.Button(master, text="Analyze", command=self.analyze_file)
        self.analyze_button.pack(padx=10, pady=10) 

        # Create a Text widget to display the results
        self.result_text = tk.Text(master, state="disabled", height=20)
        self.result_text.pack(padx=20, pady=20)
    

    def browse_file(self):
        #Get the file path and store it to self.file_paths=[]
        self.file_paths = filedialog.askopenfilenames(filetypes=[("PowerPoint files", "*.pptx"), 
                                                           ("PDF files", "*.pdf")])
        #Display the path of file that we upload.        
        self.file_label.config(text='\n'.join(self.file_paths))

    def analyze_file(self):
        #Get the file_path from the file path label
       
        #Create a FileAnalyzer object and analyze the file
        file_analyzer = FileAnalyzer(self.file_paths)

        #Get the result from the analyzer
        results = file_analyzer.run_multiple(self.file_paths)

        #Display results
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.delete("1.0", tk.END) #Remove the content of text for next test.
        
        for result in results:
            self.result_text.insert(tk.END, f"File Name: {result['file_name']}\n")
            self.result_text.insert(tk.END, f"Author: {result['author']}\n")
            self.result_text.insert(tk.END, f"Date: {result['date']}\n")
            self.result_text.insert(tk.END, f"The main content of the report is: {result['The main content of the report is']}\n\n")

'''
analyze_file():
這個方法用來分析選定的文件並將結果顯示在GUI的Text widget中。
line65-68:
首先，從self.file_paths屬性中獲取選定文件的路徑，然後創建一個FileAnalyzer對象，使用run_multiple方法對選定的文件進行分析，並得到結果。
line70-73:
接著，使用self.result_text.config(state="normal")方法將Text widget的狀態設置為“normal”，以便向其中插入文本，然後使用self.result_text.delete("1.0", "end")方法清空Text widget的內容，以便顯示新的結果。
line75-79
接下來，使用一個for循環將每個文件的結果逐一插入到Text widget中，使用insert方法將結果格式化為指定的字符串格式，然後插入到Text widget中。
'''

        

#Create the window
root = tk.Tk()

#Create the FileAnalyzerUI object
file_analyzer_ui = FileAnalyzerUI(root)

#Run UI
root.mainloop()