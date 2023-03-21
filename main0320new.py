import tkinter as tk
from tkinter import filedialog
from FileAnalyzer import FileAnalyzer

class FileAnalyzerUI:
    def __init__(self,master):
        self.master = master
        master.title("File Analyzer")

        #creat a menu bar
        menubar = tk.Menu(master)
        master.config(menu=menubar)

        # Create a File menu with a "Open" option
        file_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open", command=self.open_file)

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
        self.result_text = tk.Text(master, state="disabled", height=10)
        self.result_text.pack(padx=10, pady=10)

    def open_file(self):
        #Get the file path
        file_path = filedialog.askopenfilename()
        #update the file path label
        self.file_label.config(text=file_path)

    def browse_file(self):
        #Get the file path
        file_path = filedialog.askopenfilename()
        #update the file path label
        self.file_label.config(text=file_path)

    def analyze_file(self):
        #Get the file_path from the file path label
        file_path = self.file_label.cget("text")

        #Create a FileAnalyzer object and analyze the file
        file_analyzer = FileAnalyzer(file_path)

        #Get the result from the analyzer
        result = file_analyzer.run(file_path)

        #Display results
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")

        #Convert each dictionary in the result to a multi-line string
        result_strings = []
        for r in result:
            result_string = ""
            for key, value in r.items():
                result_string += f"{key}: {value}\n"
            result_strings.append(result_string)
  
        #Join the result strings with new line characters and insert into the Text widget
        self.result_text.insert("end", "\n\n".join(result_strings))
        self.result_text.config(state="disabled")
 


#Create the window
root = tk.Tk()

#Create the FileAnalyzerUI object
file_analyzer_ui = FileAnalyzerUI(root)

#Run UI
root.mainloop()