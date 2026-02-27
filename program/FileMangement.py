from tkinter import filedialog
def browseForExcel() -> str:
    return filedialog.askopenfilename(initialdir = "/documents",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),("All files","*.*")))
    
    #Link_file.set(str(Link_loc.split('/')[-1]))
        
def browseFolders() -> str:
    return filedialog.askdirectory(initialdir="/documents", title="Vælg en distination for downloads")
    