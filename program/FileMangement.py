from tkinter import filedialog
import pickle


def browseForExcel() -> str:
    return filedialog.askopenfilename(initialdir = "/documents",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),("All files","*.*")))
        
def browseFolders() -> str:
    return filedialog.askdirectory(initialdir="/documents", title="Vælg en distination for downloads")


def Get_locs()-> dict:
    try:
        with open('program/location_data.pkl', 'rb') as inf: 
            in_data = pickle.load(inf) 
            return in_data
    except:
        return dict()
    
def Set_locs(data): 

    #pickle and save in a file
    with open('program/location_data.pkl', 'wb') as outf:
        pickle.dump(data, outf) 
        print(data)
    