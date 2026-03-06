from tkinter import filedialog
import pickle

LocationData = 'program/location_data.pkl'

def browseForExcel() -> str:
    return filedialog.askopenfilename(initialdir = "/documents",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),("All files","*.*")))
        
def browseFolders() -> str:
    return filedialog.askdirectory(initialdir="/documents", title="Vælg en distination for downloads")


def Get_locs(location = LocationData)-> dict:
    try:
        with open(location, 'rb') as inf: 
            in_data = pickle.load(inf) 
            return in_data
    except:
        return dict()
    
def Set_locs(data, location = LocationData): 

    #pickle and save in a file
    with open(location, 'wb') as outf:
        pickle.dump(data, outf) 
        
    