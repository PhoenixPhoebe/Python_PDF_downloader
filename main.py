import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from  program import download_files as dw
from program import FileMangement as fm
import pickle


import time
def main():
    root = tk.Tk()

    shw_det = False
    Link_loc = ""
    Link_file = tk.StringVar()
    Dwn_loc = ""
    Dwn_name = tk.StringVar()
    MD_File = ""
    MD_File_name = tk.StringVar()

    #app info
    root.title("PDF downloader")
    root.minsize(root.winfo_width(), root.winfo_height())

    try:
        with open('location_data.pkl', 'rb') as inf: 
            in_data = pickle.load(inf) 
            print(in_data)
    except:
        pass


    def Change_link_excel():
        global Link_loc 
        Link_loc = fm.browseForExcel()

        Link_file.set(str(Link_loc.split('/')[-1]))
        
    def Change_dwn_folder():
        global Dwn_loc
        Dwn_loc = fm.browseFolders()
        Dwn_name.set(str(Dwn_loc.split('/')[-1]))

    def Change_data_ecxel():
        global MD_File
        MD_File = fm.browseForExcel()
        MD_File_name.set(str(MD_File.split('/')[-1]))


    ###locations
    # Links location 
    tk.Label(root, text="Excel ark med links").grid(column=0, row=0, columnspan=2, sticky=tk.W)
    tk.Entry(root, state="readonly", textvariable=Link_file).grid(column=0, row=1)
    tk.Button(root, text="Ændre", command=Change_link_excel).grid(column=1, row=1)

    #Column for finding links
    tk.Label(root, text="Kolonne").grid(column=2, row=0)
    col1 = tk.Entry(root, textvariable="AL", width=5).grid(column=2, row=1)
    tk.Label(root, text="kolonne 2").grid(column=3,row=0)
    col2 = tk.Entry(root, width=5).grid(column=3, row=1)
    
    #Place to save pdf's
    tk.Label(root, text="distanitioon for gemte pdf'er").grid(column=0,row=2, columnspan=2, sticky=tk.W)
    tk.Entry(root, textvariable=Dwn_name, state="readonly").grid(column=0, row=3)
    tk.Button(root, text="Ændre", command=Change_dwn_folder).grid(column=1, row=3)

    #Place to write if pdf's is downloaded
    tk.Label(root, text="Excel ark med downloaded status").grid(column=0,row=4, columnspan=2, sticky=tk.W)
    tk.Entry(root, state="readonly", textvariable=MD_File_name).grid(column=0, row=5)
    tk.Button(root, text="Ændre", command=Change_data_ecxel).grid(column=1, row=5)

    #check box to see detalied status
    tk.Checkbutton(root, text="Vis Detaljeret status", variable=shw_det).grid(column=0, row=6, sticky=tk.W)

    tk.Label(root, text="").grid(column=0, row=6, sticky=tk.W)
    progress_i = ttk.Progressbar(root, orient='horizontal', mode='indeterminate')
    progress_i.grid(column=0, row=7, columnspan=5, sticky=tk.EW)
    


    """
    def test():
        
        progress_i.start(10)
        root.after(5000, progress_i.stop)
        root.after(2000, test2)
        
    def test2():
        
        progress = ttk.Progressbar(root, orient='horizontal', mode='determinate', value=0)

        progress.grid(column=0, row=7, columnspan=5, sticky=tk.EW)

        for value in (20, 40, 50, 60, 80, 100):
            progress['value'] = value
            root.update_idletasks()
            time.sleep(1)
    """
    
    def download():
        pass


    

    # Download button
    button = tk.Button(root, text="download pdfs", command = download)
    button.grid(column=2, row=2, columnspan=2, rowspan=4, sticky=tk.NSEW)

    

    root.mainloop()

if __name__ == '__main__':
    main()