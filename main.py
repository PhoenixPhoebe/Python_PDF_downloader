import tkinter as tk
from tkinter import ttk
from  program import download_files as dw
import pickle

import time
def main():
    root = tk.Tk()

    shw_det = False



    #app info
    root.title("PDF downloader")
    root.minsize(root.winfo_width(), root.winfo_height())



    #locations 
    tk.Label(root, text="Excel ark med links").grid(column=0, row=0, columnspan=2, sticky=tk.W)
    tk.Entry(root, state="disabled").grid(column=0, row=1)
    tk.Label(root, text="distanitioon for gemte pdf'er").grid(column=0,row=2, columnspan=2, sticky=tk.W)
    tk.Entry(root, textvariable="TEST", state="readonly").grid(column=0, row=3)
    tk.Label(root, text="Excel ark med downloaded status").grid(column=0,row=4, columnspan=2, sticky=tk.W)
    tk.Entry(root, state="disabled").grid(column=0, row=5)
    
    tk.Label(root, text="Kolonne").grid(column=2, row=0)
    col1 = tk.Entry(root, textvariable="AL", width=5).grid(column=2, row=1)
    tk.Label(root, text="kolonne 2").grid(column=3,row=0)
    col2 = tk.Entry(root, width=5).grid(column=3, row=1)


    tk.Button(root, text="Ændre").grid(column=1, row=1)
    tk.Button(root, text="Ændre").grid(column=1, row=3)
    tk.Button(root, text="ændre").grid(column=1, row=5)

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
    

    # Download button
    button = tk.Button(root, text="download pdfs", command = dw.Run)
    button.grid(column=2, row=2, columnspan=2, rowspan=4, sticky=tk.NSEW)

    

    root.mainloop()

if __name__ == '__main__':
    main()