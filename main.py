import tkinter as tk
from  program import download_files as dw
def main():
    root = tk.Tk()

    root.title("PDF downloader")
    # Widgets are added here
    button = tk.Button(root, text="download pdfs", command = dw.Run)
    button.pack()

    root.mainloop()

if __name__ == '__main__':
    main()