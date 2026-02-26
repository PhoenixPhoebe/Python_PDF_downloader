import tkinter as tk
def main():
    root = tk.Tk()

    root.title("PDF downloader")
    # Widgets are added here
    button = tk.Button(root, text="download")
    button.pack()

    root.mainloop()

if __name__ == '__main__':
    main()