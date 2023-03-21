from GUI.main import App

import os
import pathlib

if __name__ == "__main__":
    root = App()
    root.title('outlook2Teradata')

    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    # Simply set the theme
    root.tk.call("source", os.path.join(pathlib.Path().resolve(),r"src\GUI\azure.tcl"))
    root.tk.call("set_theme", "dark")


    # Set a minsize for the window, and place it in the middle
    root.update()
    
    root.minsize(root.winfo_width(), root.winfo_height())
    x_cordinate = int((root.winfo_screenwidth() / 2) - (root.winfo_width() / 2))
    y_cordinate = int((root.winfo_screenheight() / 2) - (root.winfo_height() / 2))
    root.geometry("+{}+{}".format(x_cordinate, y_cordinate-20))

    root.mainloop()