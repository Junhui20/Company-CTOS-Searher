import logging
from app.gui import CTOSApp

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app = CTOSApp()
    app.mainloop()
