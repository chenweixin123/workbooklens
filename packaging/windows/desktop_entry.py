from multiprocessing import freeze_support

from workbooklens.console import configure_utf8_redirected_streams
from workbooklens.desktop import main

if __name__ == "__main__":
    configure_utf8_redirected_streams()
    freeze_support()
    main()
