from multiprocessing import freeze_support

from workbooklens.cli import app
from workbooklens.console import configure_utf8_redirected_streams


def main() -> None:
    configure_utf8_redirected_streams()
    freeze_support()
    app()


if __name__ == "__main__":
    main()
