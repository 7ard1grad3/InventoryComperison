from rich import pretty

from lib.excel import Excel

if __name__ == '__main__':
    # Install pretty console
    pretty.install()
    # Do basic validation
    app = Excel()
    app.start_comparison()
    input("Press Enter key to close")

