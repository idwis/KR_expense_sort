# KR_expense_sort

Sorts the expenses of Korean ETFs by category.

## Prerequisites

*   A Python environment is required. Anaconda is recommended on Windows (as it was used for development).
*   Google Chrome must also be installed.

## Installation (Executable - Recommended)

1.  Download the latest release of the executable from [Releases](https://github.com/idwis/KR_expense_sort/releases).
2.  Download the latest version of [ChromeDriver](https://googlechromelabs.github.io/chrome-for-testing/).
3.  Unzip the ChromeDriver file into the same directory as the executable.

## Installation (Python Source)

1.  Clone the repository.
2.  Download the latest version of [ChromeDriver](https://googlechromelabs.github.io/chrome-for-testing/).
3.  Unzip the downloaded file into the `KR_expense_sort` directory (the root of this project).
4.  Run `pip install selenium pandas xlrd>=2.0.1 openpyxl` to install the required modules.

## How to Run

*   **Executable:** Run `exec.exe`.
*   **Python Source:** Run `python exec.py` in the terminal.

After running, select the desired option number and press Enter. The results will be available in the `processed_data` directory after a few minutes (approximately 30 seconds to 2 minutes).