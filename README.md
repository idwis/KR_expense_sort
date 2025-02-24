# KR_expense_sort

Sorts the expenses of Korean ETFs by category.

## Prerequisites

A Python environment is required. Anaconda is recommended on Windows (as it was used for development).

Chrome should be also installed.

## Installation on executables (Recommended)

1. Download the latest version of [ChromeDriver](https://googlechromelabs.github.io/chrome-for-testing/).
3. Unzip the downloaded file into the `KR_expense_sort` directory (the root of this project).
4. Run `pip install selenium pandas xlrd>=2.0.1 openpyxl` to install the required modules.

## Installation on Python

1. Clone the repository.
2. Download the latest version of [ChromeDriver](https://googlechromelabs.github.io/chrome-for-testing/).
3. Unzip the downloaded file into the `KR_expense_sort` directory (the root of this project).
4. Run `pip install selenium pandas xlrd>=2.0.1 openpyxl` to install the required modules.

## How to Run

1. Run `python exec.py`.
2. Select the desired option number and press Enter.
3. Wait a few minutes (approximately 30 seconds to 2 minutes). The results will be available in the `processed_data` directory.