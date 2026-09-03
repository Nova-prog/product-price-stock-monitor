# product-price-stock-monitor

Monitor public product price, stock, and delivery information, then export CSV and Excel files.

This is a small Python tool for tracking products listed on public shopping or supplier websites. It is intended for lightweight personal or internal use, such as checking lab consumables, office supplies, tools, or other regularly purchased items.

## Features

- Reads target products from `products.csv`
- Checks public product pages
- Records price, stock, and delivery information when available
- Writes the latest result to CSV
- Appends historical results to CSV
- Exports an Excel workbook
- Can be run manually or daily with Windows Task Scheduler

## Requirements

- Python 3.10 or later
- Windows, macOS, or Linux
- Python packages listed in `requirements.txt`

## Installation

Clone or download this repository, then install dependencies.

```bash
python -m pip install -r requirements.txt
