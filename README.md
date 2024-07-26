
---

# NSE Top 5 Gainers and Losers in F&O

This Python project fetches the top 5 gainers and losers from the NSE (National Stock Exchange) Future and Options (F&O) stocks segment. It retrieves the data from the NSE website, processes it, and outputs the symbols and percentage changes of the top gainers and losers.

## Features

- Fetches stock data from NSE's F&O segment.
- Identifies and displays the top 5 gainers (highest percentage change).
- Identifies and displays the top 5 losers (lowest percentage change).
- Automatically fetches and updates the data every 15 minutes during market hours.

## Prerequisites

- Python 3.x
- `requests` library
- `xlwings` library (if you plan to use Excel integration) (optional)
- `json` library (part of Python standard library)

## Installation

1. Clone this repository to your local machine:

   ```bash
   git clone https://github.com/rahulgedam/fetch-nse-top-gainer-loser-fno.git
   cd fetch-nse-top-gainer-loser-fno
   ```

2. Install the required Python packages:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the script to fetch and display the top gainers and losers:

   ```bash
   python fetchglnsefno.py
   ```

2. The script will output the top 5 gainers and losers based on the percentage change (`pChange`) and will update every 15 minutes during market hours.

## Project Structure

- `fetchglnsefno.py`: Main script that fetches and processes the stock data.

## Example Output

```
Top 5 Positive pChange:
Symbol: XYZ, pChange: 5.67
Symbol: ABC, pChange: 4.89
Symbol: DEF, pChange: 4.56
Symbol: GHI, pChange: 4.23
Symbol: JKL, pChange: 4.01

Top 5 Negative pChange:
Symbol: MNO, pChange: -5.67
Symbol: PQR, pChange: -4.89
Symbol: STU, pChange: -4.56
Symbol: VWX, pChange: -4.23
Symbol: YZA, pChange: -4.01
```

## Notes

- Ensure you have an active internet connection to fetch the data from NSE.
- The script sets the necessary headers to mimic a browser request to bypass potential restrictions by NSE.

## License

Public


