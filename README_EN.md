# Maximum.md Category Parser

Universal Python script for parsing **any product category** from the **maximum.md** website.  
It collects data from all pages of a category, downloads images, and saves everything into an Excel file.

The category URL is set in the code.

---

## Features

- Opens the specified product category
- Navigates through **all pagination pages**
- Parses only products **in stock**
- Saves the following data for each product:
  - Image
  - Name
  - Current price
  - Old price (if available)
  - Product URL
- Downloads images to the `images/` folder
- Creates an Excel file with images embedded in cells

---

## Setting the Category

In `main.py`, change only this line:

```python
BASE_URL = "https://maximum.md/ru/kompyuternaya-tehnika/monitory/monitory/"
```
You can replace it with any category URL, e.g.:
Laptops
Graphics cards
Appliances
Accessories

Tech Stack

Python
Selenium (Chrome)
WebDriver Manager
Requests
OpenPyXL

Installation
1. Copy the project
git clone <repo_url>
cd maximum_parser

2. Install dependencies
pip install -r requirements.txt

3. Check Chrome

Make sure Google Chrome is installed.
ChromeDriver will be downloaded automatically.

Usage
python main.py


Output:

Images → ./images/

Excel → maximum_monitors.xlsx (filename can be changed)

Features

Uses a real browser to process JS content
Increased wait timeouts for stability
Supports desktop + mobile pagination
Resilient to partially loaded pages

Possible Improvements

Headless mode
Multithreading (multiple browsers)
Multiple categories in one run
Separate Excel sheets for each category
Proxy / anti-blocking
Logging instead of print
