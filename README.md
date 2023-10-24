# Inventory Automation

## Overview

This repository contains an automation script. The script is designed to enhance the efficiency of managing inventory data using Python and the `openpyxl` library. It performs various tasks related to inventory analysis, supplier management, and product data calculation.

## Features

1. **Efficient Data Processing:** The script loads an Excel workbook, processes the inventory data efficiently, and calculates essential inventory metrics.

2. **Supplier Analysis:** It calculates the number of products per supplier and the total value of inventory per supplier, providing valuable insights into supplier performance.

3. **Product Filtering:** The script identifies products with an inventory quantity of less than 10, making it easier to manage low-stock items.

4. **Batch Cell Updates:** Instead of updating cells one by one, it uses batch cell updates for improved performance when writing back to the Excel workbook.

## How to Use

1. **Requirements:**

   - Python 3.x
   - `openpyxl` library (can be installed using `pip install openpyxl`)

2. **Running the Script:**

   - Clone this repository to your local machine.

   - Place your inventory data in an Excel file named "inventory.xlsx" in the root directory.

   - Run the script:

     ```shell
     python main.py
     ```

   - The script will process the data and save the results in a new Excel file named "inventory_with_total_value.xlsx."

3. **Results:**

   - The script will print the following results to the console:
     - Products per supplier
     - Total value of inventory per supplier
     - Products with inventory less than 10
