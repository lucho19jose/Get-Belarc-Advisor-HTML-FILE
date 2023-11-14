# Inventory Management Script - Belarc Advisor

This script is used to manage an inventory of computer systems. It extracts information about each system and writes it to an Excel file.

## Code Overview

The script reads data from various sources, including system files and hardware information. It then writes this data to an Excel file named 'Inventario.xlsx'.

The data includes:

- File name
- Computer name
- Operating system information
- PC model
- PC serial number
- Processor information
- Mainboard information
- Device names
- RAM
- IP address
- MAC address

The data is written to the Excel file as a new row for each system.

## Usage

Run the script using Node.js. The script will automatically create the 'Inventario.xlsx' file in the same directory.

## Dependencies

This script requires the 'exceljs' module to create the Excel file. Install it using npm:

```bash
npm install exceljs
