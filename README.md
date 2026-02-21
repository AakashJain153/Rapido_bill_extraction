# Rapido Bill Extractor

A Python automation tool that extracts ride details from Rapido PDF bills, renames them in a structured format, and generates an Excel summary with clickable hyperlinks.

---

## Features

- Extracts:
  - Ride Date
  - Ride ID
  - Vehicle Number
  - Pickup Location
  - Drop Location
  - Fare Amount
- Creates renamed copies of PDFs in `Refined/`
- Generates Excel summary with:
  - Auto column width
  - Clickable hyperlinks to refined PDFs
- Keeps original files untouched

---

## Output Structure

After running:
Selected Folder/
│
├── Original_PDFs.pdf
└── Refined/
	├── YYYYMMDD_Amount.pdf
	├── YYYYMMDD_Amount.pdf
	└── Rapido_Bills_Summary.xlsx