# Excel Photo Matcher

A Windows desktop tool that inserts product photos into an Excel file by matching each image filename to the **Accounting Code** column.

**Developed by Reza Gholami**

---

## What it does

Many inventory and accounting sheets keep an accounting code for every product, while the photos live in a separate folder and are named after those same codes. Putting the pictures into Excel by hand is slow and error-prone.

This app automates that workflow:

1. You select an Excel workbook (`.xlsx` / `.xlsm`)
2. You select the folder that contains the product photos
3. The program finds the **Accounting Code** column (`کد حسابداری`)
4. For each row, it looks for a photo whose name matches that code
5. The photo is placed in the **Photo** column (`عکس`)
6. A new file is saved next to the original: `original_با_عکس.xlsx`

The original workbook is never overwritten.

## Features

- Persian desktop UI, no command line required
- Standalone `.exe` — end users do not need Python
- Automatic detection of the accounting-code and photo columns
- Creates the photo column if it is missing
- Matches exact names such as `1203001.jpg` and prefixes such as `1203001-front.png`
- Supports Persian and Arabic digits in both Excel cells and filenames
- Keeps image aspect ratio and embeds a higher-resolution copy so photos stay readable
- Progress log plus a list of codes that had no matching photo
- Safe output: writes a new workbook instead of replacing the source file

## Requirements

### To run the ready-made program

- Windows 10 or 11
- The file `افزودن عکس به اکسل.exe`
- Microsoft Excel is **not** required to generate the file; it is only needed if you want to open the result

### To run or build from source

- Python 3.10+
- Packages in `requirements.txt`:
  - `openpyxl`
  - `Pillow`

```bash
pip install -r requirements.txt
python app.py
```

### Input files the user must provide

| Item | Rule |
|------|------|
| Excel file | Must contain a column named **کد حسابداری** (Accounting Code) |
| Photo folder | Image files named after those codes |
| Supported images | `.jpg` `.jpeg` `.png` `.bmp` `.gif` `.webp` `.tif` `.tiff` |
| Example | Code `1203001` matches `1203001.jpg` or `1203001.png` |

Close the Excel file before running the program so it can be saved.

## How to use

1. Double-click `افزودن عکس به اکسل.exe`
2. Choose the Excel file
3. Choose the photos folder
4. Click **شروع و ذخیره اکسل**
5. Open the new file `..._با_عکس.xlsx`

## Build a standalone executable

```bash
pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean ExcelPhoto.spec
```

The executable is created in `dist/ExcelPhoto.exe`.

## Tech stack

- Python
- Tkinter (desktop UI)
- openpyxl (Excel read/write and image embedding)
- Pillow (image orientation, resize, and quality)
- PyInstaller (standalone Windows build)

## Author

**Reza Gholami**

Developed as a practical desktop utility for accounting and inventory teams who need product photos inside Excel without installing extra software.
