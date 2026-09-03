# 📊 Custom Template Data Engine

A simple desktop tool for cleaning, formatting, and exporting Excel and CSV data.

The application provides a spreadsheet-style interface where you can select cells and quickly apply formatting presets or create your own custom formatting rules.

---

## 🚀 Easy Start — Windows `.exe`

You do not need to install Python to use the application.

Download the latest `.exe` from the repository and run it:

```text
DataCleaner.exe
```

Once opened, you can immediately load an Excel or CSV file and start cleaning your data.

> **Note:** Windows may show a security warning when running an unsigned `.exe`.
> If you trust the source, use the Windows option to allow the application to run.

---

## ✨ Features

### 📁 Open Data

Supports:

* `.xlsx`
* `.xls`
* `.csv`

Files are loaded into an interactive spreadsheet-style interface.

### 🧹 Data Cleaning

Quick formatting presets include:

* Clear NaN & Empty Cells
* Title Case
* UPPERCASE
* lowercase
* Letters Only
* Numbers Only

### 🛠️ Custom Formatting

Create your own formatting rules using the Custom Mask system.

### ↩️ Undo

Undo your last changes with support for up to 20 previous states.

### 💾 Export

Save your processed data as:

* Excel (`.xlsx`)
* CSV (`.csv`)

---

## 🔤 Custom Mask System

The Custom Mask feature allows you to extract and format letters and numbers from existing values.

| Code | Description                                           |  Example  |
| :--: | ----------------------------------------------------- | :-------: |
| `/N` | Extract the next number                               |  `/N/N/N` |
| `/L` | Extract the next letter as uppercase                  |   `/L/L`  |
| `/l` | Extract the next letter as lowercase                  |   `/l/l`  |
| `/+` | Extract the remaining characters of the previous type |   `/N/+`  |
| `/*` | Keep the remaining original text                      |  `ID-/*`  |
| `/-` | Stop processing                                       |  `/N/N/-` |
| Text | Printed exactly as written                            | `ID-/N/N` |

## 🔤 Custom Mask Examples

| Input             | Mask           | Output             | Description                                      |
| ----------------- | -------------- | ------------------ | ------------------------------------------------ |
| `AB123456XYZ`     | `/L/L/N/N/N/-` | `AB123`            | Extract 2 letters, then 3 numbers, and stop      |
| `AB123456XYZ`     | `/L/L/N/+`     | `AB123456`         | Extract 2 letters, then all remaining numbers    |
| `ABC123456`       | `/l/l/l/N/N/N` | `abc123`           | Extract letters as lowercase, then 3 numbers     |
| `Order-12345-ABC` | `ID-/N/+`      | `ID-12345`         | Add literal text and extract all numbers         |
| `ABC123-XYZ456`   | `ID-/*`        | `ID-ABC123-XYZ456` | Add literal text and keep the remaining raw text |

---

## 📦 Installation

### Option 1 — Use the `.exe`

Recommended for normal users.

1. Go to the [Releases page](../../releases) and download the latest `DataCleaner.exe` file.
2. Run the application.
3. Open your Excel or CSV file.
4. Select the cells you want to modify.
5. Choose a formatting preset or create a custom mask.
6. Click **Apply to Selected Cells**.
7. Export your finished file.

No Python installation or additional packages are required when using the `.exe`.

### Option 2 — Run the Source Code

The repository also contains the original Python source code for anyone who wants to inspect, modify, or develop the application.

### Requirements

* Python 3
* pandas
* tkinter
* tksheet

Install the dependencies:

```bash
pip install pandas openpyxl tksheet
```

Then run the Python program:

```bash
python main.py
```

Replace `main.py` with the project's actual Python entry-point file if it has a different name.

---

## 🧰 Built With

* **Python**
* **Tkinter** — Desktop user interface
* **tksheet** — Spreadsheet-style data grid
* **Pandas** — Data processing
* **OpenPyXL** — Excel file handling

---

## 📂 Project Structure

The repository contains both the ready-to-use application and its source code.

```text
Custom-Template-Data-Engine/
│
├── DataCleaner.exe    Ready-to-use Windows application
├── Source Code/       Python source code
└── README.md
```

> The exact folder and file names may vary depending on the current repository structure.

---

## ⚠️ Disclaimer

This is a student and learning project created for personal and educational use.

Parts of the project were developed with the assistance of AI tools.

---

## 📚 Project Purpose

The project was created as a practical Python application for working with spreadsheet data.

It combines:

* Desktop GUI development
* Data cleaning
* Excel/CSV processing
* Custom parsing
* File exporting
* Undo functionality

The project is intended to be useful while also serving as a learning example for Python development.