# County Setting Totals

County Setting Totals is a Java application that processes county election setup spreadsheets and generates a formatted summary Excel workbook containing machine totals and supply counts for all counties.

The program reads all county `.xlsx` election plan files from a folder, extracts machine totals, and outputs a professionally formatted spreadsheet with alternating colors, filters, frozen headers, and automatic column sizing.

---

# Features

- Reads all county election spreadsheets automatically
- Processes multiple counties at once
- Generates a single summary workbook
- Alphabetically sorts counties
- Automatically calculates:
  - Scan machines
  - ADA machines
  - Print machines
  - Paper rolls
  - Red seals
  - Blue/Yellow/Black seals
- Professional Excel formatting
- Auto-sized columns
- Alternating row colors
- Frozen header row
- Excel filters enabled

---

# Technologies Used

- Java
- Gradle
- Apache POI

---

# Project Structure

```text
CountySettingTotals/
├── build.gradle
├── settings.gradle
├── counties/
│   ├── Adair Election Plan.xlsx
│   ├── Allen Election Plan.xlsx
│   └── ...
└── src/
    └── main/
        └── java/
            └── CountySettingTotals.java
```

---

# Requirements

- Java 21+ (recommended)
- Gradle
- IntelliJ IDEA or VS Code

---

# Setup Instructions

## 1. Clone Repository

```bash
git clone <repository-url>
```

---

## 2. Open Project

Open the project in:

- IntelliJ IDEA
- VS Code

---

## 3. Add County Files

Place all county election spreadsheets inside:

```text
counties/
```

Supported format:

```text
.xlsx
```

Example:

```text
counties/Adair Election Plan.xlsx
```

---

# Build Project

```bash
./gradlew build
```

---

# Run Program

```bash
./gradlew run
```

---

# Output

The program creates:

```text
County_Machine_Totals.xlsx
```

The output spreadsheet contains:

| County | Scan | ADA | Print | Paper Rolls | Red Seals | Blue/Yellow/Black | Notes |
|---|---|---|---|---|---|---|---|

---

# Calculation Rules

## Paper Rolls

```text
Scan + ADA + Print
```

---

## Red Seals

```text
Scan + ADA + Print
```

---

## Blue/Yellow/Black Seals

```text
Total Scan Machines
```

---

# Spreadsheet Formatting

The generated workbook includes:

- Styled headers
- Alternating pastel blue and white rows
- Borders
- Auto filters
- Frozen top row
- Auto-sized columns
- Notes column

---

# Example Console Output

```text
Processed: Adair Election Plan.xlsx
Processed: Allen Election Plan.xlsx
Processed: Anderson Election Plan.xlsx

DONE!
Created spreadsheet: County_Machine_Totals.xlsx
```

---

# Dependencies

Apache POI is automatically installed through Gradle.

Dependency:

```gradle
implementation 'org.apache.poi:poi-ooxml:5.2.5'
```

---

# Author

Brennen Gabriel
