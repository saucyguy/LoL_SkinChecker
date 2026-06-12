# 📊 League of Legends Champion Skin Summary Dashboard

This Python script automates the generation of a comprehensive, production-grade Excel tracking ledger by connecting directly to your running League of Legends client via the **League Client Update (LCU) API**. It cross-references your personal inventory with live server database assets to map out your progression toward a complete skin collection.

---

### 🚀 Key Features
* **LCU API Live Sync:** Automatically locates your local client install profile, fetches active user credential endpoints, and parses your owned skin collection natively.
* **Hextech Reroll Metrics:** Dynamically reads server-side global skin databases to calculate how many **Rollable** and **Unrollable** skins remain in the wild for every single champion.
* **Completion Analytics:** Includes an instantaneous `% of Skins Owned` tracking column per champion, formatted cleanly as whole numbers.
* **Automated Clean Layouts:** Outputs an executive spreadsheet containing zebra-striped tables, high-contrast structural headers, accounting-style formatting (replacing cluttering zeros with clean dashes `-`), auto-adjusting column widths, and dynamic right-hand champion density distributions.

---

### 📂 Dashboard Architecture

The output workbook is split into two cleanly formatted tabs (with default Excel gridlines hidden for a cleaner look):

1. **Champion Skin Summary:** An analytical control panel breaking down your ownership density, completion percentages, and dynamic brackets tracking how many champions sit at specific skin thresholds. It also features a margin spacer column separating the master table from a right-side list of champions sorted by exactly how many skins you own for them.
2. **Owned Skin Details:** A detailed item master log compiling every owned skin, individual numeric server IDs, and explicit skin tier rarities (Epic, Legendary, Mythic, Ultimate).

---
## 🛠️ How to Use

### 1. Prerequisites

Ensure Python is installed, along with the required packages. Install the script dependencies via pip:

```bash
pip install requests openpyxl
```

### 2. Launch the League Client

The League of Legends client must be **open and logged in** before running the script. The script relies on the local League Client Update (LCU) API, which is only available while a client session is active.

### 3. Configure Installation Paths

If you installed League of Legends in a custom directory, open `skin_checker.py` and add your path to the `possible_paths` array inside the `get_lockfile_data()` function so the script can locate your client lockfile:

```python
possible_paths = [
    "YOUR_CUSTOM_PATH/League of Legends/lockfile",
    "D:/P1/Riot Games/League of Legends/lockfile",
    "C:/Riot Games/League of Legends/lockfile",
    "D:/Riot Games/League of Legends/lockfile",
]
```

### 4. Run the Script

Execute the application from your terminal or command prompt:

```bash
python skin_checker.py
```

### 5. View Your Dashboard

Once the script finishes, a structured Excel file named `my_league_skins.xlsx` is generated in the project directory, fully populated with your live data.
