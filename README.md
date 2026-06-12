## 📊 Champion Skin Summary Dashboard

This script automates the generation of a comprehensive, production-grade Excel tracking ledger by connecting directly to your running League of Legends client via the **League Client Update (LCU) API**. It cross-references your personal inventory with live server database assets to map out your progression toward a complete skin collection.

### Key Features
* **LCU API Live Sync:** Automatically locates your local client install profile, fetches active user credential endpoints, and parses your owned skin collection natively.
* **Hextech Reroll Metrics:** Dynamically reads server-side global skin databases to calculate how many **Rollable** and **Unrollable** skins remain in the wild for every single champion.
* **Completion Analytics:** Includes an instantaneous `% of Skins Owned` tracking column per champion, formatted cleanly as whole numbers.
* **Automated Clean Layouts:** Outputs an executive spreadsheet containing zebra-striped tables, high-contrast structural headers, accounting-style formatting (replacing cluttering zeros with clean dashes `-`), auto-adjusting column widths, and dynamic right-hand champion density distributions.

### Dashboard Architecture
1. **Champion Skin Summary:** An analytical control panel breaking down your ownership density, completion percentages, and dynamic brackets tracking how many champions sit at specific skin thresholds.
2. **Owned Skin Details:** A detailed item master log compiling every owned skin, individual numeric server IDs, and explicit skin tier rarities (Epic, Legendary, Mythic, Ultimate).
