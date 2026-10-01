# auw-financial

## Tools

### `tools/supplier_tracker.py` — monthly supplier spend & rate tracker

Builds an Excel report from the Accounts Inventory ledger that separates price
changes from volume changes, supplier by supplier, and flags spend spikes,
stopped/new suppliers, item rate jumps, and suppliers whose prices rose faster
than the rest of the market.

```
pip install openpyxl
python tools/supplier_tracker.py "Accounts Inventory.xlsx" \
    --system "AUW Financial System.xlsx" \
    --sites Misfa,Nizwa --base 2025-10:2026-02 --to 2026-08 \
    --out supplier_tracker.xlsx
```

`--system` adds purchases entered only in the AUW Financial System (refs starting
`INV/`). Drop it once the ledger has been re-imported, or those rows are counted twice.
Run monthly after the month's invoices are entered; review the **Flags** sheet first.

The **Expected vs Purchased** sheet compares kilograms bought with the kitchen's
weekly usage per site, taken from `tools/consumption_norms.csv` (one row per site and
item group: weekly kg, usable share of purchased weight, item-name keywords). Update
that file when the menu or headcount changes; pass `--norms ""` to skip the sheet.
