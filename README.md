# Meraki AP Migration

## Overview
**Meraki AP Migration** (`file_check_and_ap_replacement.py`) is a Python automation tool designed to manage the lifecycle of Cisco Meraki wireless networks. It processes Excel-based work orders to automate network renaming, Access Point (AP) replacements, additions, removals, and relocations.

The script acts as a bridge between Excel planning documents and the Meraki Dashboard API. It enforces strict **business logic validation** before executing any changes to ensure the physical reality (inventory) matches the logical plan (Excel).

## Prerequisites

### 1. Environment
- **Python 3.x**
- Required Libraries:
  ```bash
  pip install meraki pandas openpyxl
  ```

### 2. API Authentication
You must set your Meraki API key as an environment variable:
- **Variable Name**: `MK_CSM_KEY`
- **Value**: Your Meraki Dashboard API Key

### 3. Required Files
- **`RetailerName.xlsx`**: A database file mapping Partner IDs (Column A) to Company Names (Column B). This must be in the same directory as the script.

---

## 📂 Directory Structure
The script relies on a strict folder hierarchy to process files and route them based on the outcome.

```text
.
├── file_check_and_ap_replacement.py
├── RetailerName.xlsx
├── migration_process.log  (Generated)
├── changes_audit.log      (Generated)
└── Root/
    ├── Update floorplans/
    │   ├── Excel to be processed/   <-- INPUT: Files requiring map updates
    │   ├── Update Floorplans/       <-- OUTPUT: Successful files needing review
    │   └── Manual handling/         <-- OUTPUT: Failed validation
    └── No floorplan updates/
        ├── Excel to be processed/   <-- INPUT: Files with only logical changes
        ├── Done/                    <-- OUTPUT: Successful files (No review needed)
        ├── Update Floorplans/       <-- OUTPUT: Files that unexpectedly need review
        └── Manual handling/         <-- OUTPUT: Failed validation
```

---

## 🔍 Validation Logic Breakdown
Before any changes are made, the script runs a **7-Step Validation Process** on every Excel file. If *any* step fails, the file is rejected and moved to `Manual handling`.

### 1. File Integrity & Duplicates
- **Recursive Scan**: The script scans the entire `Root` directory tree.
- **Rule**: If the same filename exists in two different folders (e.g., in `Done` and `Excel to be processed`), execution halts immediately to prevent data overwrites or processing the wrong version.

### 2. Header & Structure
- **Header Search**: Locates the row containing "Network name".
- **Column Check**: Verifies the exact order of: `Network name`, `Access Point`, `Model`, `Serial number`, `Decision`, `Location tag`, `New Model`, `New Serialnumber`, `New Access Point Name`.

### 3. Retailer Verification
- **Partner ID**: Reads the Partner ID from the Excel file (Cell B2).
- **Lookup**: Validates this ID against `RetailerName.xlsx`. If the ID is missing or invalid, the file is rejected.

### 4. Network Consistency ("The Anchor")
- **Anchor Method**: The script picks the first serial number found in the file and queries the API to find its current Network ID. This becomes the "Target Network."
- **Cross-Check**: Every other serial number in the file is checked against the API.
- **Rule**: If any serial in the file belongs to a *different* network than the Anchor, the file is rejected. This prevents cross-contamination between store networks.

### 5. The "Orphan" Check
- **Inventory Scan**: The script queries the live network for *all* currently claimed wireless devices.
- **Comparison**: It compares the live inventory against the serials listed in the Excel file.
- **Rule**: If the live network contains APs that are **not** listed in the Excel file, validation fails. The Excel file must account for every device in the network (even if the decision is just "Keep").

### 6. Device Status Validation
- **Replace/Remove**: The *Old* AP is checked. It generally expects the device to be `offline` or `dormant` before removal to prevent service interruption.
- **Add/New**: The *New* AP is checked. It must be `online` or `alerting` (plugged in) to be successfully provisioned and renamed.

### 7. Final State Sequence Validation
- **Logic**: The script calculates what the network will look like **after** the migration to ensure the AP numbering is contiguous (e.g., AP01, AP02, AP03...).
- **Data Sources**: It extracts the AP number based on the **Decision**:
  - **Keep / Keep/Relocate**: It reads the existing `Access Point` column (e.g., extracting "01" from "Old-Name-AP01").
  - **Replace / Add**: It reads the `New Access Point Name` column (e.g., extracting "02" from "New-Name-AP02").
  - **Remove**: These rows are ignored as they won't exist in the final network.
- **Rule**: It collects all resulting numbers, sorts them, and checks for gaps.
  - *Example Failure*: If the file keeps `AP01` and adds `AP03`, but removes `AP02`, the validation fails because the final sequence `[1, 3]` has a gap.

---

## ⚙️ Migration Decision Process
Once validation passes, the script iterates through the Excel rows and executes actions based on the **Decision** column.

| Decision | Logic & Actions Performed |
| :--- | :--- |
| **Replace** | 1. **Inherit**: Fetches GPS/Map coordinates and notes from the **Old AP**.<br>2. **Update**: Applies these coordinates, the new Name, and Tags to the **New AP**.<br>3. **Queue**: Adds the Old AP to a "Removal Queue" (processed at the end).<br>4. **Audit**: Logs the swap details. |
| **Remove** | 1. **Queue**: Adds the Old AP to the "Removal Queue".<br>2. **Audit**: Logs the pending removal. |
| **Add** | 1. **Claim**: Claims the New AP into the network (if not already there).<br>2. **Update**: Sets the Name and Tags.<br>3. **Flag**: Marks the file for "Floor Plan Review" because the new AP has no previous coordinates to inherit. |
| **Keep** | 1. **Rename**: Enforces the standard naming convention on the existing AP.<br>2. **Retag**: Updates tags based on the `Location tag` column. |
| **Keep/Relocate**| 1. **Process**: Same actions as "Keep".<br>2. **Flag**: Explicitly marks the file for "Floor Plan Review" so an admin can manually move the map marker. |

### Post-Processing: The Removal Queue
To avoid conflicts during processing, devices are not deleted immediately. The script collects all serials marked for removal (via `Remove` or `Replace`) and executes a **Bulk Removal** API call only after all other row operations are successful.

---

## 🚀 Usage Guide

1.  **Prepare Files**: Place your Excel work orders in the `Excel to be processed` folder within the appropriate `Root` subdirectory.
2.  **Run Script**:
    ```bash
    python file_check_and_ap_replacement.py
    ```
3.  **Validation Phase**:
    - The script scans and validates all files.
    - **Failed files** are immediately moved to `Manual handling`. Check `migration_process.log` for the specific reason (e.g., "Orphan AP found", "Partner ID invalid").
    - **Passed files** remain pending.
4.  **Confirmation**: The script reports the number of passed/failed files. Type `yes` to proceed with the migration of the valid files.
5.  **Completion**:
    - Files are moved to `Done` or `Update Floorplans`.
    - Check `changes_audit.log` for a business-readable summary of changes.

## Naming Conventions
The script enforces the following naming standard to ensure global consistency:

- **Network Name**: `Country-Region-PartnerID-CompanyName`
  - *Example*: `IT-LUC-12345-BestRetailer`
- **Access Point Name**: `Country-Region-PartnerID-APxx`
  - *Example*: `IT-LUC-12345-AP01`

---

## ⚠️ Caveats & Best Practices

### 1. No Warranty / Exclusion of Responsibility
**This script modifies production network configurations via the API.**
By using this tool, you acknowledge that you are solely responsible for its execution. The authors and contributors accept no liability for network outages, configuration loss, or hardware issues resulting from the use of this script.

### 2. Test in a Controlled Environment
**Never run this script against a production organization for the first time.**
Always validate the workflow in a "Sandbox" or "Lab" organization first. Use a test Excel file with 1-2 Access Points to verify that the `Replace`, `Remove`, and `Rename` logic behaves exactly as expected in your specific environment.

### 3. Scale Slowly
**Do not process hundreds of sites in a single batch immediately.**
Start with **one** file. Verify the physical results (APs came up, clients connected, maps updated). Then increase to a batch of 5, then 10. Monitor the `migration_process.log` closely for API rate limiting or unexpected timeouts.

### 4. Handling of "Alerting" APs (Strict Default)
**Current Behavior:**
By default, the script enforces a **Strict Health Check**.
- **Logic:** Any AP involved in a `Keep`, `Add`, or `Replace` operation must be **100% `online`** (Green status).
- **Consequence:** If an AP is `alerting` (Orange status—e.g., "Power Supply Low" or "High DNS Latency"), the file **will fail validation** and move to `Manual handling`. This ensures you do not migrate a network that is already in a degraded state.

**The Code Enforcing Strict Mode:**
```python
# ... inside process_file_validation ...

        for idx, row in df.iterrows():
            d = row['Decision']
            old_sn = str(row['Serial number']) if pd.notna(row['Serial number']) else None
            new_sn = str(row['New Serialnumber']) if pd.notna(row['New Serialnumber']) else None

            if new_sn:
                # ... (inventory checks) ...
                st = status_map.get(new_sn, 'unknown')
                
                # STRICT CHECK: Only 'online' is allowed
                if st not in ["online"]: 
                    logging.error(f"{filename}: New Serial {new_sn} status {st} invalid.")
                    return False, None, None

            if old_sn:
                st = status_map.get(old_sn, 'unknown')
                if d in ["Replace", "Remove"] and st not in ["offline", "dormant"]:
                    logging.error(f"{filename}: Serial {old_sn} status {st} invalid for {d}.")
                    return False, None, None
                
                # STRICT CHECK: Only 'online' is allowed for kept devices
                elif d in ["Keep", "Keep/Relocate"] and st not in ["online"]:
                    logging.error(f"{filename}: Serial {old_sn} status {st} invalid for {d}.")
                    return False, None, None
```

---

## 🛠 Option: How to Allow "Alerting" APs
If your environment frequently has minor alerts (e.g., specific traffic shaping warnings) and you wish to proceed with migration despite them, you can **relax the validation logic**.

**To allow `alerting` APs, modify the two lines in `file_check_and_ap_replacement.py` as shown below:**

1.  **Find this line (New APs):**
    ```python
    if st not in ["online"]:
    ```
    **Change to:**
    ```python
    if st not in ["online", "alerting"]:
    ```

2.  **Find this line (Existing APs):**
    ```python
    elif d in ["Keep", "Keep/Relocate"] and st not in ["online"]:
    ```
    **Change to:**
    ```python
    elif d in ["Keep", "Keep/Relocate"] and st not in ["online", "alerting"]:
    ```

---