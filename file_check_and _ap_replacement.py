import os
import shutil
import pandas as pd
import meraki
import logging
import sys
import re
import time
import unicodedata
from datetime import datetime


# --- CONFIGURATION & CONSTANTS ---
API_KEY = os.environ.get("MK_CSM_KEY", "YOUR_API_KEY_HERE")

# Base Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RETAILER_DB_FILE = os.path.join(BASE_DIR, 'RetailerName.xlsx')
PROCESS_LOG_FILE = os.path.join(BASE_DIR, 'migration_process.log')
AUDIT_LOG_FILE = os.path.join(BASE_DIR, 'changes_audit.log')

# Folder Structure Definitions
ROOT_FOLDER = os.path.join(BASE_DIR, 'Root')
PATH_UPDATE_FP = os.path.join(ROOT_FOLDER, 'Update floorplans')
PATH_NO_UPDATE_FP = os.path.join(ROOT_FOLDER, 'No floorplan updates')

DIR_MANUAL_HANDLING = "Manual handling"
DIR_DONE = "Done"
DIR_OUTPUT_UPDATE_FP = "Update Floorplans"

PATHS_TO_PROCESS = [
    os.path.join(PATH_UPDATE_FP, 'Excel to be processed'),
    os.path.join(PATH_NO_UPDATE_FP, 'Excel to be processed')
]

# Validation Lists
MODELS_TO_SWAP = [] 
MODELS_TO_KEEP = []

# Excel Headers
EXPECTED_HEADERS = [
    "Network name", "Access Point", "Model", "Serial number", 
    "Decision", "Location tag", "New Model", "New Serialnumber", "New Access Point Name"
]
EXPECTED_HEADERS_LOWER = [h.lower() for h in EXPECTED_HEADERS]

# Logic Constants
VALID_DECISIONS = ["Replace", "Remove", "Keep", "Keep/Relocate", "Add"]
VALID_LOCATIONS = ["Workshop", "Showroom", "Outdoors"]
NEW_AP_TAG = "NEW-AP"

# --- CUSTOM LOGGING FORMATTER ---
class AsciiSafeFormatter(logging.Formatter):
    """
    Custom logging formatter that automatically normalizes text 
    (converts 'é' to 'e') and removes unsupported characters 
    to prevent Windows console crashes.
    """
    def format(self, record):
        # 1. Format the message normally (adds timestamp, level, etc.)
        original_msg = super().format(record)
        
        # 2. Normalize Unicode (NFD form separates 'é' into 'e' + '´')
        nfkd_form = unicodedata.normalize('NFD', original_msg)
        
        # 3. Filter out non-spacing marks (accents) and reconstruct string
        ascii_text = "".join([c for c in nfkd_form if not unicodedata.category(c).startswith('Mn')])
        
        # 4. Remove the specific replacement character that caused your crash
        clean_msg = ascii_text.replace('\ufffd', '')
        
        return clean_msg

# --- LOGGING SETUP ---
def setup_logging():
    """Configures logging to both file and console with auto-normalization."""
    
    # Use our custom formatter instead of the default logging.Formatter
    log_formatter = AsciiSafeFormatter('%(asctime)s - %(levelname)s - %(message)s')
    
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Clear existing handlers to prevent duplicates if function is called twice
    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    # File Handler
    # We still use utf-8 for the file because files support it better than consoles
    file_handler = logging.FileHandler(PROCESS_LOG_FILE, mode='a', encoding='utf-8')
    file_handler.setFormatter(log_formatter)
    root_logger.addHandler(file_handler)

    # Console Handler
    # The custom formatter ensures only ASCII characters reach the console
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    root_logger.addHandler(console_handler)

    logging.info("\n" + "="*80)
    logging.info(f"RUN STARTED AT: {datetime.now()}")
    logging.info("="*80)

def append_audit_log(entry_text):
    """Appends a detailed entry to the rolling audit log file."""
    try:
        with open(AUDIT_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(entry_text + "\n")
    except Exception as e:
        logging.error(f"Failed to write to Audit Log: {e}")

# --- HELPER FUNCTIONS ---

def show_user_organizations(dashboard):
    """
    Validates API access and displays available organizations.
    Returns the list of organizations.
    """
    logging.info("Verifying API Key access...")
    try:
        orgs = dashboard.organizations.getOrganizations()
        orgs_sorted = sorted(orgs, key=lambda x: x['name'])        
        
        # Log summary to file
        log_summary = f"API Key validated. Access granted to {len(orgs)} organizations:\n"
        for org in orgs_sorted:
            log_summary += f"  - {org['name']} (ID: {org['id']})\n"
        
        logging.info(log_summary.strip())
        return orgs

    except Exception as e:
        logging.critical(f"Failed to validate API Key or fetch Organizations: {e}")
        print("\nCRITICAL ERROR: Unable to fetch organizations. Check your API Key.")
        sys.exit(1)

def check_for_duplicate_filenames_recursive(root_directory):
    """
    Recursively checks for duplicate filenames across the directory tree 
    to prevent data overwrites or processing errors.
    """
    logging.info(f"Scanning for duplicate filenames recursively in: {root_directory}")
    
    file_map = {}
    duplicates_found = False

    for dirpath, _, filenames in os.walk(root_directory):
        for f in filenames:
            if f.endswith((".xlsx", ".xls")):
                full_path = os.path.join(dirpath, f)
                
                if f not in file_map:
                    file_map[f] = []
                file_map[f].append(full_path)

    for filename, paths in file_map.items():
        if len(paths) > 1:
            duplicates_found = True
            logging.error(f"DUPLICATE FILE FOUND: '{filename}'")
            for p in paths:
                logging.error(f" - Location: {p}")

    if duplicates_found:
        return False
    
    return True

def move_processed_file(file_path, df, success, review_items=None):
    """
    Moves processed files to the appropriate destination folder based on 
    success status and floor plan review requirements.
    """
    filename = os.path.basename(file_path)
    parent_dir = os.path.dirname(file_path) 
    category_dir = os.path.dirname(parent_dir) 
    category_name = os.path.basename(category_dir)

    target_folder_name = ""

    if not success:
        target_folder_name = DIR_MANUAL_HANDLING
    else:
        # Determine target based on category and content
        if category_name == os.path.basename(PATH_UPDATE_FP): 
            target_folder_name = DIR_OUTPUT_UPDATE_FP
        elif category_name == os.path.basename(PATH_NO_UPDATE_FP):
            decisions = df['Decision'].unique().tolist()
            has_review_items = bool(review_items)
            
            # If review is needed or specific decisions exist, force "Update Floorplans"
            if has_review_items or "Keep/Relocate" in decisions or "Add" in decisions:
                target_folder_name = DIR_OUTPUT_UPDATE_FP
            else:
                target_folder_name = DIR_DONE
        else:
            logging.warning(f"Unknown folder category '{category_name}'. Moving to Manual handling.")
            target_folder_name = DIR_MANUAL_HANDLING

    target_path = os.path.join(category_dir, target_folder_name)
    if not os.path.exists(target_path):
        os.makedirs(target_path)

    # Ensure unique filename to prevent overwrites
    base_name, extension = os.path.splitext(filename)
    counter = 1
    unique_filename = filename
    
    while os.path.exists(os.path.join(target_path, unique_filename)):
        unique_filename = f"{base_name}_{counter}{extension}"
        counter += 1

    dest_file_path = os.path.join(target_path, unique_filename)

    try:
        shutil.move(file_path, dest_file_path)
        logging.info(f"Moved file to: {target_folder_name} ({unique_filename})")
    except Exception as e:
        logging.error(f"CRITICAL: Could not move file {filename}: {e}")

def load_retailer_db():
    """Loads and validates the Retailer Name Excel database."""
    if not os.path.exists(RETAILER_DB_FILE):
        logging.critical(f"Retailer DB file missing: {RETAILER_DB_FILE}")
        sys.exit(1)
    try:
        df = pd.read_excel(RETAILER_DB_FILE, header=None)
        
        # Validate uniqueness of Partner IDs
        ids = df.iloc[:, 0].astype(str).str.strip()
        if not ids.is_unique:
            logging.critical("Retailer DB Validation Failed: Partner IDs are not unique.")
            sys.exit(1)
            
        retailer_map = {}
        for index, row in df.iterrows():
            p_id = str(row.iloc[0]).strip()
            c_name = str(row.iloc[1]).strip()
            
            if not p_id or not c_name or c_name.lower() == 'nan':
                 logging.critical(f"Retailer DB Validation Failed: Empty data at row {index}")
                 sys.exit(1)
            retailer_map[p_id] = c_name
        return retailer_map
    except Exception as e:
        logging.critical(f"Error reading Retailer DB: {e}")
        sys.exit(1)

def get_global_inventory(dashboard):
    """Fetches the complete wireless inventory for all organizations."""
    logging.info("Fetching global Meraki inventory...")
    inventory = {}
    try:
        orgs = dashboard.organizations.getOrganizations()
        for org in orgs:
            org_id = org['id']
            devices = dashboard.organizations.getOrganizationDevices(
                org_id, total_pages='all', productTypes=['wireless']
            )
            for dev in devices:
                dev['orgId'] = org_id
                inventory[dev['serial']] = dev
    except meraki.APIError as e:
        logging.critical(f"Meraki API Error: {e}")
        sys.exit(1)
    return inventory

def get_device_statuses(dashboard, org_id, serials):
    """Retrieves availability status (online/offline) for specific serials."""
    if not serials: return {}
    try:
        statuses = dashboard.organizations.getOrganizationDevicesAvailabilities(
            org_id, total_pages='all', serials=serials
        )
        return {item['serial']: item['status'] for item in statuses}
    except Exception as e:
        logging.error(f"Error fetching statuses for Org {org_id}: {e}")
        return {}

def validate_final_ap_sequence(df):
    """
    Checks if the resulting AP numbering (AP01, AP02...) contains duplicates.
    Gaps in the sequence (e.g., 01, 02, 05) are allowed.
    """
    seen_numbers = set()
    regex = re.compile(r'AP(\d{2})$')
    
    for index, row in df.iterrows():
        decision = str(row['Decision']).strip()
        target_name = ""
        
        # Determine the target name based on the decision
        if decision in ["Keep", "Keep/Relocate"]:
            target_name = str(row['Access Point']).strip()
        elif decision in ["Replace", "Add"]:
            target_name = str(row['New Access Point Name']).strip()
        elif decision == "Remove":
            continue 
        
        # Extract the AP number (e.g., '01' from 'AP01')
        match = regex.search(target_name)
        if match:
            ap_num = match.group(1)
            
            # Check for duplicates
            if ap_num in seen_numbers:
                return False, f"Duplicate AP number found: AP{ap_num} (Check row {index + 2})"
            
            seen_numbers.add(ap_num)
    
    if not seen_numbers:
        return False, "No APs found in final configuration."
    
    # If we reached here, there are no duplicates. 
    # We no longer compare against a range, so gaps are perfectly fine.
    return True, f"Validation OK: {len(seen_numbers)} unique APs identified (Gaps allowed)."

def sanitize_network_name(name):
    """Sanitizes network name to comply with Meraki API constraints."""
    # Allow only alphanumeric, whitespace, and specific symbols (. @ # -)
    name = re.sub(r'[^\w\s.@#-]', '', name)
    
    # Collapse multiple spaces
    name = re.sub(r'\s+', ' ', name)
    
    # Trim and truncate
    name = name.strip()
    return name[:128]

# --- MIGRATION LOGIC (INTEGRATED) ---

def migrate_meraki_aps_internal(dashboard, network_id, ap_df, naming_prefix, delay=0, new_ap_tag="NEW-AP"):
    """
    Executes the AP migration logic based on the Excel decisions.
    Updates names, tags, and floor plan positions.
    Returns audit logs and a list of items requiring manual floor plan review.
    """
    audit_actions = []
    fp_review_list = []
    floor_plan_cache = {}
    
    serials_to_remove = []

    def get_floor_plan_id(net_id, serial):
        if net_id not in floor_plan_cache:
            try:
                fps = dashboard.networks.getNetworkFloorPlans(net_id)
                mapping = {}
                for fp in fps:
                    fp_id = fp.get('floorPlanId')
                    for dev in fp.get('devices', []):
                        if dev.get('serial'):
                            mapping[dev['serial']] = fp_id
                floor_plan_cache[net_id] = mapping
            except meraki.APIError:
                floor_plan_cache[net_id] = {}
        return floor_plan_cache[net_id].get(serial)

    def construct_ap_name(source_string):
        source_string = source_string.strip()
        match = re.search(r'(AP\d{2})(?:-N)?$', source_string)
        if match:
            suffix = match.group(1) 
            return f"{naming_prefix}-{suffix}"
        else:
            return f"{naming_prefix}-{source_string}"

    # --- PROCESS ROWS ---
    for index, row in ap_df.iterrows():
        try:
            old_serial = str(row['Serial number']).strip() if pd.notna(row['Serial number']) else ""
            old_ap_name = str(row['Access Point']).strip() if pd.notna(row['Access Point']) else "Unknown-Name"
            decision = str(row['Decision']).strip()
            loc_tag = str(row['Location tag']).strip() if pd.notna(row['Location tag']) else ""
            new_serial = str(row['New Serialnumber']).strip() if pd.notna(row['New Serialnumber']) else ""
            new_ap_input = str(row['New Access Point Name']).strip() if pd.notna(row['New Access Point Name']) else ""
            
            logging.info(f"Processing {decision} for Serial: {old_serial}")

            # --- DECISION: REMOVE ---
            if decision == "Remove":
                serials_to_remove.append(old_serial)
                msg = f"MARKED FOR REMOVAL: {old_serial} (Old Name: '{old_ap_name}')"
                logging.info(msg)
                audit_actions.append(msg)

            # --- DECISION: REPLACE ---
            elif decision == "Replace":
                constructed_name = construct_ap_name(new_ap_input)
                
                source_ap = dashboard.devices.getDevice(old_serial)
                target_fp_id = get_floor_plan_id(network_id, old_serial)
                
                # Prepare tags
                new_device_info = dashboard.devices.getDevice(new_serial)
                tags = new_device_info.get('tags', [])
                new_device_name = new_device_info.get('name', [])

                if loc_tag and loc_tag not in tags: tags.append(loc_tag)
                #if new_ap_tag and new_ap_tag not in tags: tags.append(new_ap_tag)

                # Update new device with old device's position
                update_params = {
                    'serial': new_serial,
                    'name': constructed_name,
                    'lat': source_ap.get('lat'),
                    'lng': source_ap.get('lng'),
                    'notes': source_ap.get('notes'),
                    'address': source_ap.get('address'),
                    'tags': tags, 
                    'moveMapMarker': True
                }
                if target_fp_id: update_params['floorPlanId'] = target_fp_id
                
                dashboard.devices.updateDevice(**update_params)
                serials_to_remove.append(old_serial)
                
                msg = f"REPLACED: {old_serial} ('{old_ap_name}') with {new_serial} (Renamed '{new_device_name}' -> '{constructed_name}'), Tags: {tags}. Old serial marked for removal."
                logging.info(msg)
                audit_actions.append(msg)

                if not target_fp_id:
                    fp_review_list.append({
                        "serial": new_serial,
                        "name": constructed_name,
                        "reason": "Source AP had no Floor Plan"
                    })

            # --- DECISION: KEEP / RELOCATE ---
            elif decision in ["Keep", "Keep/Relocate"]:
                constructed_name = construct_ap_name(old_ap_name)
                
                current_ap = dashboard.devices.getDevice(old_serial)
                tags = current_ap.get('tags', [])
                if loc_tag and loc_tag not in tags: tags.append(loc_tag)

                dashboard.devices.updateDevice(serial=old_serial, name=constructed_name, tags=tags)
                
                if old_ap_name != constructed_name:
                    msg = f"KEPT: {old_serial} (Renamed: '{old_ap_name}' -> '{constructed_name}', Tags: {tags})"
                else:
                    msg = f"KEPT: {old_serial} (Name: '{constructed_name}', Tags: {tags})"
                logging.info(msg)
                audit_actions.append(msg)

                current_fp_id = get_floor_plan_id(network_id, old_serial)
                
                if decision == "Keep/Relocate":
                    fp_review_list.append({
                        "serial": old_serial,
                        "name": constructed_name,
                        "reason": "Relocation requested (Verify Position)"
                    })
                elif decision == "Keep" and not current_fp_id:
                    fp_review_list.append({
                        "serial": old_serial,
                        "name": constructed_name,
                        "reason": "Not assigned to any Floor Plan"
                    })

            # --- DECISION: ADD ---
            elif decision == "Add":
                constructed_name = construct_ap_name(new_ap_input)
                
                try:
                    dashboard.networks.claimNetworkDevices(network_id, [new_serial])
                except:
                    pass 
                
                new_device_info = dashboard.devices.getDevice(new_serial)
                tags = new_device_info.get('tags', [])
                new_device_name = new_device_info.get('name', [])

                if loc_tag and loc_tag not in tags: tags.append(loc_tag)
                #if new_ap_tag and new_ap_tag not in tags: tags.append(new_ap_tag)
                    
                dashboard.devices.updateDevice(serial=new_serial, name=constructed_name, tags=tags)
                msg = f"ADDED: {new_serial} (Renamed '{new_device_name}' -> '{constructed_name}'), Tags: {tags}"
                logging.info(msg)
                audit_actions.append(msg)

                fp_review_list.append({
                    "serial": new_serial,
                    "name": constructed_name,
                    "reason": "New AP (No Floor Plan assigned)"
                })

            if delay > 0: time.sleep(delay)

        except Exception as e:
            err_msg = f"ERROR processing {decision} on row {index}: {e}"
            logging.error(err_msg)
            audit_actions.append(err_msg)
            raise e 

    # --- BULK REMOVAL ---
    if serials_to_remove:
        logging.info(f"Processing removal of {len(serials_to_remove)} devices...")
        for serial in serials_to_remove:
            try:
                dashboard.networks.removeNetworkDevices(network_id, serial)
                msg = f"CONFIRMED REMOVAL: {serial}"
                logging.info(msg)
                audit_actions.append(msg)
            except Exception as e:
                err_msg = f"FAILED TO REMOVE {serial}: {e}"
                logging.error(err_msg)
                audit_actions.append(err_msg)
                
    return audit_actions, fp_review_list

# --- VALIDATION LOGIC ---

def process_file_validation(file_path, retailer_db, global_inv, dashboard):
    """
    Validates the Excel file content against business rules and Meraki inventory.
    Returns: success (bool), dataframe, metadata
    """
    filename = os.path.basename(file_path)
    logging.info(f"Validating file: {filename}")

    try:
        df_raw = pd.read_excel(file_path, header=None)
        
        # 1. Validate Partner ID
        try:
            partner_id = str(df_raw.iloc[1, 1]).strip()
        except:
            logging.error(f"{filename}: Cannot read Partner ID.")
            return False, None, None

        if partner_id not in retailer_db:
            logging.error(f"{filename}: Partner ID '{partner_id}' invalid.")
            return False, None, None
        
        company_name = retailer_db[partner_id]

        # 2. Locate and Validate Headers
        header_idx = None
        for i, row in df_raw.iterrows():
            if str(row[0]).strip().lower() == "network name":
                header_idx = i
                break
        
        if header_idx is None:
            logging.error(f"{filename}: Header missing.")
            return False, None, None

        found = [str(h).strip().lower() for h in df_raw.iloc[header_idx, 0:9].tolist()] #type: ignore
        if found != EXPECTED_HEADERS_LOWER:
            logging.error(f"{filename}: Header mismatch.")
            return False, None, None

        df = pd.read_excel(file_path, skiprows=header_idx, names=EXPECTED_HEADERS) #type: ignore
        df = df.iloc[:, 0:9].dropna(how='all')
        
        if df.empty:
            logging.error(f"{filename}: Empty data.")
            return False, None, None

        # 3. Row-by-Row Logic Validation
        swap_list_norm = [m.upper() for m in MODELS_TO_SWAP]
        keep_list_norm = [m.upper() for m in MODELS_TO_KEEP]

        for idx, row in df.iterrows():
            d = str(row['Decision']).strip()
            loc = str(row['Location tag']).strip() if pd.notna(row['Location tag']) else ""
            curr_model = str(row['Model']).strip().upper() if pd.notna(row['Model']) else ""
            
            n_mod = str(row['New Model']).strip() if pd.notna(row['New Model']) else ""
            n_sn = str(row['New Serialnumber']).strip() if pd.notna(row['New Serialnumber']) else ""
            n_ap = str(row['New Access Point Name']).strip() if pd.notna(row['New Access Point Name']) else ""
            
            new_info_empty = (n_mod == "" and n_sn == "" and n_ap == "")
            new_info_full = (n_mod != "" and n_sn != "" and "AP" in n_ap)

            # Model Compatibility Check
            if d in ["Replace", "Remove"] and swap_list_norm:
                if curr_model not in swap_list_norm:
                    logging.error(f"{filename} Row {idx}: Model '{curr_model}' not in Swap List.")
                    return False, None, None
            elif d in ["Keep", "Keep/Relocate"] and keep_list_norm:
                if curr_model not in keep_list_norm:
                    logging.error(f"{filename} Row {idx}: Model '{curr_model}' not in Keep List.")
                    return False, None, None

            # Decision Logic Check
            if d == "Replace":
                if loc not in VALID_LOCATIONS:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. Invalid Location tag: '{loc}'. Expected one of: {VALID_LOCATIONS}")
                    return False, None, None
                if not new_info_full:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. Missing required New AP details.")
                    return False, None, None

            elif d == "Remove":
                if loc != "":
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. Location tag must be empty.")
                    return False, None, None
                if not new_info_empty:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. New AP columns must be empty.")
                    return False, None, None

            elif d in ["Keep", "Keep/Relocate"]:
                if loc not in VALID_LOCATIONS:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. Invalid Location tag: '{loc}'. Expected one of: {VALID_LOCATIONS}")
                    return False, None, None
                if not new_info_empty:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. New AP columns must be empty.")
                    return False, None, None

            elif d == "Add":
                if pd.notna(row['Serial number']):
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. The 'Serial number' column (Old AP) must be empty.")
                    return False, None, None
                if loc not in VALID_LOCATIONS:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. Invalid Location tag: '{loc}'. Expected one of: {VALID_LOCATIONS}")
                    return False, None, None
                if not new_info_full:
                    logging.error(f"{filename} Row {idx}: Validation Failed for '{d}'. Missing required New AP details.")
                    return False, None, None

            else:
                logging.error(f"{filename} Row {idx}: Unknown Decision '{d}'. Allowed: {VALID_DECISIONS}")
                return False, None, None

        # --- 4. Identify Target Network & Validate Consistency ---
        
        # Collect ALL serials from the file (Old and New)
        old_serials = df['Serial number'].dropna().astype(str).tolist()
        new_serials = df['New Serialnumber'].dropna().astype(str).tolist()
        
        # Combine and remove duplicates
        all_file_serials = list(set(old_serials + new_serials))

        if not all_file_serials:
            logging.error(f"{filename}: No serial numbers found in file.")
            return False, None, None

        # Pick the first serial to identify the "Anchor" Network
        first_sn = all_file_serials[0]
        
        if first_sn not in global_inv:
            logging.error(f"{filename}: Anchor Serial {first_sn} not found in API Inventory.")
            return False, None, None
            
        target_net_id = global_inv[first_sn].get('networkId')
        target_org_id = global_inv[first_sn].get('orgId')

        if not target_net_id:
            logging.error(f"{filename}: Anchor Serial {first_sn} is currently Unused (not in a network).")
            return False, None, None

        # Validate that ALL serials in the file belong to this SAME network
        for sn in all_file_serials:
            if sn not in global_inv:
                logging.error(f"{filename}: Serial {sn} not found in API Inventory.")
                return False, None, None
            
            sn_net_id = global_inv[sn].get('networkId')
            
            if sn_net_id != target_net_id:
                logging.error(f"{filename}: Network Mismatch! Serial {sn} is in network '{sn_net_id}', but file targets '{target_net_id}'.")
                return False, None, None

        # Validate Sequence (AP01, AP02...)
        seq_ok, seq_msg = validate_final_ap_sequence(df)
        if not seq_ok:
            logging.error(f"{filename}: {seq_msg}")
            return False, None, None
        
        # 5. Check for Orphaned APs (APs in Network but missing from File)
        try:
            # Updated to use Organization endpoint with filtering
            net_devices = dashboard.organizations.getOrganizationDevices(
                target_org_id, 
                total_pages='all',
                networkIds=[target_net_id],
                productTypes=['wireless']
            )
            
            network_ap_serials = set(d['serial'] for d in net_devices)

        except Exception as e:
            logging.error(f"{filename}: Failed to fetch network devices for count check: {e}")
            return False, None, None

        excel_old_serials = set(df['Serial number'].dropna().astype(str).tolist())
        excel_new_serials = set(df['New Serialnumber'].dropna().astype(str).tolist())
        all_excel_serials = excel_old_serials.union(excel_new_serials)

        missing_in_file = network_ap_serials - all_excel_serials
        
        if missing_in_file:
            logging.error(f"{filename}: VALIDATION FAILED. The file is missing {len(missing_in_file)} APs that are currently in the network.")
            logging.error(f"Missing Serials (Orphans): {missing_in_file}")
            return False, None, None

       # 6. Verify Network Consistency
        # Note: Primary validation was done in Section 4. 
        # We assume 'all_excel_serials' from Section 5 contains valid inventory items.
        for sn in all_excel_serials:
            # Double check against global inventory to prevent key errors
            if sn not in global_inv:
                logging.error(f"{filename}: Serial {sn} found in file but not in Global Inventory.")
                return False, None, None
            
            # Ensure no cross-network pollution
            if global_inv[sn].get('networkId') != target_net_id:
                logging.error(f"{filename}: Network mismatch on serial {sn}. Expected {target_net_id}.")
                return False, None, None

        # 7. Verify Device Statuses
        # Use the set from Section 5 to fetch statuses for Old AND New devices in one call
        status_map = get_device_statuses(dashboard, target_org_id, list(all_excel_serials))

        for idx, row in df.iterrows():
            d = row['Decision']
            old_sn = str(row['Serial number']) if pd.notna(row['Serial number']) else None
            new_sn = str(row['New Serialnumber']) if pd.notna(row['New Serialnumber']) else None

            if new_sn:
                if new_sn not in global_inv:
                    logging.error(f"{filename}: New Serial {new_sn} not in inventory.")
                    return False, None, None
                tags = global_inv[new_sn].get('tags', [])
                if NEW_AP_TAG not in tags:
                    logging.error(f"{filename}: New Serial {new_sn} missing {NEW_AP_TAG}.")
                    return False, None, None
                st = status_map.get(new_sn, 'unknown')
                if st in ["online"]:
                        pass # This is the ideal state
                elif st in ["alerting"]:
                    logging.warning(f"{filename} Row {idx}: WARNING - Serial {new_sn} device status is '{st}'")
                else:
                    # Fail if status is 'unknown' or anything else
                    logging.error(f"{filename}: New Serial {new_sn} status {st} invalid.")
                    return False, None, None

            if old_sn:
                st = status_map.get(old_sn, 'unknown')
                if d in ["Replace", "Remove"] and st not in ["offline", "dormant"]:
                    logging.error(f"{filename}: Serial {old_sn} status {st} invalid for {d}.")
                    return False, None, None
                elif d in ["Keep", "Keep/Relocate"]:
                    if st in ["online"]:
                        pass # This is the ideal state
                    elif st in ["offline", "dormant","alerting"]:
                        logging.warning(f"{filename} Row {idx}: WARNING - Decision is '{d}' for Serial {old_sn}, but device status is '{st}'")
                    else:
                        # Fail if status is 'unknown' or anything else
                        logging.error(f"{filename}: Serial {old_sn} status {st} invalid for {d}.")
                        return False, None, None

        # Get Network Metadata
        try:
            current_net = dashboard.networks.getNetwork(target_net_id)
            current_net_name = current_net['name']
            parts = current_net_name.split('-')
            country = parts[0] if len(parts) > 0 else "XX"
            region = parts[1] if len(parts) > 1 else "XXX"
        except:
            current_net_name = "Unknown"
            country = "XX"
            region = "XXX"

        return True, df, {
            'partner_id': partner_id,
            'company_name': company_name,
            'network_id': target_net_id,
            'old_network_name': current_net_name,
            'country': country,
            'region': region
        }

    except Exception as e:
        logging.error(f"CRITICAL ERROR on {filename}: {e}")
        return False, None, None

# --- MAIN EXECUTION ---

def main():
    setup_logging()
    
    # --- STEP 0: PRE-FLIGHT CHECKS ---
    if not check_for_duplicate_filenames_recursive(ROOT_FOLDER):
        logging.critical("Execution stopped due to duplicate filenames found in the Root directory tree.")
        print("\n" + "!"*60)
        print("CRITICAL ERROR: Duplicate filenames detected.")
        print("Check the log file for specific locations.")
        print("!"*60 + "\n")
        sys.exit(1)

    dashboard = meraki.DashboardAPI(
        API_KEY,
        suppress_logging=True,
        print_console=False,
        wait_on_rate_limit=True,
        maximum_retries=5,
        nginx_429_retry_wait_time=60,
        retry_4xx_error=False,
        single_request_timeout=60
    )
    
    show_user_organizations(dashboard)
    
    retailer_db = load_retailer_db()
    global_inv = get_global_inventory(dashboard)
    
    passed_files = []
    failed_count = 0
    
    # --- STEP 1: VALIDATION PHASE ---
    for folder in PATHS_TO_PROCESS:
        if not os.path.exists(folder):
            logging.warning(f"Folder missing: {folder}")
            continue
            
        for f in os.listdir(folder):
            if f.endswith((".xlsx", ".xls")):
                f_path = os.path.join(folder, f)
                success, df, meta = process_file_validation(f_path, retailer_db, global_inv, dashboard)
                
                if success:
                    passed_files.append({'path': f_path, 'df': df, 'meta': meta})
                else:
                    failed_count += 1
                    move_processed_file(f_path, None, False)

    print("\n" + "="*40)
    print(f"VALIDATION SUMMARY")
    print(f"Passed: {len(passed_files)}")
    print(f"Failed: {failed_count}")
    print("="*40 + "\n")

    # Halt execution if any files failed validation
    if failed_count > 0:
        print(f"ATTENTION: {failed_count} file(s) failed validation.")
        print("These files have been moved to the 'Manual handling' folder.")
        print("All files that passed validation are still in the 'Excel to be processed' folder.")
        print("-" * 60)
        print("ACTION REQUIRED: Please resolve the issues with the failed files (if needed)")
        print("and run the script again to process the remaining valid files.")
        print("-" * 60 + "\n")
        return 

    if not passed_files:
        logging.info("No files found to process. Exiting.")
        return

    while True:
        ans = input("All files passed validation. Proceed with Network Changes? (yes/no): ").lower().strip()
        if ans == "no":
            sys.exit(0)
        elif ans == "yes":
            break

    # --- STEP 2: MIGRATION PHASE ---
    logging.info("STARTING MIGRATION...")
    
    for item in passed_files:
        path = item['path']
        df = item['df']
        meta = item['meta']
        filename = os.path.basename(path)
        
        try:
            raw_name = f"{meta['country']}-{meta['region']}-{meta['partner_id']}-{meta['company_name']}"
            final_name = sanitize_network_name(raw_name)
            ap_naming_prefix = f"{meta['country']}-{meta['region']}-{meta['partner_id']}"
            
            logging.info(f"Processing {filename}")
            logging.info(f"Renaming Network {meta['network_id']} -> {final_name}")
            
            dashboard.networks.updateNetwork(meta['network_id'], name=final_name)
            
            actions_log, review_items = migrate_meraki_aps_internal(
                dashboard=dashboard,
                network_id=meta['network_id'],
                ap_df=df,
                naming_prefix=ap_naming_prefix,
                delay=0,
                new_ap_tag=NEW_AP_TAG
            )
            
            review_text = ""
            if review_items:
                review_text = "\n\n*** FLOOR PLAN REVIEW REQUIRED ***\n"
                review_text += f"{'SERIAL':<16} | {'NAME':<30} | {'REASON'}\n"
                review_text += "-"*80 + "\n"
                for item in review_items:
                    review_text += f"{item['serial']:<16} | {item['name']:<30} | {item['reason']}\n"
                logging.info(review_text)

            audit_entry = (
                f"--------------------------------------------------\n"
                f"DATE: {datetime.now()}\n"
                f"FILE: {filename}\n"
                f"NETWORK ID: {meta['network_id']}\n"
                f"NAME CHANGE: '{meta['old_network_name']}' -> '{final_name}'\n"
                f"PARTNER: {meta['partner_id']} ({meta['company_name']})\n"
                f"ACTIONS:\n" + "\n".join([f"  - {a}" for a in actions_log])
            )

            if review_text:
                audit_entry += review_text

            audit_entry += "\n--------------------------------------------------"
            append_audit_log(audit_entry)
            
            move_processed_file(path, df, True,review_items)
            logging.info(f"Successfully processed {filename}")
            
        except Exception as e:
            error_msg = f"CRITICAL ERROR processing {filename}: {e}"
            logging.error(error_msg)
            append_audit_log(f"FAILED PROCESSING {filename}: {e}")
            move_processed_file(path, df, False)
            
            print("\n" + "!"*60)
            print(f"⚠️  EXECUTION PAUSED due to error in file: {filename}")
            print(f"Error Details: {e}")
            print("The file has been moved to 'Manual handling'.")
            print("!"*60 + "\n")
            
            while True:
                user_choice = input("Do you want to continue with the next file? (yes/no): ").lower().strip()
                if user_choice in ['yes', 'y']:
                    break 
                elif user_choice in ['no', 'n']:
                    sys.exit(1)
                else:
                    print("Please answer 'yes' or 'no'.")

if __name__ == "__main__":
    main()