import pandas as pd
import re
from fuzzywuzzy import fuzz, process
from collections import defaultdict

class alltrans:
    def __init__(self):
        self.standard_columns = [
            'Driver Name', 'License Number', 'Date of Birth', 'License State',
            'Expiration Date', 'Status', 'Hire Date', 'Comment'
        ]

    def clean_column_names(self, df):
        df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_', regex=False)
        return df

    def normalize_name(self, name):
        if pd.isna(name) or not name:
            return []
        name = str(name).lower()
        name = re.sub(r'\b(mr|mrs|ms|dr|jr|sr|iii|ii|iv)\b', '', name, flags=re.IGNORECASE)
        name = re.sub(r'[^a-z\s]', '', name)
        name = re.sub(r'\s+', ' ', name).strip()
        parts = name.split()
        if not parts:
            return []

        formats = [' '.join(parts)]
        if len(parts) > 1:
            formats += [
                f"{parts[0]} {parts[-1]}", f"{parts[-1]} {parts[0]}",
                f"{parts[0]}{parts[-1]}", f"{parts[-1]}{parts[0]}"
            ]
        if len(parts) > 2:
            first, last = parts[0], parts[-1]
            initials = ''.join(p[0] for p in parts[1:-1])
            formats += [
                f"{first} {initials} {last}", f"{first} {initials}{last}",
                f"{first}{initials} {last}", f"{first}{initials}{last}"
            ]
        return list(set(formats))

    def names_match(self, name1, name2):
        if pd.isna(name1) or pd.isna(name2) or not name1 or not name2:
            return False
        formats1 = self.normalize_name(name1)
        formats2 = self.normalize_name(name2)
        for f1 in formats1:
            for f2 in formats2:
                if f1 == f2:
                    return True
                if fuzz.token_set_ratio(f1, f2) >= 95:
                    return True
                if fuzz.partial_ratio(f1, f2) >= 96:
                    return True
                if fuzz.token_sort_ratio(f1, f2) >= 98:
                    return True
        return False

    def get_valid_column(self, df, purpose, default_names, required=True):
        for col in default_names:
            if col in df.columns:
                return col
        for col_name in default_names:
            match, score = process.extractOne(col_name, df.columns, scorer=fuzz.ratio)
            if score > 80:
                return match
        if required and len(df.columns) > 0:
            return df.columns[0]
        return None

    def extract_drivers_from_mvr(self, mvr_df):
        """Extract drivers from MVR data with specific column mapping"""
        mvr_df = self.clean_column_names(mvr_df)
        drivers = []

        # Map MVR columns to our standard format
        name_mapping = {
            'driver_full_name': 'Driver Name',
            'name': 'Driver Name',
            'cdl_number': 'License Number', 
            'license_number': 'License Number',
            'driver_date_of_birth': 'Date of Birth',
            'date_of_birth': 'Date of Birth',
            'dob': 'Date of Birth',
            'license_state': 'License State',
            'state': 'License State'
        }
        
        # Find the actual column names in the MVR data
        actual_mapping = {}
        for std_col in ['Driver Name', 'License Number', 'Date of Birth', 'License State']:
            possible_names = [k for k, v in name_mapping.items() if v == std_col]
            for possible_name in possible_names:
                for actual_col in mvr_df.columns:
                    if possible_name in actual_col.lower():
                        actual_mapping[std_col] = actual_col
                        break
                if std_col in actual_mapping:
                    break
        
        # Extract driver information with duplicate removal
        if 'Driver Name' in actual_mapping:
            name_col = actual_mapping['Driver Name']
            seen_licenses = set()
            
            for _, row in mvr_df.iterrows():
                license_num = str(row[actual_mapping['License Number']]).strip() if 'License Number' in actual_mapping else ''
                
                # Skip duplicates
                if license_num and license_num in seen_licenses:
                    continue
                if license_num:
                    seen_licenses.add(license_num)
                
                driver_data = {
                    'Driver Name': row[name_col] if name_col in row else '',
                    'License Number': license_num,
                    'Date of Birth': row[actual_mapping['Date of Birth']] if 'Date of Birth' in actual_mapping else '',
                    'License State': row[actual_mapping['License State']] if 'License State' in actual_mapping else '',
                    'Expiration Date': '',
                    'Status': '',
                    'Hire Date': '',
                    'Comment': 'EXTRACTED FROM MVR'
                }
                drivers.append(driver_data)

        return pd.DataFrame(drivers)

    def normalize_license(self, license_str):
        """Normalize license number - handle spaces but keep exact match possibility"""
        if pd.isna(license_str) or not license_str:
            return None
        license_str = str(license_str).strip()
        if license_str in ('', 'nan', 'none', 'NaN', 'None'):
            return None
        # Remove extra spaces but keep the basic format
        license_str = re.sub(r'\s+', ' ', license_str).strip()
        return license_str.upper()

    def match_drivers(self, client_df, mvr_df, hire_date_col, dob_col, license_col):
        """Optimized matching function that preserves your existing variable names"""
        results = []
        matched_client_indices = set()
        matched_mvr_indices = set()

        # Clean both datasets
        client_df_clean = self.clean_column_names(client_df.copy())
        
        # Get client column names
        client_name_col = self.get_valid_column(client_df_clean, "driver names", ['name', 'driver_name', 'full_name'])
        license_num_col = self.get_valid_column(client_df_clean, "license number", ['license_number', 'cdl_number', 'lic_number'], False)
        
        # Normalize parameter column names
        hire_date_col_clean = hire_date_col.strip().lower().replace(' ', '_') if hire_date_col else None
        dob_col_clean = dob_col.strip().lower().replace(' ', '_') if dob_col else None
        license_col_clean = license_col.strip().lower().replace(' ', '_') if license_col else None

        # Build license number hash map from MVR data for O(1) lookups
        mvr_license_map = {}
        for mvr_idx, mvr_row in mvr_df.iterrows():
            license_num = self.normalize_license(mvr_row.get('License Number', ''))
            if license_num and license_num not in mvr_license_map:
                mvr_license_map[license_num] = (mvr_idx, mvr_row)

        # STAGE 1: LICENSE NUMBER MATCHING (Primary)
        if license_num_col and license_num_col in client_df_clean.columns:
            for client_idx, client_row in client_df_clean.iterrows():
                if client_idx in matched_client_indices:
                    continue
                    
                client_license = self.normalize_license(client_row.get(license_num_col))
                if not client_license:
                    continue
                
                # O(1) lookup in hash map
                if client_license in mvr_license_map:
                    mvr_idx, mvr_row = mvr_license_map[client_license]
                    
                    if mvr_idx not in matched_mvr_indices:
                        result = {
                            'Driver Name': client_row[client_name_col],
                            'License Number': client_row.get(license_num_col, ''),
                            'Date of Birth': client_row[dob_col_clean] if dob_col_clean and dob_col_clean in client_row else '',
                            'License State': client_row[license_col_clean] if license_col_clean and license_col_clean in client_row else '',
                            'Expiration Date': mvr_row.get('Expiration Date', ''),
                            'Status': mvr_row.get('Status', ''),
                            'Hire Date': client_row[hire_date_col_clean] if hire_date_col_clean and hire_date_col_clean in client_row else '',
                            'Comment': 'MATCH FOUND (License)'
                        }
                        results.append(result)
                        matched_client_indices.add(client_idx)
                        matched_mvr_indices.add(mvr_idx)

        # STAGE 2: NAME + DOB MATCHING (Fallback)
        for client_idx, client_row in client_df_clean.iterrows():
            if client_idx in matched_client_indices:
                continue
                
            client_name = client_row[client_name_col]
            client_dob = client_row[dob_col_clean] if dob_col_clean and dob_col_clean in client_row else None
            
            for mvr_idx, mvr_row in mvr_df.iterrows():
                if mvr_idx in matched_mvr_indices:
                    continue
                    
                mvr_name = mvr_row['Driver Name']
                mvr_dob = mvr_row['Date of Birth']
                
                # Check name match
                if self.names_match(client_name, mvr_name):
                    # If DOB is available, check it too
                    dob_match = True
                    if client_dob and mvr_dob:
                        try:
                            client_dob_str = str(client_dob).split()[0]
                            mvr_dob_str = str(mvr_dob).split()[0]
                            dob_match = (client_dob_str == mvr_dob_str)
                        except:
                            dob_match = False
                    
                    if dob_match:
                        result = {
                            'Driver Name': client_name,
                            'License Number': mvr_row['License Number'],
                            'Date of Birth': client_dob if client_dob else mvr_dob,
                            'License State': client_row[license_col_clean] if license_col_clean and license_col_clean in client_row else mvr_row['License State'],
                            'Expiration Date': mvr_row.get('Expiration Date', ''),
                            'Status': mvr_row.get('Status', ''),
                            'Hire Date': client_row[hire_date_col_clean] if hire_date_col_clean and hire_date_col_clean in client_row else '',
                            'Comment': 'MATCH FOUND (Name+DOB)'
                        }
                        results.append(result)
                        matched_client_indices.add(client_idx)
                        matched_mvr_indices.add(mvr_idx)
                        break

        # Add unmatched client drivers as "MISSING MVR"
        for client_idx, client_row in client_df_clean.iterrows():
            if client_idx not in matched_client_indices:
                result = {
                    'Driver Name': client_row[client_name_col],
                    'License Number': client_row.get(license_num_col, '') if license_num_col else '',
                    'Date of Birth': client_row[dob_col_clean] if dob_col_clean and dob_col_clean in client_row else '',
                    'License State': client_row[license_col_clean] if license_col_clean and license_col_clean in client_row else '',
                    'Expiration Date': '',
                    'Status': '',
                    'Hire Date': client_row[hire_date_col_clean] if hire_date_col_clean and hire_date_col_clean in client_row else '',
                    'Comment': 'MISSING MVR'
                }
                results.append(result)

        # Add unmatched MVR drivers as "Extra MVR records"
        for mvr_idx, mvr_row in mvr_df.iterrows():
            if mvr_idx not in matched_mvr_indices:
                result = {
                    'Driver Name': mvr_row['Driver Name'],
                    'License Number': mvr_row['License Number'],
                    'Date of Birth': mvr_row['Date of Birth'],
                    'License State': mvr_row['License State'],
                    'Expiration Date': mvr_row.get('Expiration Date', ''),
                    'Status': mvr_row.get('Status', ''),
                    'Hire Date': '',
                    'Comment': 'Extra MVR record (no client match)'
                }
                results.append(result)

        return pd.DataFrame(results)