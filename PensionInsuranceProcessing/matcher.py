import os
import openpyxl
from difflib import SequenceMatcher
from cleaner import clean_pdf_company_name, normalize_for_matching

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def _find_file(filename, fallback_rel=None):
    if os.path.exists(filename):
        return filename
    in_script = os.path.join(SCRIPT_DIR, filename)
    if os.path.exists(in_script):
        return in_script
    if fallback_rel:
        in_fallback = os.path.join(os.path.dirname(SCRIPT_DIR), fallback_rel)
        if os.path.exists(in_fallback):
            return in_fallback
    return filename

class CompanyMatcher:
    def __init__(self, 
                 instrument_master_path="MAIN_Instrument_Master_october.xlsx",
                 company_master_path="company_master_oct.xlsx",
                 mapper_file_path="company_mapping_master.xlsx",
                 fuzzy_threshold=0.88):
        self.instrument_master_path = _find_file(instrument_master_path, "MAIN_Instrument_Master_october.xlsx")
        self.company_master_path = _find_file(company_master_path, r"october\company_master_oct.xlsx")
        self.mapper_file_path = _find_file(mapper_file_path, "company_mapping_master.xlsx")
        self.fuzzy_threshold = fuzzy_threshold
        
        self.mapper_approved = {}   # key -> dict
        self.mapper_review = {}     # key -> dict
        self.mapper_pending = set() # keys already in pending
        self.mapper_rejected = set()
        
        self.im_exact = {}
        self.im_case = {}
        self.im_norm = {}
        self.im_keys_list = []
        
        self.cm_exact = {}
        self.cm_case = {}
        self.cm_norm = {}
        self.cm_keys_list = []
        
        self.new_pending_records = []
        
        self._load_mapper()
        self._load_instrument_master()
        self._load_company_master()

    def _load_mapper(self):
        if not os.path.exists(self.mapper_file_path):
            return
        wb = openpyxl.load_workbook(self.mapper_file_path, data_only=True)
        ws = wb.active
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or not any(row):
                continue
            raw_name = str(row[0] or "").strip()
            cleaned_name = str(row[1] or "").strip()
            mapped_db = str(row[2] or "").strip()
            inst = str(row[3] or "").strip() or "Equity"
            nature = str(row[4] or "").strip() or "EQ"
            status = str(row[5] or "").strip().capitalize()
            source = str(row[6] or "").strip()
            level = str(row[7] or "").strip()
            comments = str(row[8] or "").strip()
            
            entry = {
                "db_name": mapped_db,
                "instrument": inst,
                "nature": nature,
                "status": status,
                "source": source,
                "approximation_level": level or f"Mapper ({status})",
                "comments": comments
            }
            
            if status == "Approved":
                if raw_name:
                    self.mapper_approved[raw_name] = entry
                    self.mapper_approved[normalize_for_matching(raw_name)] = entry
                if cleaned_name:
                    self.mapper_approved[cleaned_name] = entry
                    self.mapper_approved[normalize_for_matching(cleaned_name)] = entry
            elif status == "Review":
                if raw_name:
                    self.mapper_review[raw_name] = entry
                    self.mapper_review[normalize_for_matching(raw_name)] = entry
                if cleaned_name:
                    self.mapper_review[cleaned_name] = entry
                    self.mapper_review[normalize_for_matching(cleaned_name)] = entry
            elif status == "Pending":
                if raw_name: self.mapper_pending.add(raw_name)
                if cleaned_name: self.mapper_pending.add(cleaned_name)
            elif status == "Rejected":
                if raw_name: self.mapper_rejected.add(raw_name)
                if cleaned_name: self.mapper_rejected.add(cleaned_name)

    def _load_instrument_master(self):
        if not os.path.exists(self.instrument_master_path):
            print(f"Warning: Instrument master {self.instrument_master_path} not found.")
            return
        wb = openpyxl.load_workbook(self.instrument_master_path, data_only=True)
        ws = wb["Final"] if "Final" in wb.sheetnames else wb.active
        
        for row in list(ws.iter_rows(values_only=True))[1:]:
            search_key = str(row[1]).strip() if row[1] is not None else ""
            db_name = str(row[2]).strip() if row[2] is not None else ""
            inst = str(row[4]).strip() if row[4] is not None else ""
            nature = str(row[5]).strip() if row[5] is not None else ""
            
            if not search_key or not db_name:
                continue
                
            # Follow simple lookup rule: first occurrence gets stored
            if search_key not in self.im_exact:
                self.im_exact[search_key] = (db_name, inst, nature)
            
            c_k = search_key.lower()
            if c_k not in self.im_case:
                self.im_case[c_k] = (db_name, inst, nature)
                
            n_k = normalize_for_matching(search_key)
            if n_k and n_k not in self.im_norm:
                self.im_norm[n_k] = (db_name, inst, nature)
                self.im_keys_list.append((n_k, (db_name, inst, nature)))

    def _load_company_master(self):
        if not os.path.exists(self.company_master_path):
            print(f"Warning: Company master {self.company_master_path} not found.")
            return
        wb = openpyxl.load_workbook(self.company_master_path, data_only=True)
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
        
        for row in list(ws.iter_rows(values_only=True))[1:]:
            com_name = str(row[0]).strip() if row[0] is not None else ""
            if not com_name:
                continue
            
            # Rule: Company master matches are always Equity and EQ
            target = (com_name, "Equity", "EQ")
            if com_name not in self.cm_exact:
                self.cm_exact[com_name] = target
                
            c_k = com_name.lower()
            if c_k not in self.cm_case:
                self.cm_case[c_k] = target
                
            n_k = normalize_for_matching(com_name)
            if n_k and n_k not in self.cm_norm:
                self.cm_norm[n_k] = target
                self.cm_keys_list.append((n_k, target))

    def match(self, raw_name: str, clean_name: str) -> dict:
        """
        Executes hierarchical matching:
        0. Mapper File (Approved / Review)
        1. Exact Match in Instrument Master
        2. Exact Match in Company Master
        3. Approximation Match in Instrument Master (lowercase, normalized, fuzzy)
        4. Approximation Match in Company Master (lowercase, normalized, fuzzy)
        5. Fallback / Unmatched (saved for Pending review)
        """
        raw_name_str = (raw_name or "").strip()
        clean_name_str = (clean_name or "").strip()
        norm_name_str = normalize_for_matching(clean_name_str)
        
        # 0. Check Mapper File: Approved
        for key in [raw_name_str, clean_name_str, norm_name_str]:
            if key and key in self.mapper_approved:
                entry = self.mapper_approved[key]
                return {
                    "db_name": entry["db_name"],
                    "instrument": entry["instrument"],
                    "nature": entry["nature"],
                    "approximation_level": f"Mapper Approved: {entry['approximation_level']}",
                    "status": "Approved",
                    "source": "Mapper Master"
                }
                
        # 0. Check Mapper File: Review
        for key in [raw_name_str, clean_name_str, norm_name_str]:
            if key and key in self.mapper_review:
                entry = self.mapper_review[key]
                return {
                    "db_name": entry["db_name"],
                    "instrument": entry["instrument"],
                    "nature": entry["nature"],
                    "approximation_level": f"Mapper Review: {entry['approximation_level']}",
                    "status": "Review",
                    "source": "Mapper Master"
                }

        # 1. Exact Match in Instrument Master
        if clean_name_str and clean_name_str in self.im_exact:
            db_name, inst, nature = self.im_exact[clean_name_str]
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": "Exact Match (Instrument Master)",
                "status": "Approved",
                "source": "Instrument Master"
            }

        # 2. Exact Match in Company Master
        if clean_name_str and clean_name_str in self.cm_exact:
            db_name, inst, nature = self.cm_exact[clean_name_str]
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": "Exact Match (Company Master)",
                "status": "Approved",
                "source": "Company Master"
            }

        # 3. Approximation Match in Instrument Master
        # 3a. Case-insensitive
        clean_lower = clean_name_str.lower()
        if clean_lower and clean_lower in self.im_case:
            db_name, inst, nature = self.im_case[clean_lower]
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": "Converted to lowercase (Instrument Master)",
                "status": "Approved",
                "source": "Instrument Master"
            }

        # 3b. Normalized Suffix / Punctuation / OCR
        if norm_name_str and norm_name_str in self.im_norm:
            db_name, inst, nature = self.im_norm[norm_name_str]
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": "Normalized suffix/punctuation/OCR (Instrument Master)",
                "status": "Approved",
                "source": "Instrument Master"
            }

        # 4. Approximation Match in Company Master
        # 4a. Case-insensitive
        if clean_lower and clean_lower in self.cm_case:
            db_name, inst, nature = self.cm_case[clean_lower]
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": "Converted to lowercase (Company Master)",
                "status": "Approved",
                "source": "Company Master"
            }

        # 4b. Normalized Suffix / Punctuation / OCR
        if norm_name_str and norm_name_str in self.cm_norm:
            db_name, inst, nature = self.cm_norm[norm_name_str]
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": "Normalized suffix/punctuation/OCR (Company Master)",
                "status": "Approved",
                "source": "Company Master"
            }

        # 3c / 4c. High-confidence Fuzzy Matching
        # First check Instrument Master fuzzy
        best_im = None
        best_im_score = 0
        if norm_name_str:
            for k_norm, val in self.im_keys_list:
                score = SequenceMatcher(None, norm_name_str, k_norm).ratio()
                if score > best_im_score:
                    best_im_score = score
                    best_im = val
                    
        if best_im_score >= self.fuzzy_threshold and best_im:
            db_name, inst, nature = best_im
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": f"Fuzzy token match ({best_im_score:.0%}) (Instrument Master)",
                "status": "Review",
                "source": "Instrument Master"
            }

        # Next check Company Master fuzzy
        best_cm = None
        best_cm_score = 0
        if norm_name_str:
            for k_norm, val in self.cm_keys_list:
                score = SequenceMatcher(None, norm_name_str, k_norm).ratio()
                if score > best_cm_score:
                    best_cm_score = score
                    best_cm = val
                    
        if best_cm_score >= self.fuzzy_threshold and best_cm:
            db_name, inst, nature = best_cm
            return {
                "db_name": db_name,
                "instrument": inst,
                "nature": nature,
                "approximation_level": f"Fuzzy token match ({best_cm_score:.0%}) (Company Master)",
                "status": "Review",
                "source": "Company Master"
            }

        # 5. Unmatched
        self._record_pending(raw_name_str, clean_name_str)
        return {
            "db_name": "",
            "instrument": "",
            "nature": "",
            "approximation_level": "Unmatched - Requires Manual Review",
            "status": "Pending",
            "source": "None"
        }

    def _record_pending(self, raw_name: str, clean_name: str):
        if not raw_name and not clean_name:
            return
        if raw_name in self.mapper_pending or clean_name in self.mapper_pending:
            return
        self.new_pending_records.append({
            "raw_name": raw_name,
            "cleaned_name": clean_name,
            "mapped_db": "",
            "instrument": "",
            "nature": "",
            "status": "Pending",
            "source": "Pipeline Unmatched",
            "approximation_level": "None",
            "comments": "Automated log: Needs manual review"
        })
        if raw_name: self.mapper_pending.add(raw_name)
        if clean_name: self.mapper_pending.add(clean_name)

    def save_pending_to_mapper(self):
        """Appends new pending rows to company_mapping_master.xlsx without overwriting existing rows."""
        if not self.new_pending_records:
            return
        if not os.path.exists(self.mapper_file_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Company_Mappings"
            ws.append([
                'RAW_NAME', 'CLEANED_NAME', 'MAPPED_DATABASE_NAME', 'INSTRUMENT_NAME',
                'NATURE', 'STATUS', 'MATCH_SOURCE', 'APPROXIMATION_LEVEL', 'COMMENTS'
            ])
        else:
            wb = openpyxl.load_workbook(self.mapper_file_path)
            ws = wb.active

        for rec in self.new_pending_records:
            ws.append([
                rec["raw_name"],
                rec["cleaned_name"],
                rec["mapped_db"],
                rec["instrument"],
                rec["nature"],
                rec["status"],
                rec["source"],
                rec["approximation_level"],
                rec["comments"]
            ])
        wb.save(self.mapper_file_path)
        print(f"Appended {len(self.new_pending_records)} pending review entries to {self.mapper_file_path}")
        self.new_pending_records = []
