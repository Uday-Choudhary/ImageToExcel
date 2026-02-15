"""
Spatial Table Extractor
Analyzes OCR data to reconstruct tables based on spatial alignment.
"""

import json
import re
import numpy as np
from collections import defaultdict


class SpatialTableExtractor:
    def __init__(self):
        self.min_confidence = 0.3  # Filter low-confidence noise

    # ─── Text Cleaning ───────────────────────────────────────────────

    def clean_text(self, text):
        """
        Fixes common OCR errors using regex.
        - S/s -> $ (start of number)
        - O/o -> 0 (numeric context)
        - i/I/l -> 1 (numeric context)
        - B -> 8 (numeric context)
        """
        if not text:
            return ""

        text = re.sub(r'^[Ss](?=\d)', '$', text)                          # S100 → $100
        text = re.sub(r'(?<=[\d$.,])[Oo]|[Oo](?=[\d$.,])', '0', text)    # O → 0
        text = re.sub(r'(?<=[\d$.,])[iIl]|[iIl](?=[\d$.,])', '1', text)  # i/I/l → 1
        text = re.sub(r'(?<=[\d$.,])B|B(?=[\d$.,])', '8', text)          # B → 8
        text = re.sub(r'(?<=[\d$.,])\||\|(?=[\d$.,])', '1', text)        # | → 1
        return text.strip('_').strip()

    # ─── Geometry Helpers ─────────────────────────────────────────────

    def _compute_overlap(self, w_min, w_max, b_start, b_end):
        """
        Compute the overlap ratio between a word's X range [w_min, w_max]
        and a column boundary [b_start, b_end].
        Returns the fraction of the word that falls within the column.
        """
        overlap_start = max(w_min, b_start)
        overlap_end = min(w_max, b_end)
        overlap = max(0, overlap_end - overlap_start)
        word_width = max(w_max - w_min, 1)  # avoid division by zero
        return overlap / word_width

    @staticmethod
    def _vertical_overlap(a_min, a_max, b_min, b_max):
        """Overlap ratio of two vertical ranges relative to the smaller."""
        overlap = max(0, min(a_max, b_max) - max(a_min, b_min))
        smaller = min(a_max - a_min, b_max - b_min)
        return overlap / smaller if smaller > 0 else 0

    # ─── Row Clustering ──────────────────────────────────────────────

    def _cluster_rows(self, words):
        """
        Group words into rows using vertical overlap (>=35%) with
        center-distance fallback.
        """
        if not words:
            return []

        words.sort(key=lambda w: w['y_center'])
        rows = []
        current = [words[0]]

        for word in words[1:]:
            in_row = any(
                self._vertical_overlap(m['y_min'], m['y_max'],
                                       word['y_min'], word['y_max']) >= 0.35
                for m in current
            )
            if not in_row:
                median_y = np.median([w['y_center'] for w in current])
                avg_h = np.mean([w['height'] for w in current])
                in_row = abs(word['y_center'] - median_y) < avg_h * 0.5

            if in_row:
                current.append(word)
            else:
                current.sort(key=lambda w: w['x_min'])
                rows.append(current)
                current = [word]

        if current:
            current.sort(key=lambda w: w['x_min'])
            rows.append(current)

        return rows

    # ─── Header Detection ────────────────────────────────────────────

    _PRIMARY_KW = [
        # Core transaction terms
        "description", "particulars", "desc", "qty", "quantity", "gty", "qnty", "unit", "price", "rate", "amount", "amcunt", "amt",
        "gross", "net", "vat", "sku", "item", "um",
        # Product/Inventory
        "product", "article", "code", "barcode", "upc", "part", "partno",
        "catalog", "model", "reference", "ref", "serial", "batch",
        # Financial
        "cost", "value", "discount", "subtotal", "balance", "debit",
        "credit", "payment", "charge", "fee", "expense", "revenue",
        # Measurements & Units
        "weight", "volume", "length", "width", "height", "size",
        "dimension", "capacity", "measure", "metric",
        # Dates & Time
        "date", "time", "period", "duration", "from", "to", "due",
        "issued", "expiry", "validity", "timestamp",
        # People & Entities
        "name", "customer", "vendor", "supplier", "client", "employee",
        "staff", "contact", "person", "entity", "organization",
        # Identification
        "number", "id", "identifier", "account", "invoice", "order",
        "transaction", "receipt", "voucher", "ticket", "bill",
        # Status & Classification
        "status", "category", "type", "class", "grade", "level",
        "rank", "priority", "stage", "phase", "condition",
        # Location
        "location", "address", "city", "state", "country", "region",
        "zone", "area", "warehouse", "store", "branch",
        # Nutrition (food labels)
        "calories", "protein", "carbs", "fat", "sugar", "sodium",
        "fiber", "vitamin", "mineral", "serving", "nutrition",
        # Medical/Healthcare
        "diagnosis", "medication", "dosage", "patient", "treatment",
        "prescription", "test", "result", "symptom", "procedure",
        "service", "billed", "paid", "outstanding",
        # Education
        "subject", "course", "grade", "score", "marks", "student",
        "teacher", "semester", "term", "credits", "gpa",
        # HR/Payroll
        "salary", "wage", "hours", "overtime", "allowance", "deduction",
        "bonus", "commission", "benefits", "leave", "attendance",
        # Shipping/Logistics
        "tracking", "shipment", "delivery", "freight", "carrier",
        "origin", "destination", "package", "container", "pallet",
        # Manufacturing
        "production", "output", "yield", "defect", "quality",
        "machine", "operator", "shift", "process", "assembly",
    ]

    _SECONDARY_KW = [
        # General qualifiers
        "no.", "total", "tax", "subtotal", "per", "label",
        # Nutrition details
        "calories", "serving", "nutrition", "value", "daily",
        "dv", "percent", "intake", "recommended", "rda",
        # Financial modifiers
        "taxable", "nontaxable", "exempt", "included", "exclusive",
        "inclusive", "before", "after", "adjusted", "unadjusted",
        # Quantities & Counts
        "count", "pieces", "units", "pack", "box", "case",
        "dozen", "pair", "set", "lot", "bundle",
        # Percentages & Ratios
        "percentage", "ratio", "proportion", "share", "margin",
        "markup", "variance", "change", "growth", "decline",
        # Time periods
        "daily", "weekly", "monthly", "quarterly", "yearly",
        "annual", "biannual", "fiscal", "calendar", "ytd",
        # Performance metrics
        "target", "actual", "planned", "forecast", "budget",
        "variance", "achievement", "performance", "kpi", "metric",
        # Quality & Compliance
        "approved", "rejected", "pending", "verified", "certified",
        "compliant", "standard", "specification", "tolerance", "limit",
        # Actions & Operations
        "add", "remove", "update", "modify", "delete",
        "create", "cancel", "approve", "reject", "submit",
        # Comparisons
        "previous", "current", "next", "last", "first",
        "minimum", "maximum", "average", "median", "range",
        # Document types
        "invoice", "quote", "estimate", "proposal", "contract",
        "agreement", "statement", "report", "summary", "detail",
        # Currency & Pricing
        "currency", "exchange", "forex", "conversion", "msrp",
        "retail", "wholesale", "list", "sale", "clearance",
        # Inventory management
        "stock", "available", "reserved", "allocated", "backorder",
        "reorder", "safety", "onhand", "intransit", "committed",
        # Medical specific
        "frequency", "route", "strength", "form", "indication",
        "contraindication", "side", "effect", "interaction", "warning",
        # Project management
        "task", "milestone", "deliverable", "dependency", "resource",
        "effort", "progress", "completion", "baseline", "revision",
        # eCommerce
        "cart", "wishlist", "checkout", "shipping", "handling",
        "coupon", "promo", "refund", "return", "exchange",
        # Accounting
        "journal", "ledger", "trial", "reconciliation", "accrual",
        "depreciation", "amortization", "asset", "liability", "equity",
        # HR specific
        "department", "position", "title", "employment", "probation",
        "appraisal", "increment", "resignation", "termination", "hire",
        # Laboratory/Testing
        "sample", "specimen", "analysis", "measurement", "reading",
        "observation", "finding", "conclusion", "method", "protocol",
        # Retail
        "aisle", "shelf", "bin", "rack", "display",
        "pos", "till", "register", "tender", "change",
    ]

    def _find_header_row(self, rows, max_scan=25):
        """
        Find the best header row by scanning all candidate rows and
        picking the one with the highest keyword score.
        
        Requirements:
          - Row must contain >=2 OCR word detections (avoids titles)
          - Needs >=1 primary keyword + >=2 total matches, OR >=3 secondary
          - Among all qualifying rows, picks the one with most matches
        """
        best_idx, best_row, best_score = -1, [], 0

        for i, row in enumerate(rows[:max_scan]):
            # Skip rows with only 1 OCR detection (likely a title, not headers)
            if len(row) < 2:
                continue

            text = " ".join(w['text'].lower().strip('_').strip() for w in row)
            pri = [k for k in self._PRIMARY_KW if re.search(r'\b' + re.escape(k) + r'\b', text)]
            sec = [k for k in self._SECONDARY_KW if re.search(r'\b' + re.escape(k) + r'\b', text)]
            
            total = len(pri) + len(sec)
            qualifies = (pri and total >= 2) or (len(sec) >= 3)
            
            if qualifies and total > best_score:
                best_score = total
                best_idx = i
                best_row = row

        return best_idx, best_row

    # ─── Split Header Merging ────────────────────────────────────────

    def _merge_split_headers(self, rows, header_row_index):
        """
        Handle headers that span multiple OCR lines (e.g. 'Gross_' on one line
        and 'worth' on the next). Merge the continuation into the header row.
        """
        if header_row_index < 0 or header_row_index >= len(rows) - 1:
            return rows

        header_row = rows[header_row_index]
        next_row = rows[header_row_index + 1]

        next_texts = [w for w in next_row if w['text'].strip()]

        # Check if next row has data (numbers) — if so, don't merge
        for w in next_texts:
            t = w['text'].strip()
            if re.match(r'^[\$Ss]?\d+[.,]?\d*$', t):
                return rows

        header_y_max = max(w['y_max'] for w in header_row)
        next_y_min = min(w['y_min'] for w in next_row)
        avg_height = np.mean([w['height'] for w in header_row])
        y_gap = next_y_min - header_y_max

        if len(next_texts) <= 3 and y_gap < avg_height * 1.0:
            merged_any = False
            avg_width = np.mean([h['x_max'] - h['x_min'] for h in header_row])
            for word in next_texts:
                best = min(header_row, key=lambda h: abs(word['x_center'] - h['x_center']))
                if abs(word['x_center'] - best['x_center']) < avg_width * 1.5:
                    best['text'] = best['text'].rstrip('_') + " " + word['text']
                    merged_any = True

            if merged_any:
                rows = rows[:header_row_index + 1] + rows[header_row_index + 2:]

        return rows

    # ─── Column Boundaries ───────────────────────────────────────────

    @staticmethod
    def _boundaries_from_headers(header_words, all_words):
        """Build column boundaries from header positions."""
        header_words.sort(key=lambda w: w['x_min'])
        lo = min(w['x_min'] for w in all_words)
        hi = max(w['x_max'] for w in all_words)

        bounds = [lo]
        for i in range(1, len(header_words)):
            mid = (header_words[i - 1]['x_max'] + header_words[i]['x_min']) / 2
            bounds.append(mid)
        bounds.append(hi + 20)
        return bounds

    @staticmethod
    def _boundaries_from_gutters(rows, all_words):
        """Detect column boundaries via vertical gap (gutter) analysis."""
        lo = min(w['x_min'] for w in all_words)
        hi = max(w['x_max'] for w in all_words)
        width = int(hi + 20)
        avg_h = np.mean([w['height'] for w in all_words])
        min_gap = max(avg_h * 0.25, 8)

        hist = np.zeros(width, dtype=float)
        for row in rows:
            cov = np.zeros(width, dtype=int)
            for w in row:
                cov[max(0, int(w['x_min'])):min(width, int(w['x_max']))] = 1
            hist += cov

        gaps = []
        in_gap, start = False, 0
        for x in range(int(lo), min(int(hi), width)):
            if hist[x] <= 0:
                if not in_gap:
                    in_gap, start = True, x
            elif in_gap:
                in_gap = False
                if x - start > min_gap:
                    gaps.append((start, x))

        bounds = [lo]
        for gs, ge in gaps:
            bounds.append((gs + ge) / 2)
        bounds.append(hi + 20)
        return bounds

    # ─── Continuation Row Merging ────────────────────────────────────

    def _merge_continuation_rows(self, table_rows, num_columns):
        """
        Merge 'continuation rows' into the previous data row.
        A continuation row has text only in description columns
        and empty data columns.
        """
        if not table_rows or num_columns < 2:
            return table_rows

        # Identify which columns typically hold numeric data
        data_cols = {
            i for i in range(num_columns)
            if any(re.match(r'^[\d$.,% ]+$', row[i].strip().replace(' ', ''))
                   for row in table_rows if row[i].strip())
        }
        if not data_cols:
            data_cols = set(range(2, num_columns))

        merged = []
        for row in table_rows:
            if not merged:
                merged.append(row)
                continue

            has_text = any(row[i].strip() for i in range(num_columns) if i not in data_cols)
            data_empty = all(not row[i].strip() for i in range(num_columns) if i in data_cols)

            if has_text and data_empty:
                for i in range(num_columns):
                    if row[i].strip():
                        merged[-1][i] = (merged[-1][i] + " " + row[i]).strip()
            else:
                merged.append(row)

        return merged

    # ─── Improved Header & Metadata Analysis ─────────────────────────

    def _process_header_layout(self, header_rows):
        """
        Split header rows into 'left' (Sender) and 'right' (Receiver) columns
        if a significant horizontal gap exists.
        """
        if not header_rows:
            return {"left": [], "right": []}

        # Calculate page width approximation
        all_words = [w for r in header_rows for w in r]
        if not all_words:
            return {"left": [], "right": []}
            
        min_x = min(w['x_min'] for w in all_words)
        max_x = max(w['x_max'] for w in all_words)
        page_width = max_x - min_x
        
        left_col = []
        right_col = []

        for row in header_rows:
            # Sort words by X
            row.sort(key=lambda w: w['x_min'])
            
            # Find largest gap in this row
            max_gap = 0
            gap_idx = -1
            
            for i in range(len(row) - 1):
                gap = row[i+1]['x_min'] - row[i]['x_max']
                if gap > max_gap:
                    max_gap = gap
                    gap_idx = i
            
            # If gap is > 20% of page width, split
            if max_gap > page_width * 0.20:
                left_part = row[:gap_idx+1]
                right_part = row[gap_idx+1:]
                
                l_text = " ".join(w['text'] for w in left_part).strip()
                r_text = " ".join(w['text'] for w in right_part).strip()
                
                if l_text: left_col.append(l_text)
                if r_text: right_col.append(r_text)
            else:
                # No split - assume distinct block logic or single col
                # If aligned left (< 30% width), add to left. If aligned right (> 50%), add to right.
                text = " ".join(w['text'] for w in row).strip()
                avg_x = np.mean([w['x_center'] for w in row])
                
                if avg_x < (min_x + page_width * 0.4):
                    left_col.append(text)
                elif avg_x > (min_x + page_width * 0.6):
                    right_col.append(text)
                else:
                    # Center align - add to left for now as default
                    left_col.append(text)

        return {"left": left_col, "right": right_col}

    def _extract_metadata(self, text_lines):
        """
        Extract key metadata (Invoice #, Date, Amounts) using Regex.
        """
        metadata = {}
        full_text = "\n".join(text_lines)
        
        # Regex Patterns
        # Note: [ \t] prevents matching across newlines for the key-value separator, 
        # avoiding "TAX INVOICE \n Date" -> Invoice=Date
        patterns = {
            "invoice_no": r'(?i)\b(?:invoice|inv|ref)\b[ \t]*[:#.]?[ \t]*([a-zA-Z0-9-]*\d[a-zA-Z0-9-]*)',
            "date": r'(?i)\b(?:date|dated)\b\s*[:.]?\s*([A-Za-z]{3,9}\s+\d{1,2},?\s*\d{4}|\d{2}[/-]\d{2}[/-]\d{2,4})',
            "total_amount": r'(?i)\b(?:total|due|balance)\b\s*(?:amount)?\s*[:.]?\s*[\$Ss]?\s*([\d,]+\.?\d{2})'
        }

        for key, pattern in patterns.items():
            match = re.search(pattern, full_text)
            if match:
                val = match.group(1).strip()
                # Clean up extracted value
                if key == "total_amount":
                    val = val.replace(',', '').replace(' ', '')
                metadata[key] = val
        
        return metadata

    # ─── Main Extraction ─────────────────────────────────────────────

    def extract_full_data(self, json_path):
        """
        Extract ALL data from an EasyOCR JSON file, separated into:
        - Header Info: Text appearing above the table (e.g. company, address)
        - Table: Structured rows and columns
        - Footer Info: Text appearing below the table (e.g. totals, tax)
        
        Returns:
        {
            "header_info": [["Line 1"], ["Line 2"]],
            "table": {"headers": [...], "rows": [...]},
            "footer_info": [["Subtotal", "$100"], ...]
        }
        """
        try:
            with open(json_path, 'r') as f:
                data = json.load(f)
        except Exception as e:
            print(f"    Error loading JSON: {e}")
            return None

        if not data:
            return None

        # Filter low-confidence noise
        data = [item for item in data if item.get('confidence', 0) >= self.min_confidence]

        if not data:
            return None

        # Build word objects with spatial coordinates
        words = []
        for item in data:
            bbox = np.array(item['bbox'])
            mn, mx = np.min(bbox, axis=0), np.max(bbox, axis=0)
            words.append({
                'text':       item['text'],
                'x_min':      float(mn[0]), 'x_max': float(mx[0]),
                'y_min':      float(mn[1]), 'y_max': float(mx[1]),
                'x_center':   float(np.mean(bbox[:, 0])),
                'y_center':   float(np.mean(bbox[:, 1])),
                'height':     float(mx[1] - mn[1]),
                'confidence': item.get('confidence', 0),
            })

        if not words:
            return None

        # Step 1: Cluster into rows
        rows = self._cluster_rows(words)
        if not rows:
            return None

        # Step 2: Find header using expanded keyword system
        h_idx, h_words = self._find_header_row(rows)

        # Step 3: Split into sections
        header_rows_raw = []
        table_rows_raw = []
        footer_rows_raw = []

        if h_idx != -1:
            # Header info is everything above the header row
            header_rows_raw = rows[:h_idx]
            
            # Identify where table ends and footer begins
            split_idx = len(rows) 
            
            # Simple heuristic: scan rows after header for footer keywords
            # or significantly different structure (e.g. just label and value)
            footer_keywords = ["total", "subtotal", "tax", "amount", "due", "balance", "shipping", "handling"]
            
            for i in range(h_idx + 1, len(rows)):
                row_text = " ".join(w['text'].lower() for w in rows[i])
                if any(row_text.startswith(k) for k in footer_keywords):
                    split_idx = i
                    break
                if sum(1 for k in footer_keywords if k in row_text) >= 1 and len(rows[i]) <= 3:
                     split_idx = i
                     break

            # Need to fix split headers BEFORE splitting rows, 
            # but current _merge_split_headers assumes it operates on the whole list.
            # We will merge split headers on the whole 'rows' list first if needed.
            if h_idx != -1:
                rows = self._merge_split_headers(rows, h_idx)
                # Re-fetch h_words in case it changed (though usually index doesn't change)
                h_words = rows[h_idx]
                
            # Re-slice because rows might have changed length/content (though merge only merges into header usually)
            # Actually _merge_split_headers removes elements, so indices shift.
            # But it only merges h_idx+1 into h_idx. So h_idx stays same.
            
            # Recalculate split_idx since rows might have shifted? 
            # _merge_split_headers can reduce length.
            # Let's re-run split detection on the potentially modified rows
            header_rows_raw = rows[:h_idx]
            
            split_idx = len(rows)
            for i in range(h_idx + 1, len(rows)):
                row_text = " ".join(w['text'].lower() for w in rows[i])
                if any(row_text.startswith(k) for k in footer_keywords):
                    split_idx = i
                    break
                if sum(1 for k in footer_keywords if k in row_text) >= 1 and len(rows[i]) <= 3:
                     split_idx = i
                     break
            
            table_rows_raw = rows[h_idx + 1 : split_idx]
            footer_rows_raw = rows[split_idx:]

            print(f"    Layout: header-based (row {h_idx})")
            h_sorted = sorted(h_words, key=lambda w: w['x_min'])
            headers = [self.clean_text(w['text']) for w in h_sorted]
            # Use table words + header words for column boundaries
            table_words_flat = [w for r in table_rows_raw for w in r] + h_sorted
            bounds = self._boundaries_from_headers(h_sorted, table_words_flat if table_words_flat else h_sorted)
        else:
            print("    Layout: gutter-based")
            headers = []
            bounds = []
            header_rows_raw = []
            table_rows_raw = rows
            footer_rows_raw = []
            
            all_w = [w for r in rows for w in r]
            bounds = self._boundaries_from_gutters(rows, all_w)
            n = len(bounds) - 1
            headers = [f"Column {i+1}" for i in range(n)]

        # --- Process Header Info ---
        # Old flat method
        # header_info = []
        # for row in header_rows_raw:
        #     row.sort(key=lambda w: w['x_min'])
        #     text = " ".join(w['text'] for w in row).strip()
        #     if text:
        #         header_info.append([text])
        
        # New Split Method
        header_split = self._process_header_layout(header_rows_raw)
        
        # Gather all text lines for metadata extraction
        all_header_lines = header_split["left"] + header_split["right"]
        metadata = self._extract_metadata(all_header_lines)

        # --- Process Table Data ---
        num_cols = len(bounds) - 1
        table_data = []
        
        for row in table_rows_raw:
            mapped = [""] * num_cols
            for w in row:
                txt = self.clean_text(w['text'])
                best_col, best_ov = -1, 0
                for i in range(num_cols):
                    ov = self._compute_overlap(w['x_min'], w['x_max'],
                                               bounds[i], bounds[i + 1])
                    if ov > best_ov:
                        best_ov, best_col = ov, i
                if best_col == -1:
                    for i in range(num_cols):
                        if bounds[i] <= w['x_center'] < bounds[i + 1]:
                            best_col = i
                            break
                if best_col != -1:
                    mapped[best_col] = (mapped[best_col] + " " + txt).strip()
            if any(c.strip() for c in mapped):
                table_data.append(mapped)

        if h_idx != -1:
            table_data = self._merge_continuation_rows(table_data, num_cols)

        # Check if first data row is a header (gutter fallback)
        if not headers and table_data:
            first = " ".join(table_data[0]).lower()
            hk = ["description", "qty", "price", "amount", "total",
                  "label", "value", "calories", "serving", "nutrition"]
            if sum(1 for k in hk if k in first) >= 2:
                headers = table_data[0]
                table_data = table_data[1:]

        table_data = [[c.strip() for c in row] for row in table_data]
        table_data = [row for row in table_data if any(row)]

        # --- Process Footer Info ---
        footer_info = []
        for row in footer_rows_raw:
            row.sort(key=lambda w: w['x_min'])
            texts = [self.clean_text(w['text']) for w in row]
            
            # Split label/value if last item matches money pattern
            if len(texts) >= 2:
                last = texts[-1]
                # Matches $10.00, 10.00, $ 10.00
                if re.match(r'^[\$Ss]?\s*\d+[.,]?\d*$', last):
                    label = " ".join(texts[:-1])
                    footer_info.append([label, last])
                else:
                    footer_info.append([" ".join(texts)])
            else:
                footer_info.append([" ".join(texts)])

        return {
            "header_split": header_split,
            "metadata": metadata,
            "table": {"headers": headers, "rows": table_data},
            "footer_info": footer_info
        }

    def extract_from_json(self, json_path):
        """
        Wrapper for backward compatibility. Returns just the table dict.
        """
        full = self.extract_full_data(json_path)
        if full and full.get("table"):
            return full["table"]
        return None