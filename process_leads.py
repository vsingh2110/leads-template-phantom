#!/usr/bin/env python3
"""
process_leads.py — Convert Digital Marketing Leads CSVs to linkedin_upload.csv
"""

import csv
import re
import os
import sys

# ---------------------------------------------------------------------------
# Country detection
# ---------------------------------------------------------------------------

# Ordered list of (prefix_digits_only, country_code) — longest first so that
# e.g. +971 is tested before +97, +91 before +9, etc.
PHONE_PREFIX_MAP = [
    # 4-digit prefixes
    ("8801", "BD"),
    # 3-digit prefixes
    ("880", "BD"),
    ("977", "NP"),
    ("998", "UZ"),
    ("971", "AE"),
    ("974", "QA"),
    ("973", "BH"),
    ("972", "IL"),
    ("968", "OM"),
    ("967", "YE"),
    ("966", "SA"),
    ("965", "KW"),
    ("964", "IQ"),
    ("963", "SY"),
    ("960", "MV"),
    ("852", "HK"),
    ("591", "BO"),
    ("593", "EC"),
    ("254", "KE"),
    ("256", "UG"),
    ("260", "ZM"),
    ("251", "ET"),
    ("234", "NG"),
    ("233", "GH"),
    ("243", "CD"),
    ("218", "LY"),
    ("216", "TN"),
    ("213", "DZ"),
    ("212", "MA"),
    ("211", "SS"),
    ("252", "SO"),
    # 2-digit prefixes
    ("93", "AF"),
    ("92", "PK"),
    ("91", "IN"),
    ("90", "TR"),
    ("86", "CN"),
    ("82", "KR"),
    ("81", "JP"),
    ("66", "TH"),
    ("65", "SG"),
    ("64", "NZ"),
    ("63", "PH"),
    ("62", "ID"),
    ("61", "AU"),
    ("60", "MY"),
    ("58", "VE"),
    ("55", "BR"),
    ("54", "AR"),
    ("44", "GB"),
    ("43", "AT"),
    ("41", "CH"),
    ("39", "IT"),
    ("36", "HU"),
    ("34", "ES"),
    ("33", "FR"),
    ("32", "BE"),
    ("31", "NL"),
    ("30", "GR"),
    ("27", "ZA"),
    ("20", "EG"),
    # 1-digit prefixes
    ("7", "RU"),
    ("1", "US"),
]


def detect_country_from_phone(phone_str):
    """
    Return ISO country code if phone_str has an explicit international prefix
    (+XX or 00XX at the start).  Bare digits (no leading + or 00) → None.
    """
    if not phone_str:
        return None
    s = str(phone_str).strip()
    # Remove common formatting characters but keep leading + and digits
    s = re.sub(r"[\s\-\.\(\)]", "", s)
    if not s:
        return None

    # Extract digits after the explicit international prefix
    if s.startswith("+"):
        digits = re.sub(r"\D", "", s[1:])  # strip non-digits after +
        has_prefix = True
    elif s.startswith("00"):
        digits = re.sub(r"\D", "", s[2:])  # strip non-digits after 00
        has_prefix = True
    else:
        # No explicit international prefix — treat as local (Indian) number
        return None

    if not has_prefix or not digits:
        return None

    for prefix, country in PHONE_PREFIX_MAP:
        if digits.startswith(prefix):
            return country

    return None


def detect_country_from_text(text):
    """Detect country from free-form text (Location, Country lines)."""
    if not text:
        return None
    text_lower = text.lower()
    country_keywords = {
        "india": "IN",
        "indian": "IN",
        "bangladesh": "BD",
        "nepal": "NP",
        "pakistan": "PK",
        "uae": "AE",
        "united arab emirates": "AE",
        "dubai": "AE",
        "saudi": "SA",
        "saudi arabia": "SA",
        "kuwait": "KW",
        "qatar": "QA",
        "bahrain": "BH",
        "oman": "OM",
        "iraq": "IQ",
        "syria": "SY",
        "jordan": "JO",
        "turkey": "TR",
        "china": "CN",
        "hong kong": "HK",
        "singapore": "SG",
        "malaysia": "MY",
        "thailand": "TH",
        "philippines": "PH",
        "indonesia": "ID",
        "australia": "AU",
        "new zealand": "NZ",
        "japan": "JP",
        "south korea": "KR",
        "korea": "KR",
        "uk": "GB",
        "united kingdom": "GB",
        "england": "GB",
        "germany": "DE",
        "france": "FR",
        "spain": "ES",
        "italy": "IT",
        "netherlands": "NL",
        "belgium": "BE",
        "switzerland": "CH",
        "austria": "AT",
        "greece": "GR",
        "portugal": "PT",
        "usa": "US",
        "united states": "US",
        "america": "US",
        "canada": "CA",
        "brazil": "BR",
        "argentina": "AR",
        "africa": "ZA",
        "south africa": "ZA",
        "nigeria": "NG",
        "ghana": "GH",
        "kenya": "KE",
        "ethiopia": "ET",
        "egypt": "EG",
        "morocco": "MA",
        "uzbekistan": "UZ",
        "russia": "RU",
        "yemen": "YE",
        "somalia": "SO",
        "israel": "IL",
    }
    # Check multi-word first (longer matches)
    for kw in sorted(country_keywords.keys(), key=len, reverse=True):
        if kw in text_lower:
            return country_keywords[kw]
    return None


# ---------------------------------------------------------------------------
# Name cleaning
# ---------------------------------------------------------------------------

PREFIX_RE = re.compile(
    r"^((?:Dr|Mr|Mrs|Ms|Prof(?:essor)?|Er|Smt|Shri|Sh)\.?\s+|"
    r"(?:Dr|Mr|Mrs|Ms|Prof(?:essor)?|Er)\.\s*(?=\S))",
    re.IGNORECASE,
)


def clean_name(raw):
    """Strip prefixes, tildes, extra whitespace. Return (firstname, lastname)."""
    if not raw:
        return "", ""
    name = str(raw).strip()
    # Strip tilde
    name = name.lstrip("~").strip()
    # Strip prefixes repeatedly (e.g. "Dr. Mr. Someone")
    for _ in range(5):
        m = PREFIX_RE.match(name)
        if m:
            name = name[m.end():].strip()
        else:
            break
    # Remove embedded newlines / carriage returns
    name = re.sub(r"[\r\n]+", " ", name).strip()
    # Collapse multiple spaces
    name = re.sub(r"\s{2,}", " ", name)
    parts = name.split(None, 1)
    if not parts:
        return "", ""
    firstname = parts[0].strip()
    lastname = parts[1].strip() if len(parts) > 1 else ""
    return firstname, lastname


# ---------------------------------------------------------------------------
# Filtering
# ---------------------------------------------------------------------------

FILTER_PHRASES = [
    "wrong query", "fake query", "wrong no", "wrong number",
    "no is not correct", "invalid no", "no. is incorrect",
    "irrelevant", "irrelevent", "job seeker", "not needed",
    "no. does not exist", "person is fake",
]


def should_filter_row(remarks):
    """Return True if the row should be removed based on remarks text."""
    if not remarks:
        return False
    r = str(remarks).lower()
    return any(phrase in r for phrase in FILTER_PHRASES)


SECTION_HEADER_RE = re.compile(
    r"^(january|february|march|april|may|june|july|august|september|"
    r"october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec"
    r"|\d{4}|sl\.?\s*no|s\.no|date)$",
    re.IGNORECASE,
)


def is_header_or_empty(value):
    if not value or not str(value).strip():
        return True
    return bool(SECTION_HEADER_RE.match(str(value).strip()))


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def read_csv(path):
    for enc in ["utf-8-sig", "utf-8", "latin-1", "cp1252"]:
        try:
            with open(path, encoding=enc, newline="") as f:
                rows = list(csv.reader(f))
            return rows
        except (UnicodeDecodeError, Exception):
            continue
    return []


EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")


def extract_email(text):
    if not text:
        return ""
    m = EMAIL_RE.search(str(text))
    return m.group(0).lower().strip() if m else ""


def clean_field(val):
    """Strip whitespace and remove embedded newlines."""
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"[\r\n]+", " ", s).strip()
    s = re.sub(r"\s{2,}", " ", s)
    if s.lower() in ("nan", "none", "n/a", "na", "-", "--", "---"):
        return ""
    return s


def make_record(email, firstname, lastname, jobtitle, company, country):
    return {
        "email": clean_field(email).lower(),
        "firstname": clean_field(firstname),
        "lastname": clean_field(lastname),
        "jobtitle": clean_field(jobtitle),
        "employeecompany": clean_field(company),
        "country": clean_field(country),
        "googleaid": "",
    }


# ---------------------------------------------------------------------------
# Freeform text parser (for after-25-july-2023, 2024, 2025, 2026, intl)
# ---------------------------------------------------------------------------

def parse_freeform(text):
    """
    Parse a freeform details block and return dict with keys:
    name, email, phone, company, jobtitle, country
    """
    if not text:
        return {}
    text = str(text)
    result = {}

    # Email
    m = EMAIL_RE.search(text)
    if m:
        result["email"] = m.group(0).lower().strip()

    # Phone — look for explicit prefixed numbers or long digit strings
    phone_patterns = [
        r"\+[\d\s\-\.\(\)]{7,20}",   # +XX...
        r"00\d[\d\s\-\.]{6,18}",     # 00XX...
        r"\b\d{10,15}\b",            # bare long numbers
    ]
    for pp in phone_patterns:
        pm = re.search(pp, text)
        if pm:
            result["phone"] = pm.group(0).strip()
            break

    # Name — look for labelled lines first
    name_m = re.search(
        r"(?:^|\n)\s*(?:Name|Full\s*name)\s*[:\-]\s*(.+)",
        text, re.IGNORECASE
    )
    if name_m:
        result["name"] = name_m.group(1).strip()
    else:
        # Fall back: first non-empty line that looks like a name
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            # Skip lines that are clearly phone/email/url/date
            if re.match(r"^[\d\s\+\-\.\(\)]{5,}$", line):
                continue
            if "@" in line or "http" in line.lower():
                continue
            if re.match(
                r"^(Source|Mobile|Number|Phone|Need|Profile|Location|"
                r"City|Country|Hospital|Clinic|Enquiry|Email|Gmail|"
                r"Whatsapp|What\'s|Requirement|Date)\s*[:\-]",
                line, re.IGNORECASE
            ):
                continue
            # Skip forex/occupation-only lines
            if re.match(r"^(Forex|Trader|Engineer|Doctor|Nurse)\s*$", line, re.IGNORECASE):
                continue
            result["name"] = line
            break

    # Company / Hospital
    comp_m = re.search(
        r"(?:Hospital|Clinic|Hospital/Clinic|Centre|Center|Company|Organisation|Organization)\s*[:\-]\s*(.+)",
        text, re.IGNORECASE
    )
    if comp_m:
        result["company"] = comp_m.group(1).strip()

    # Job title — extract from Profile/Designation line
    prof_m = re.search(
        r"(?:Profile|Designation)\s*[:\-]\s*(.+)", text, re.IGNORECASE
    )
    if prof_m:
        profile_text = prof_m.group(1).strip()
        at_idx = profile_text.lower().find(" at ")
        title_part = profile_text[:at_idx].strip() if at_idx >= 0 else profile_text
        company_part = profile_text[at_idx + 4:].strip() if at_idx >= 0 else ""
        if "jobtitle" not in result:
            result["jobtitle"] = title_part
        if "company" not in result and company_part:
            result["company"] = company_part

    # Country
    country_m = re.search(
        r"(?:Country|Location)\s*[:\-]\s*(.+)", text, re.IGNORECASE
    )
    if country_m:
        country_text = country_m.group(1).strip()
        cc = detect_country_from_text(country_text)
        if cc:
            result["country"] = cc
        else:
            result["country_text"] = country_text

    if "country" not in result and "phone" in result:
        cc = detect_country_from_phone(result["phone"])
        if cc:
            result["country"] = cc

    if "country" not in result:
        # Try to detect country from full text (city/location mentions)
        city_m = re.search(r"(?:City|Location)\s*[:\-]\s*(.+)", text, re.IGNORECASE)
        if city_m:
            cc = detect_country_from_text(city_m.group(1))
            if cc:
                result["country"] = cc

    return result


# ---------------------------------------------------------------------------
# Sheet parsers
# ---------------------------------------------------------------------------

def _parse_contact_cell(raw_text):
    """
    Parse a raw (unstripped) cell value that may contain newlines, commas,
    email, phone and name mixed together.
    Returns (name, email, phone).
    """
    text = str(raw_text).strip() if raw_text else ""
    email_out = ""
    phone_out = ""
    name_out = ""

    # Handle comma-separated: Name,Surname,email,phone
    if "," in text and "@" in text:
        parts = [p.strip() for p in text.split(",")]
        name_parts = []
        for p in parts:
            if "@" in p:
                email_out = extract_email(p)
            elif re.match(r"^[\d\s\+\-\.]{7,}$", p):
                phone_out = p
            else:
                name_parts.append(p)
        name_out = " ".join(name_parts).strip()
        return name_out, email_out, phone_out

    # Single-line "phone name" pattern (digits then name, separated by space or comma)
    stripped = text.replace("\n", " ").strip()
    # Try "phone name" with space separator
    single_m = re.match(r"^(\+?[\d\s\-\.]{10,16})\s+(.{3,50})$", stripped)
    if single_m and not re.search(r"[a-zA-Z]", single_m.group(1)):
        phone_out = single_m.group(1).strip()
        name_out = single_m.group(2).strip()
        # Strip trailing location like "from ASSAM"
        name_out = re.sub(r"\s+from\s+\S+$", "", name_out, flags=re.IGNORECASE).strip()
        return name_out, email_out, phone_out
    # Try "phone, name" with comma separator
    comma_phone_m = re.match(r"^(\d{10,15}),\s*(.{3,50})$", stripped)
    if comma_phone_m:
        phone_out = comma_phone_m.group(1).strip()
        name_out = comma_phone_m.group(2).strip()
        return name_out, email_out, phone_out

    # Split on newlines and classify each line
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    name_candidates = []
    for line in lines:
        em = extract_email(line)
        if em:
            email_out = email_out or em
            continue
        if re.match(r"^[\d\s\+\-\.\(\)]{7,}$", line) or re.match(r"^\+\d[\d\s\-\.]{5,}$", line):
            phone_out = phone_out or line.strip()
            continue
        if re.match(r"^(Whatsapp|Call me|Call|http|Tel|Fax|Mob|Note|From)\b", line, re.IGNORECASE):
            continue
        if line.count(",") >= 2 and len(line) > 40:
            continue
        if re.search(r"\d.*\b(road|street|nagar|colony|plot|shop|sector|floor)\b", line, re.IGNORECASE):
            continue
        # Skip lines starting with digits (likely phone numbers mixed with text)
        if re.match(r"^\d{7,}", line):
            # But try to extract the name part after the digits
            dm = re.match(r"^\d{7,}\s+(.{3,})$", line)
            if dm:
                candidate = dm.group(1).strip()
                candidate = re.sub(r"\s+from\s+\S+$", "", candidate, flags=re.IGNORECASE).strip()
                if not re.search(r"\d", candidate):
                    name_candidates.append(candidate)
            digit_m = re.match(r"^(\d+)", line)
            if digit_m:
                phone_out = phone_out or digit_m.group(1)
            continue
        # Skip lines that look like company/institution names
        if re.search(r"\b(diagnostic|diagnostics|hospital|clinic|centre|center|"
                     r"medical|healthcare|pharma|labs|laboratory|institute|"
                     r"scanning|imaging)\b", line, re.IGNORECASE):
            continue
        # Remove trailing "from LOCATION"
        line = re.sub(r"\s+from\s+\S+.*$", "", line, flags=re.IGNORECASE).strip()
        name_candidates.append(line)

    # From name candidates pick the best (short, no digits, person-like)
    for cand in name_candidates:
        words = cand.split()
        if 1 <= len(words) <= 4 and not re.search(r"\d", cand):
            name_out = cand
            break
    # Fallback: first candidate that doesn't contain digits
    if not name_out:
        for cand in name_candidates:
            if not re.search(r"\d", cand):
                name_out = cand
                break

    return name_out, email_out, phone_out


def parse_followups(rows):
    """
    Followups — NO header row.
    Columns: [0]=slno, [1]=date, [2]=center name, [3]=address, [4]=pincode,
              [5]=contact person, [6]=contact no, [7]=email, [8]=remarks
    Some rows have data shifted to unexpected columns.
    Uses raw (pre-clean_field) values to detect multiline cells.
    """
    records = []
    for row in rows:
        # Pad to at least 9 columns
        row = row + [""] * (9 - len(row))

        # Skip completely empty rows
        if all(not str(c).strip() for c in row[:8]):
            continue

        remarks = clean_field(row[8])
        if should_filter_row(remarks):
            continue

        # Use raw values (preserving newlines) for cells that may be multiline
        raw4 = str(row[4]) if row[4] else ""
        raw5 = str(row[5]) if row[5] else ""
        raw6 = str(row[6]) if row[6] else ""

        center = clean_field(row[2])
        phone = ""
        email = extract_email(row[7])

        name = ""
        extracted_phone = ""
        extracted_email = ""

        # Check if col[4] looks like a contact entry (multiline or comma+email)
        col4_is_contact = (
            "\n" in raw4 or
            ("," in raw4 and "@" in raw4) or
            re.match(r"^\d{10,}\s+\S", raw4.strip())
        )

        col5_stripped = raw5.strip()

        if col4_is_contact:
            # Prefer col[4] as the contact data source
            name, extracted_email, extracted_phone = _parse_contact_cell(raw4)
        elif col5_stripped:
            # col[5] is the expected contact person column
            # Reject if it looks like a whatsapp/note entry
            if re.match(r"^(Whatsapp|Phone|Call|Tel|oen\b)", col5_stripped, re.IGNORECASE):
                # Fall back to col[4]
                if raw4.strip() and not re.match(r"^\d{5,6}$", raw4.strip()):
                    name, extracted_email, extracted_phone = _parse_contact_cell(raw4)
            elif "\n" in raw5 or ("," in raw5 and "@" in raw5):
                name, extracted_email, extracted_phone = _parse_contact_cell(raw5)
            else:
                col5_clean = clean_field(raw5)
                m = re.match(r"^(\d{10,15})\s+(.+)$", col5_clean)
                if m:
                    extracted_phone = m.group(1)
                    name = m.group(2).strip()
                else:
                    name = col5_clean

        # If still no name, try col[4] as fallback
        if not name and raw4.strip():
            raw4s = raw4.strip()
            if re.match(r"^\d{5,6}$", raw4s):
                pass  # real pincode
            else:
                name, extracted_email, extracted_phone = _parse_contact_cell(raw4)

        # Try col[6] for multi-line name+phone data (some rows have it there)
        if not name and "\n" in raw6:
            name2, em2, ph2 = _parse_contact_cell(raw6)
            if name2:
                name = name2
                if not phone:
                    phone = ph2
                if not email:
                    email = em2

        # Fill in extracted values
        if not email:
            email = extracted_email
        if not phone:
            phone = extracted_phone or clean_field(row[6])

        # Derive company from col[3] if center is empty (shifted rows)
        if not center:
            col3 = clean_field(row[3])
            if col3 and not re.match(r"^\d{5,6}$", col3):
                if len(col3) < 60 and col3.count(",") <= 2:
                    center = col3

        fn, ln = clean_name(name)
        if not fn:
            continue

        country = detect_country_from_phone(phone) or "IN"
        addr = clean_field(row[3])
        if country == "IN":
            country = detect_country_from_text(addr) or "IN"

        records.append(make_record(email, fn, ln, "", center, country))
    return records


def parse_calls_to_do(rows):
    """
    Calls to Do — has header row.
    Columns: sl no., Date, center name, address, pincode, contact person, contact no, email, remarks
    """
    records = []
    header_found = False
    col_map = {}

    for row in rows:
        if not header_found:
            # Detect header row
            row_lower = [str(c).lower().strip() for c in row]
            if any("sl no" in c or "contact person" in c for c in row_lower):
                header_found = True
                for i, h in enumerate(row_lower):
                    if "sl no" in h:
                        col_map["slno"] = i
                    elif "center name" in h:
                        col_map["center"] = i
                    elif "contact person" in h:
                        col_map["person"] = i
                    elif "contact no" in h:
                        col_map["phone"] = i
                    elif "email" in h:
                        col_map["email"] = i
                    elif "remark" in h:
                        col_map["remarks"] = i
                    elif "adress" in h or "address" in h:
                        col_map["address"] = i
                    elif "pincode" in h:
                        col_map["pincode"] = i
            continue

        row = row + [""] * 14
        slno = row[col_map.get("slno", 0)]
        if is_header_or_empty(slno) and is_header_or_empty(row[col_map.get("center", 2)]):
            continue

        center = clean_field(row[col_map.get("center", 2)])
        raw_person = str(row[col_map.get("person", 5)]) if row[col_map.get("person", 5)] else ""
        raw_pincode = str(row[col_map.get("pincode", 4)]) if row[col_map.get("pincode", 4)] else ""
        phone = clean_field(row[col_map.get("phone", 6)])
        email = extract_email(row[col_map.get("email", 7)])
        remarks = clean_field(row[col_map.get("remarks", 8)])
        address = clean_field(row[col_map.get("address", 3)])

        if should_filter_row(remarks):
            continue

        contact_person = clean_field(raw_person)

        # Handle comma-separated or multiline entries in contact_person
        if contact_person and ("\n" in raw_person or ("," in contact_person and "@" in contact_person)):
            name, em, ph = _parse_contact_cell(raw_person)
            contact_person = name
            if not email:
                email = em
            if not phone:
                phone = ph
        elif contact_person and "," in contact_person and "@" in contact_person:
            name, em, ph = _parse_contact_cell(contact_person)
            contact_person = name
            if not email:
                email = em
            if not phone:
                phone = ph
        elif contact_person and re.match(r"^\d{10,15},\s*\S", contact_person):
            # "phone, name" pattern
            name, em, ph = _parse_contact_cell(contact_person)
            contact_person = name
            if not phone:
                phone = ph

        # Handle "9949034747 Dr Ajaykumar" style — phone followed by name
        if contact_person:
            phone_name_m = re.match(r"^(\d{10,15})\s+(.+)$", contact_person)
            if phone_name_m:
                if not phone:
                    phone = phone_name_m.group(1)
                contact_person = phone_name_m.group(2).strip()

        # If contact_person is a pure phone number, it's a shifted row — look elsewhere
        if contact_person and re.match(r"^\d{7,}$", contact_person.strip()):
            if not phone:
                phone = contact_person.strip()
            contact_person = ""

        # If contact_person is empty, look in adjacent columns for name-like text
        if not contact_person:
            # Check pincode column for multiline name+phone data
            if raw_pincode.strip() and not re.match(r"^\d{5,6}$", raw_pincode.strip()):
                name, em, ph = _parse_contact_cell(raw_pincode)
                if name:
                    contact_person = name
                    if not email:
                        email = em
                    if not phone:
                        phone = ph
            if not contact_person:
                addr_val = clean_field(row[col_map.get("address", 3)])
                # Only use address as name if it's short and name-like
                if addr_val and len(addr_val) < 50 and not re.search(r"\d{5}", addr_val):
                    # heuristic: doesn't contain many commas
                    if addr_val.count(",") <= 1:
                        contact_person = addr_val

        fn, ln = clean_name(contact_person)
        if not fn:
            continue

        # Keep rows with no email if they have a name (LinkedIn can match on name+company)
        country = detect_country_from_phone(phone) or "IN"
        if country == "IN":
            country = detect_country_from_text(address) or "IN"

        records.append(make_record(email, fn, ln, "", center, country))
    return records


def parse_old_lead_sheet(rows):
    """
    Old Lead Sheet.
    Columns: S.No, Date, Requirement, City, State, Hospital/Centre, Person, Contact no, E-Mail ID, ...
    """
    records = []
    header_found = False
    col_map = {}

    for row in rows:
        if not header_found:
            row_lower = [str(c).lower().strip() for c in row]
            if any("s.no" in c or "hospital" in c or "e-mail" in c for c in row_lower):
                header_found = True
                for i, h in enumerate(row_lower):
                    if h in ("s.no", "sno"):
                        col_map["slno"] = i
                    elif "hospital" in h or "centre" in h:
                        col_map["company"] = i
                    elif "person" in h:
                        col_map["person"] = i
                    elif "contact no" in h or "contact_no" in h:
                        col_map["phone"] = i
                    elif "e-mail" in h or "email" in h:
                        col_map["email"] = i
                    elif "state" in h:
                        col_map["state"] = i
                    elif "remark" in h:
                        col_map["remarks"] = i
            continue

        row = row + [""] * 16
        slno = row[col_map.get("slno", 0)]
        if is_header_or_empty(slno) and is_header_or_empty(row[col_map.get("company", 5)]):
            continue

        company = clean_field(row[col_map.get("company", 5)])
        contact_person = clean_field(row[col_map.get("person", 6)])
        phone = clean_field(row[col_map.get("phone", 7)])
        email = extract_email(row[col_map.get("email", 8)])
        state = clean_field(row[col_map.get("state", 4)])
        remarks = clean_field(row[col_map.get("remarks", 11)])

        if should_filter_row(remarks):
            continue

        fn, ln = clean_name(contact_person)
        if not fn:
            continue

        country = detect_country_from_phone(phone)
        if not country:
            country = detect_country_from_text(state) or "IN"

        records.append(make_record(email, fn, ln, "", company, country))
    return records


def parse_new_lead(rows):
    """
    New Lead sheet.
    Columns: Date, Source of Lead, City, STATE, CENTER NAME, CENTER ADDRESS,
             Pincode, CONTACT PERSON DETAIL, E-MAIL ID, Contact No, REMARKS, Lead Type
    """
    records = []
    header_found = False
    col_map = {}

    for row in rows:
        if not header_found:
            row_lower = [str(c).lower().strip() for c in row]
            if any("center name" in c or "contact person" in c for c in row_lower):
                header_found = True
                for i, h in enumerate(row_lower):
                    if "center name" in h:
                        col_map["company"] = i
                    elif "contact person" in h:
                        col_map["person"] = i
                    elif "e-mail" in h or "email" in h:
                        col_map["email"] = i
                    elif "contact no" in h:
                        col_map["phone"] = i
                    elif "state" in h:
                        col_map["state"] = i
                    elif "remark" in h:
                        col_map["remarks"] = i
            continue

        row = row + [""] * 18
        company = clean_field(row[col_map.get("company", 4)])
        contact_person = clean_field(row[col_map.get("person", 7)])
        email_raw = clean_field(row[col_map.get("email", 8)])
        phone_raw = clean_field(row[col_map.get("phone", 9)])
        state = clean_field(row[col_map.get("state", 3)])
        remarks = clean_field(row[col_map.get("remarks", 10)])

        if should_filter_row(remarks):
            continue

        # Email might be in the phone column
        email = extract_email(email_raw) or extract_email(phone_raw)
        phone = phone_raw if not extract_email(phone_raw) else ""

        fn, ln = clean_name(contact_person)
        if not fn:
            continue

        country = detect_country_from_phone(phone)
        if not country:
            country = detect_country_from_text(state) or "IN"

        records.append(make_record(email, fn, ln, "", company, country))
    return records


def parse_after_march_2023(rows):
    """
    After March 2023.
    Columns: Date, Source of Lead, Requirement, City, STATE,
             CENTER/Hospital NAME, CENTER ADDRESS, Pincode,
             CONTACT PERSON DETAIL, E-MAIL ID, Contact No, REMARKS, Lead Type, Remark
    """
    records = []
    header_found = False
    col_map = {}

    for row in rows:
        if not header_found:
            row_lower = [str(c).lower().strip() for c in row]
            if any("center" in c or "hospital" in c or "contact person" in c for c in row_lower):
                header_found = True
                for i, h in enumerate(row_lower):
                    if "center" in h or "hospital" in h:
                        col_map["company"] = i
                    elif "contact person" in h:
                        col_map["person"] = i
                    elif "e-mail" in h or "email" in h:
                        col_map["email"] = i
                    elif "contact no" in h:
                        col_map["phone"] = i
                    elif "state" in h:
                        col_map["state"] = i
                    elif "remark" in h:
                        col_map.setdefault("remarks", i)
            continue

        row = row + [""] * 16
        company = clean_field(row[col_map.get("company", 5)])
        contact_person = clean_field(row[col_map.get("person", 8)])
        email_raw = clean_field(row[col_map.get("email", 9)])
        phone_raw = clean_field(row[col_map.get("phone", 10)])
        state = clean_field(row[col_map.get("state", 4)])
        remarks = clean_field(row[col_map.get("remarks", 11)])

        if should_filter_row(remarks):
            continue

        email = extract_email(email_raw) or extract_email(phone_raw)
        phone = phone_raw if not extract_email(phone_raw) else ""

        fn, ln = clean_name(contact_person)
        if not fn:
            continue

        country = detect_country_from_phone(phone)
        if not country:
            country = detect_country_from_text(state) or "IN"

        records.append(make_record(email, fn, ln, "", company, country))
    return records


def parse_freeform_sheet(rows, details_col=1, remarks_col=2):
    """
    Generic freeform sheet parser (after 25 july 2023, 2024, 2025, 2026,
    International Leads).  The first row is the header.
    """
    records = []
    first_row = True
    for row in rows:
        if first_row:
            first_row = False
            # Detect column indices from header
            row_lower = [str(c).lower().strip() for c in row]
            for i, h in enumerate(row_lower):
                if "detail" in h or "details" in h:
                    details_col = i
                elif "remark" in h:
                    remarks_col = i
            continue

        row = row + [""] * (max(details_col, remarks_col) + 2)
        # Use raw value (preserving newlines) for details so parse_freeform can split by lines
        raw_details = str(row[details_col]) if row[details_col] else ""
        details_clean = clean_field(raw_details)
        remarks = clean_field(row[remarks_col])

        if not details_clean:
            continue
        if should_filter_row(remarks) or should_filter_row(details_clean):
            continue

        parsed = parse_freeform(raw_details)
        if not parsed:
            continue

        name = parsed.get("name", "")
        fn, ln = clean_name(name)
        email = parsed.get("email", "")
        company = clean_field(parsed.get("company", ""))
        jobtitle = clean_field(parsed.get("jobtitle", ""))
        country = parsed.get("country", "")

        if not country:
            country = detect_country_from_phone(parsed.get("phone", "")) or "IN"

        if not fn and not email:
            continue

        records.append(make_record(email, fn, ln, jobtitle, company, country))
    return records


def parse_2026(rows):
    """
    2026 file has a mixed structure — some rows use structured key:value format,
    some use freeform. We use the freeform parser with details_col=1.
    """
    return parse_freeform_sheet(rows, details_col=1, remarks_col=2)


def parse_international_leads(rows):
    """
    International Leads — same freeform structure.
    """
    return parse_freeform_sheet(rows, details_col=1, remarks_col=2)


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------

def dedup(records):
    seen_emails = set()
    seen_name_company = set()
    out = []
    for r in records:
        email = r["email"].lower().strip()
        fn = r["firstname"].lower().strip()
        ln = r["lastname"].lower().strip()
        company = r["employeecompany"].lower().strip()

        if email:
            if email in seen_emails:
                continue
            seen_emails.add(email)
        else:
            key = (fn, ln, company)
            if key in seen_name_company:
                continue
            seen_name_company.add(key)
        out.append(r)
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

BASE = os.path.dirname(os.path.abspath(__file__))

INPUT_FILES = [
    ("Digital Marketing Leads  - Followups .csv", "followups"),
    ("Digital Marketing Leads  - calls to do.csv", "calls"),
    ("Digital Marketing Leads  - old lead sheet.csv", "old_lead"),
    ("Digital Marketing Leads  - new lead.csv", "new_lead"),
    ("Digital Marketing Leads  - after march 2023.csv", "after_march"),
    ("Digital Marketing Leads  - after 25 july 2023.csv", "freeform"),
    ("Digital Marketing Leads  - Leads 2024.csv", "freeform"),
    ("Digital Marketing Leads  - 2025.csv", "freeform"),
    ("Digital Marketing Leads  - 2026.csv", "freeform"),
    ("Digital Marketing Leads  - International Leads.csv", "freeform"),
]

PARSERS = {
    "followups": parse_followups,
    "calls": parse_calls_to_do,
    "old_lead": parse_old_lead_sheet,
    "new_lead": parse_new_lead,
    "after_march": parse_after_march_2023,
    "freeform": parse_freeform_sheet,
}

OUTPUT_FILE = os.path.join(BASE, "linkedin_upload.csv")
OUTPUT_HEADERS = ["email", "firstname", "lastname", "jobtitle", "employeecompany", "country", "googleaid"]


def main():
    all_records = []
    total_raw = 0

    for filename, sheet_type in INPUT_FILES:
        path = os.path.join(BASE, filename)
        if not os.path.exists(path):
            print(f"  SKIP (not found): {filename}")
            continue
        rows = read_csv(path)
        if not rows:
            print(f"  SKIP (empty): {filename}")
            continue
        parser = PARSERS[sheet_type]
        records = parser(rows)
        total_raw += len(records)
        print(f"  {filename}: {len(records)} records")
        all_records.extend(records)

    before_dedup = len(all_records)
    all_records = dedup(all_records)
    dupes_removed = before_dedup - len(all_records)

    # Final filter: must have (email) OR (firstname AND (lastname OR employeecompany))
    final = []
    filtered_no_id = 0
    for r in all_records:
        has_email = bool(r["email"])
        has_name = bool(r["firstname"])
        if has_email or has_name:
            final.append(r)
        else:
            filtered_no_id += 1

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        writer.writerows(final)

    print(f"\nSummary:")
    print(f"  Total rows processed (pre-dedup): {total_raw}")
    print(f"  Rows kept: {len(final)}")
    print(f"  Duplicates removed: {dupes_removed}")
    print(f"  Rows filtered (no identifier): {filtered_no_id}")
    print(f"  Output: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
