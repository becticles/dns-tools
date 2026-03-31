#!/usr/bin/env python3

import csv
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# ============================================
# CONFIG
# ============================================

DNS_FOLDER = Path("/Users/brian/gitfiles/msinfra-staticdns")

VIP_WORKBOOK = Path(
    "/Users/brian/Library/CloudStorage/OneDrive-UniversityofCentralFlorida/F5 Load Balancer Project_GRP - Documents/f5_VIPs.xlsx"
)

timestamp = datetime.now().strftime("%Y%m%d_%H%M")

OUTPUT_WORKBOOK = Path(
    f"/Users/brian/Downloads/f5_VIPs_with_dns_{timestamp}.xlsx"
)

SKIP_SHEET_NAMES = {"Appliances"}

EXTERNAL_IP_COLUMN = 2
INTERNAL_IP_COLUMN = 3
START_ROW = 2

OUTPUT_INT_COUNT_COLUMN = 10
OUTPUT_INT_DNS_COLUMN = 11
OUTPUT_EXT_COUNT_COLUMN = 12
OUTPUT_EXT_DNS_COLUMN = 13

LEGACY_SHEET_NAME = "Legacy VIP Candidates"

# ============================================
# HELPERS
# ============================================

def clean_text(val):
    if val is None:
        return ""
    return str(val).strip().strip('"').strip()

def clean_fqdn(val):
    return clean_text(val).rstrip(".").lower()

def clean_ip(val):
    return clean_text(val)

def is_ipv4(value):
    value = clean_ip(value)
    parts = value.split(".")
    if len(parts) != 4:
        return False
    try:
        return all(0 <= int(p) <= 255 for p in parts)
    except ValueError:
        return False

def build_fqdn(hostname, zone):
    hostname = clean_text(hostname)
    zone = clean_text(zone)

    if hostname == "@":
        return clean_fqdn(zone)

    if hostname and zone:
        return clean_fqdn(f"{hostname}.{zone}")

    if hostname:
        return clean_fqdn(hostname)

    if zone:
        return clean_fqdn(zone)

    return ""

def ptr_zone_to_ipv4(hostname, zone):
    hostname = clean_text(hostname)
    zone = clean_text(zone).lower().rstrip(".")

    if not hostname or not zone.endswith(".in-addr.arpa"):
        return ""

    zone_prefix = zone[:-len(".in-addr.arpa")]
    zone_parts = zone_prefix.split(".")
    zone_parts.reverse()

    host_parts = hostname.split(".")
    host_parts.reverse()

    ip_parts = zone_parts + host_parts

    if len(ip_parts) != 4:
        return ""

    ip = ".".join(ip_parts)

    return ip if is_ipv4(ip) else ""

def recreate_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    return wb.create_sheet(title=sheet_name)

def autosize_columns(ws, min_width=12, max_width=80):
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(min_width, min(max_len + 2, max_width))

def print_source_diagnostics():
    print("Source workbook:", VIP_WORKBOOK)
    print("Output workbook:", OUTPUT_WORKBOOK)

    if not VIP_WORKBOOK.exists():
        raise FileNotFoundError(f"Source workbook not found: {VIP_WORKBOOK}")

    stat = VIP_WORKBOOK.stat()
    modified = datetime.fromtimestamp(stat.st_mtime)

    print("Source exists: True")
    print("Source size bytes:", stat.st_size)
    print("Source modified:", modified)

def print_sheet_preview(wb):
    print("Loaded sheets:", ", ".join(wb.sheetnames))

# ============================================
# LOAD DNS CSV FILES
# ============================================

def load_dns_records(folder):
    ip_to_dns = defaultdict(set)

    csv_files = sorted(folder.glob("*.csv"))
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in {folder}")

    for csv_file in csv_files:
        print("Reading", csv_file.name)

        with open(csv_file, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)

            for row in reader:
                rtype = clean_text(row.get("RecordType", "")).upper()
                hostname = row.get("HostName", "")
                zone = row.get("Zone", "")
                rdata = row.get("RecordData", "")

                if rtype == "A":
                    fqdn = build_fqdn(hostname, zone)
                    ip = clean_ip(rdata)

                    if fqdn and is_ipv4(ip):
                        ip_to_dns[ip].add(fqdn)

                elif rtype == "PTR":
                    fqdn = clean_fqdn(rdata)
                    ip = ptr_zone_to_ipv4(hostname, zone)

                    if fqdn and is_ipv4(ip):
                        ip_to_dns[ip].add(fqdn)

    return ip_to_dns

# ============================================
# UPDATE WORKBOOK
# ============================================

def update_workbook(ip_to_dns):
    print("Opening workbook now...")
    wb = load_workbook(VIP_WORKBOOK)
    print_sheet_preview(wb)

    legacy_rows = []

    for ws in wb.worksheets:
        if ws.title in SKIP_SHEET_NAMES:
            print("Skipping sheet:", ws.title)
            continue

        print("Processing sheet:", ws.title)

        ws.cell(1, OUTPUT_INT_COUNT_COLUMN, "Matched DNS Count")
        ws.cell(1, OUTPUT_INT_DNS_COLUMN, "Matched DNS")
        ws.cell(1, OUTPUT_EXT_COUNT_COLUMN, "External DNS Count")
        ws.cell(1, OUTPUT_EXT_DNS_COLUMN, "External DNS")

        for row in range(START_ROW, ws.max_row + 1):
            if row % 500 == 0:
                print(f"  {ws.title}: row {row}/{ws.max_row}")

            internal_ip = clean_ip(ws.cell(row, INTERNAL_IP_COLUMN).value)
            external_ip = clean_ip(ws.cell(row, EXTERNAL_IP_COLUMN).value)

            internal_matches = []
            external_matches = []

            if is_ipv4(internal_ip):
                internal_matches = sorted(ip_to_dns.get(internal_ip, set()))

            if is_ipv4(external_ip):
                external_matches = sorted(ip_to_dns.get(external_ip, set()))

            ws.cell(row, OUTPUT_INT_COUNT_COLUMN, len(internal_matches))
            ws.cell(row, OUTPUT_INT_DNS_COLUMN, ", ".join(internal_matches))
            ws.cell(row, OUTPUT_EXT_COUNT_COLUMN, len(external_matches))
            ws.cell(row, OUTPUT_EXT_DNS_COLUMN, ", ".join(external_matches))

            if (
                (is_ipv4(internal_ip) or is_ipv4(external_ip))
                and len(internal_matches) == 0
                and len(external_matches) == 0
            ):
                legacy_rows.append([
                    ws.title,
                    row,
                    clean_text(ws.cell(row, 1).value),
                    external_ip,
                    internal_ip
                ])

    legacy_ws = recreate_sheet(wb, LEGACY_SHEET_NAME)

    legacy_ws.append([
        "Sheet",
        "Row",
        "VIP Name",
        "External IP",
        "Internal IP"
    ])

    for r in legacy_rows:
        legacy_ws.append(r)

    autosize_columns(legacy_ws)

    print("Saving workbook:", OUTPUT_WORKBOOK)
    wb.save(OUTPUT_WORKBOOK)

    print("Legacy VIP candidates:", len(legacy_rows))

# ============================================
# MAIN
# ============================================

print_source_diagnostics()

print("Loading DNS records...")
ip_to_dns = load_dns_records(DNS_FOLDER)
print("Unique IPs with DNS:", len(ip_to_dns))

update_workbook(ip_to_dns)

print("Done.")