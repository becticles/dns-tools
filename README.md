# F5 VIP DNS Reconciliation Tool

This tool reconciles **F5 load balancer VIP inventories** with exported **DNS records** to produce a consolidated mapping of hostnames to VIPs.

The script enriches a VIP spreadsheet by identifying DNS records associated with both internal and external IP addresses and highlights potential **legacy or unused VIP configurations** that no longer have DNS references.

The output is an updated spreadsheet that provides a clearer view of DNS relationships to load balancer VIPs.

---

## Purpose

In large environments, load balancer configurations and DNS records can drift over time. Applications are retired, migrated, or renamed, leaving behind VIP entries that no longer serve active traffic.

Manually identifying these discrepancies is time-consuming.

This tool automates the process by:

- Mapping DNS records to VIPs using internal and external IPs
- Identifying VIPs with no DNS references
- Producing a clean report for infrastructure review

---

## Features

- Parses multiple DNS export CSV files
- Supports mixed DNS record formats
- Supports both **A** and **PTR** records
- Maps DNS hostnames to VIPs using:
  - internal IP address
  - external IP address
- Adds DNS counts and hostnames to the VIP spreadsheet
- Generates a report identifying VIPs with no DNS references
- Produces timestamped output files
- Does **not modify the source workbook**

---

## Requirements

Python 3.9+

Install dependencies:

```bash
pip install -r requirements.txt
