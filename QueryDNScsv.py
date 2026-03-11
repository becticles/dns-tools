import os
import csv
import glob
import getpass

def get_csv_files():
    user = getpass.getuser()
    base_path = f"C:\\Users\\{user}\\github\\msinfra-staticdns\\"
    return glob.glob(os.path.join(base_path, '*.csv'))

def parse_csv_files(csv_files):
    dns_records = []
    for file in csv_files:
        with open(file, mode='r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 5:
                    hostname = row[0].strip().strip('"')
                    domain = row[1].strip().strip('"')
                    ip = row[4].strip().strip('"')
                    fqdn = f"{hostname}.{domain}".lower()
                    dns_records.append({"fqdn": fqdn, "ip": ip})
    return dns_records

def query_dns(records, query):
    query = query.lower()
    matches = []
    for record in records:
        if query == record["fqdn"] or query == record["ip"]:
            matches.append(record)
    return matches

def main():
    query = input("Enter FQDN or IP to search: ").strip()
    csv_files = get_csv_files()
    if not csv_files:
        print("No CSV files found.")
        return

    records = parse_csv_files(csv_files)
    results = query_dns(records, query)

    if results:
        # Sort alphabetically by FQDN
        results = sorted(results, key=lambda r: r["fqdn"])
        for r in results:
            print(f"FQDN: {r['fqdn']}, IP: {r['ip']}")
    else:
        print("No matching record found.")

if __name__ == "__main__":
    main()

