"""
list_installed_software.py
Lists all installed software on a Windows system by querying the registry.
Outputs results to the console and optionally saves to a CSV file.
"""

import winreg
import csv
import sys
from datetime import datetime


REGISTRY_PATHS = [
    # 64-bit applications
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    # 32-bit applications on 64-bit Windows
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    # Per-user installations
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
]


def get_installed_software():
    """Query the Windows registry and return a list of installed software."""
    software_list = []
    seen = set()

    for hive, path in REGISTRY_PATHS:
        try:
            registry_key = winreg.OpenKey(hive, path)
        except FileNotFoundError:
            continue

        num_subkeys, _, _ = winreg.QueryInfoKey(registry_key)

        for i in range(num_subkeys):
            try:
                subkey_name = winreg.EnumKey(registry_key, i)
                subkey = winreg.OpenKey(registry_key, subkey_name)

                def get_value(name):
                    try:
                        value, _ = winreg.QueryValueEx(subkey, name)
                        return value.strip() if isinstance(value, str) else value
                    except FileNotFoundError:
                        return ""

                name = get_value("DisplayName")
                if not name or name in seen:
                    continue  # Skip unnamed or duplicate entries

                seen.add(name)
                software_list.append({
                    "Name":         name,
                    "Version":      get_value("DisplayVersion"),
                    "Publisher":    get_value("Publisher"),
                    "Install Date": format_date(get_value("InstallDate")),
                    "Install Location": get_value("InstallLocation"),
                    "Size (MB)":    format_size(get_value("EstimatedSize")),
                })

                winreg.CloseKey(subkey)

            except (OSError, EnvironmentError):
                continue

        winreg.CloseKey(registry_key)

    return sorted(software_list, key=lambda x: x["Name"].lower())


def format_date(raw: str) -> str:
    """Convert YYYYMMDD registry date to a readable format."""
    if len(raw) == 8 and raw.isdigit():
        try:
            return datetime.strptime(raw, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass
    return raw


def format_size(raw) -> str:
    """Convert KB value from registry to MB string."""
    try:
        mb = int(raw) / 1024
        return f"{mb:.1f}"
    except (ValueError, TypeError):
        return ""


def print_table(software_list):
    """Print the software list as a formatted table."""
    if not software_list:
        print("No installed software found.")
        return

    col_widths = {
        "Name": 45,
        "Version": 18,
        "Publisher": 35,
        "Install Date": 14,
        "Size (MB)": 10,
    }

    header = (
        f"{'Name':<{col_widths['Name']}}"
        f"{'Version':<{col_widths['Version']}}"
        f"{'Publisher':<{col_widths['Publisher']}}"
        f"{'Install Date':<{col_widths['Install Date']}}"
        f"{'Size (MB)':>{col_widths['Size (MB)']}}"
    )
    separator = "-" * len(header)

    print(f"\nInstalled Software ({len(software_list)} programs found)")
    print(separator)
    print(header)
    print(separator)

    for app in software_list:
        name      = app["Name"][:col_widths["Name"] - 1]
        version   = app["Version"][:col_widths["Version"] - 1]
        publisher = app["Publisher"][:col_widths["Publisher"] - 1]
        inst_date = app["Install Date"]
        size      = app["Size (MB)"]

        print(
            f"{name:<{col_widths['Name']}}"
            f"{version:<{col_widths['Version']}}"
            f"{publisher:<{col_widths['Publisher']}}"
            f"{inst_date:<{col_widths['Install Date']}}"
            f"{size:>{col_widths['Size (MB)']}}"
        )

    print(separator)


def save_to_csv(software_list, filename=None):
    """Save the software list to a CSV file."""
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"installed_software_{timestamp}.csv"

    fields = ["Name", "Version", "Publisher", "Install Date", "Install Location", "Size (MB)"]

    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(software_list)

    print(f"\nSaved {len(software_list)} records to: {filename}")
    return filename


def ask_date_filter() -> datetime | None:
    """
    Interactively prompt the user for an optional filter date.
    Returns a datetime object, or None if the user skips filtering.
    """
    print("\n" + "=" * 60)
    print("  DATE FILTER")
    print("=" * 60)
    print("You can filter results to show only apps installed AFTER")
    print("a specific date.  Accepted formats:")
    print("   YYYY-MM-DD   e.g.  2024-01-15")
    print("   DD/MM/YYYY   e.g.  15/01/2024")
    print("   MM/DD/YYYY   e.g.  01/15/2024")
    print("Press ENTER to skip filtering and show all apps.")
    print("-" * 60)

    formats = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]

    while True:
        raw = input("Enter filter date (or press ENTER to skip): ").strip()

        if not raw:
            print("No date filter applied — showing all installed software.\n")
            return None

        for fmt in formats:
            try:
                date = datetime.strptime(raw, fmt)
                print(f"Filter set: showing apps installed after {date.strftime('%Y-%m-%d')}\n")
                return date
            except ValueError:
                continue

        print(f"  Could not parse '{raw}'. Please use YYYY-MM-DD, DD/MM/YYYY, or MM/DD/YYYY.\n")


def filter_by_date(software_list: list, after: datetime) -> list:
    """Return only apps whose install date is strictly after `after`."""
    filtered = []
    no_date = []

    for app in software_list:
        raw_date = app["Install Date"]
        if not raw_date:
            no_date.append(app)
            continue
        try:
            install_dt = datetime.strptime(raw_date, "%Y-%m-%d")
            if install_dt > after:
                filtered.append(app)
        except ValueError:
            no_date.append(app)

    print(f"  {len(filtered)} app(s) installed after {after.strftime('%Y-%m-%d')}  "
          f"| {len(no_date)} app(s) had no install date recorded (excluded).")

    return filtered


def main():
    # Basic argument parsing (no external deps)
    save_csv = "--csv" in sys.argv
    csv_file = None
    for arg in sys.argv[1:]:
        if arg.startswith("--output="):
            csv_file = arg.split("=", 1)[1]

    print("Scanning installed software from Windows Registry...")
    software_list = get_installed_software()
    print(f"Found {len(software_list)} installed applications total.")

    # Ask user for optional date filter
    filter_date = ask_date_filter()

    if filter_date:
        software_list = filter_by_date(software_list, filter_date)

    print_table(software_list)

    if save_csv:
        save_to_csv(software_list, csv_file)
    else:
        print("\nTip: Run with --csv to export results to a CSV file.")
        print("     Use --output=filename.csv to specify a custom file name.")


if __name__ == "__main__":
    main()