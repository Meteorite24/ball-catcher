import subprocess
import re
from openpyxl import Workbook

def ping(host, count=4):
    """
    Ping a host and return the packet loss percentage and statistics.

    :param host: Host to ping (IP address or domain name)
    :param count: Number of ping requests to send
    :return: A dictionary with packet loss rate and statistics
    """
    try:
        # Execute the system 'ping' command
        output = subprocess.run(
            ["ping", "-c", str(count), host],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        # Parse output using regex to find packet loss
        packet_loss_regex = r"(\d+)% packet loss"
        packet_loss = re.search(packet_loss_regex, output.stdout)
        if packet_loss:
            packet_loss_percentage = int(packet_loss.group(1))
        else:
            packet_loss_percentage = None

        # Parse output for statistics (min/avg/max/mdev)
        stats_regex = r"rtt min/avg/max/mdev = ([\d.]+)/([\d.]+)/([\d.]+)/([\d.]+)"
        stats_match = re.search(stats_regex, output.stdout)

        stats = {
            "min": float(stats_match.group(1)) if stats_match else None,
            "avg": float(stats_match.group(2)) if stats_match else None,
            "max": float(stats_match.group(3)) if stats_match else None,
            "mdev": float(stats_match.group(4)) if stats_match else None,
        }

        return {
            "packet_loss_percentage": packet_loss_percentage,
            "statistics": stats,
        }
    except Exception as e:
        print(f"Error occurred while pinging: {e}")
        return None

def export_to_excel(results, output_file="packet_loss_results.xlsx"):
    """
    Export ping results to an Excel file.

    :param results: List of dictionaries containing ping results
    :param output_file: Name of the output Excel file
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Ping Results"

    # Create headers
    headers = ["Host", "Packet Loss (%)", "Min RTT (ms)", "Avg RTT (ms)", "Max RTT (ms)", "MDev RTT (ms)"]
    ws.append(headers)

    # Add data rows
    for result in results:
        ws.append([
            result["host"],
            result["packet_loss_percentage"],
            result["statistics"]["min"],
            result["statistics"]["avg"],
            result["statistics"]["max"],
            result["statistics"]["mdev"],
        ])

    # Save to Excel file
    wb.save(output_file)
    print(f"Results exported successfully to {output_file}")

if __name__ == "__main__":
    test_hosts = ["8.8.8.8", "cloudflare.com"]  # Define the hosts to test
    results = []

    for host in test_hosts:
        print(f"Pinging {host}...")
        result = ping(host)
        if result:
            print(f"Packet loss: {result['packet_loss_percentage']}%")
            print(f"Statistics: {result['statistics']}")
            result["host"] = host
            results.append(result)
        else:
            print(f"Failed to ping {host}")

    if results:
        export_to_excel(results)
