import subprocess
import sys

def check_exchange_connection():
    """Check if we're connected to Exchange Online."""
    try:
        result = subprocess.run(
            ["powershell", "-Command", "Get-ConnectionInformation"],
            capture_output=True,
            text=True
        )
        if "Connected" not in result.stdout:
            print("Not connected to Exchange Online. Connecting...")
            subprocess.run(
                ["powershell", "-Command", "Connect-ExchangeOnline"],
                check=True
            )
            print("Connected to Exchange Online successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    check_exchange_connection() 