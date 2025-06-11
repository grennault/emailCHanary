import subprocess
import sys
from typing import List
import re
import time
import tempfile
import os

def get_input_with_retry(prompt: str) -> str:
    """Get user input with retry on empty input."""
    while True:
        try:
            user_input = input(prompt).strip()
            if user_input:
                return user_input
            print("Input cannot be empty. Please try again.")
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            sys.exit(0)

def get_email_list(prompt: str) -> List[str]:
    """Get a list of email addresses from user input."""
    print(prompt)
    print("Enter one email per line. Press Enter twice when done.")
    
    emails = []
    while True:
        try:
            email = input().strip()
            if not email:
                break
            if re.match(r"[^@]+@[^@]+\.[^@]+", email):
                emails.append(email)
            else:
                print("Invalid email format. Please try again.")
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            sys.exit(0)
    
    return emails

def run_powershell_commands(commands: List[str], session=None) -> tuple[str, subprocess.Popen]:
    """
    Run multiple PowerShell commands in a single session.
    
    Args:
        commands: List of PowerShell commands to execute
        session: Not used anymore, kept for compatibility
    
    Returns:
        tuple: (command output, None)
    """
    try:
        # Always include the connection commands to ensure we're connected
        full_commands = [
            "Import-Module ExchangeOnlineManagement",
            "Connect-ExchangeOnline"
        ] + commands
        
        # Combine all commands with semicolons
        combined_command = "; ".join(full_commands)
        result = subprocess.run(
            ["powershell", "-Command", combined_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        output = result.stdout.strip()
        if output:
            print(output)
        
        return output, None
        
    except subprocess.CalledProcessError as e:
        print(f"Error executing PowerShell commands: {e}")
        print(f"Error output: {e.stderr}")
        raise
    except Exception as e:
        print(f"Error: {e}")
        raise

def get_available_domains(session=None) -> tuple[List[str], None]:
    """
    Get list of available domains from Exchange Online.
    
    Args:
        session: Not used anymore, kept for compatibility
    
    Returns:
        tuple: (list of domains, None)
    """
    commands = [
        "Get-AcceptedDomain | Select-Object -ExpandProperty DomainName"
    ]
    domains_output, _ = run_powershell_commands(commands)
    
    # Filter out non-domain lines
    domains = []
    for line in domains_output.split('\n'):
        line = line.strip()
        # Only include valid domain lines
        if (line and 
            '.' in line and 
            ' ' not in line and 
            not line.startswith('V') and 
            not line.startswith('U') and 
            not line.startswith('F') and 
            not line.startswith('S')):
            domains.append(line)
    
    return domains, None

def create_internal_canary(name: str, domain: str, forward_to: List[str], session=None) -> tuple[str, None]:
    """Create an internal canary shared mailbox that forwards to multiple addresses."""
    # Clean up the name for the email address and create alias
    clean_name = ''.join(c.lower() for c in name if c.isalnum())
    alias = f"{clean_name}alias"
    email = f"{clean_name}@{domain}"
    
    try:
        # First, create the shared mailbox
        print("\nCreating shared mailbox...")
        create_mailbox_commands = [
            f"""
            # Create the shared mailbox
            $params = @{{
                Shared = $true
                Name = '{name}'
                Alias = '{alias}'
                PrimarySmtpAddress = '{email}'
            }}
            New-Mailbox @params

            Write-Host "`nShared mailbox created:"
            Write-Host "Email: {email}"
            """
        ]
        run_powershell_commands(create_mailbox_commands)
        
        # Then, set up forwarding
        print("\nSetting up forwarding...")
        if len(forward_to) > 1:
            # Create forwarding rules for each address
            print(f"Setting up forwarding to multiple addresses...")
            for forward_address in forward_to:
                print(f"Adding forwarding rule for: {forward_address}")
                forward_commands = [
                    f"""
                    # Set up forwarding rule
                    Set-Mailbox -Identity '{email}' `
                        -ForwardingAddress '{forward_address}' `
                        -DeliverToMailboxAndForward $true `
                        -HiddenFromAddressListsEnabled $false `
                        -DisplayName '{name}' `
                        -WindowsEmailAddress '{email}'

                    Write-Host "`nForwarding rule added:"
                    Write-Host "From: {email}"
                    Write-Host "To: {forward_address}"
                    """
                ]
                run_powershell_commands(forward_commands)
        else:
            # Single forward address
            print(f"Setting up forwarding to: {forward_to[0]}")
            forward_commands = [
                f"""
                # Set up forwarding
                Set-Mailbox -Identity '{email}' `
                    -ForwardingAddress '{forward_to[0]}' `
                    -DeliverToMailboxAndForward $true `
                    -HiddenFromAddressListsEnabled $false `
                    -DisplayName '{name}' `
                    -WindowsEmailAddress '{email}'

                Write-Host "`nForwarding configured:"
                Write-Host "From: {email}"
                Write-Host "To: {forward_to[0]}"
                """
            ]
            run_powershell_commands(forward_commands)
        
        return email, None
        
    except Exception as e:
        print(f"\nError during canary creation: {str(e)}")
        print("Attempting to clean up...")
        try:
            # Try to remove the mailbox if it was created
            cleanup_commands = [
                f"""
                Remove-Mailbox -Identity '{email}' -Confirm:$false
                Write-Host "Cleaned up mailbox: {email}"
                """
            ]
            run_powershell_commands(cleanup_commands)
        except:
            print("Cleanup failed. You may need to manually remove the mailbox.")
        raise

def setup_external_canary(name: str, external_emails: List[str], session=None) -> tuple[List[str], None]:
    """Create external canary contacts in GAL."""
    created_contacts = []
    
    for external_email in external_emails:
        commands = [
            f"""
            # Create external contact
            New-MailContact -Name '{name}' -ExternalEmailAddress '{external_email}'
            
            # Make it visible in GAL
            Set-MailContact -Identity '{external_email}' -HiddenFromAddressListsEnabled $false
            
            Write-Host "`nExternal canary contact created and configured:"
            Write-Host "Name: {name}"
            Write-Host "Email: {external_email}"
            Write-Host "Visible in GAL: Yes"
            """
        ]
        
        run_powershell_commands(commands)
        created_contacts.append(external_email)
    
    return created_contacts, None

def main():
    try:
        print("Email Canary Setup")
        print("-----------------")
        print("This script will create two types of canaries:")
        print("1. Internal canary: A shared mailbox that forwards to specified addresses")
        print("2. External canary: An external contact visible in GAL")
        print("\nThese canaries will help detect if an internal email address is compromised")
        print("and being used to send phishing emails to contacts.\n")
        
        # External Canary Setup (First)
        print("External Canary Setup")
        print("--------------------")
        while True:
            create_external = get_input_with_retry("\nWould you like to create an external canary? (yes/no): ").lower()
            if create_external in ['yes', 'no']:
                break
            print("Please answer 'yes' or 'no'")
        
        external_emails = []
        if create_external == 'yes':
            contact_name = get_input_with_retry("\nEnter the name for the external canary: ")
            external_emails = get_email_list("\nEnter external email addresses:")
            
            if not external_emails:
                print("No external email addresses provided. Skipping external canary creation.")
            else:
                print("\nCreating external canaries first...")
                external_canaries, _ = setup_external_canary(contact_name, external_emails)
                print(f"\nExternal canaries created:")
                for email in external_canaries:
                    print(f"- {email}")
        
        # Internal Canary Setup
        print("\nInternal Canary Setup")
        print("--------------------")
        name = get_input_with_retry("Enter the name for the internal canary (e.g., Security Canary): ")
        
        # Get available domains
        print("\nFetching available domains...")
        domains, _ = get_available_domains()
        
        if not domains:
            print("No domains found!")
            sys.exit(1)
            
        print("\nAvailable domains:")
        for i, domain in enumerate(domains, 1):
            print(f"{i}. {domain}")
            
        # Select domain
        if len(domains) == 1:
            selected_domain = domains[0]
            print(f"\nUsing domain: {selected_domain}")
        else:
            while True:
                try:
                    choice = int(get_input_with_retry("\nSelect a domain number: "))
                    if 1 <= choice <= len(domains):
                        selected_domain = domains[choice - 1]
                        break
                    else:
                        print("Invalid selection. Please try again.")
                except ValueError:
                    print("Please enter a valid number.")
        
        # Get forwarding addresses
        forward_to = get_email_list("\nEnter forwarding email addresses:")
        
        if not forward_to:
            print("No forwarding addresses provided. Exiting...")
            sys.exit(1)
        
        # Create the internal canary
        clean_name = ''.join(c.lower() for c in name if c.isalnum())
        print(f"\nCreating internal canary with email: {clean_name}@{selected_domain}")
        internal_email, _ = create_internal_canary(name, selected_domain, forward_to)
        
        print("\nCanary Setup Complete!")
        print(f"Internal canary: {internal_email}")
        if create_external == 'yes' and external_emails:
            print("\nExternal canaries:")
            for email in external_canaries:
                print(f"- {email}")
        print("\nThese canaries will help detect if an internal email address is compromised")
        print("and being used to send phishing emails to contacts.")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 