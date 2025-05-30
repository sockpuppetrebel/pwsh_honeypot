#!/usr/bin/env python3
"""
Microsoft Graph UPN Lookup Script using Certificate Authentication
Searches for users by first and last name and retrieves their UPNs
"""

import os
import sys
import json
from typing import List, Dict, Optional
import requests
from msal import ConfidentialClientApplication
from cryptography import x509
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend

# Microsoft Graph API endpoint
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# Your Azure AD details
TENANT_ID = "3ec00d79-021a-42d4-aac8-dcb35973dff2"
CLIENT_ID = "fe2a9efe-3000-4b02-96ea-344a2583dd52"

def load_certificate(cert_path: str, cert_password: Optional[str] = None, key_path: Optional[str] = None) -> Dict:
    """
    Load certificate from file (PFX/PEM format)
    """
    # If key_path is provided, we're dealing with separate cert and key files
    if key_path:
        with open(cert_path, 'rb') as f:
            cert_data = f.read()
        with open(key_path, 'rb') as f:
            key_data = f.read()
        
        # Load certificate
        certificate = x509.load_pem_x509_certificate(cert_data)
        
        # Load private key
        private_key = serialization.load_pem_private_key(
            key_data, password=cert_password.encode() if cert_password else None
        )
        
        # Convert to PEM format for MSAL
        private_key_pem = private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption()
        ).decode('utf-8')
        
        cert_pem = certificate.public_bytes(
            encoding=serialization.Encoding.PEM
        ).decode('utf-8')
        
        # Get thumbprint (use SHA1 for Azure)
        from cryptography.hazmat.primitives import hashes
        thumbprint = certificate.fingerprint(hashes.SHA1()).hex()
        
        return {
            "private_key": private_key_pem,
            "thumbprint": thumbprint,
            "public_certificate": cert_pem
        }
    
    # Original logic for single file
    with open(cert_path, 'rb') as f:
        cert_data = f.read()
    
    # Try to load as PFX first
    try:
        from cryptography.hazmat.primitives.serialization import pkcs12
        if cert_password:
            password = cert_password.encode()
        else:
            password = None
        
        private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
            cert_data, password, backend=default_backend()
        )
        
        # Convert to PEM format for MSAL
        private_key_pem = private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption()
        ).decode('utf-8')
        
        cert_pem = certificate.public_bytes(
            encoding=serialization.Encoding.PEM
        ).decode('utf-8')
        
        # Get thumbprint (use SHA1 for Azure)
        from cryptography.hazmat.primitives import hashes
        thumbprint = certificate.fingerprint(hashes.SHA1()).hex()
        
        return {
            "private_key": private_key_pem,
            "thumbprint": thumbprint,
            "public_certificate": cert_pem
        }
    except:
        # Try to load as PEM
        try:
            # Assume the file contains both private key and certificate
            private_key = serialization.load_pem_private_key(
                cert_data, password=cert_password.encode() if cert_password else None
            )
            
            # Extract certificate
            cert_start = cert_data.find(b'-----BEGIN CERTIFICATE-----')
            cert_end = cert_data.find(b'-----END CERTIFICATE-----') + len(b'-----END CERTIFICATE-----')
            cert_pem_data = cert_data[cert_start:cert_end]
            
            certificate = x509.load_pem_x509_certificate(cert_pem_data)
            
            private_key_pem = private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            ).decode('utf-8')
            
            cert_pem = certificate.public_bytes(
                encoding=serialization.Encoding.PEM
            ).decode('utf-8')
            
            thumbprint = certificate.fingerprint(
                certificate.signature_hash_algorithm
            ).hex()
            
            return {
                "private_key": private_key_pem,
                "thumbprint": thumbprint,
                "public_certificate": cert_pem
            }
        except Exception as e:
            print(f"Error loading certificate: {e}")
            raise

def get_access_token_with_certificate(cert_data: Dict) -> Optional[str]:
    """
    Get access token using certificate authentication
    """
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential={
            "private_key": cert_data["private_key"],
            "thumbprint": cert_data["thumbprint"],
            "public_certificate": cert_data["public_certificate"]
        }
    )
    
    # Get token for Microsoft Graph
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in result:
        return result["access_token"]
    else:
        print(f"Error obtaining token: {result.get('error')}")
        print(f"Error description: {result.get('error_description')}")
        return None

def search_user_by_name(access_token: str, first_name: str, last_name: str) -> List[Dict]:
    """
    Search for users by first and last name using Microsoft Graph API
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Build search filter
    filter_query = f"givenName eq '{first_name}' and surname eq '{last_name}'"
    
    # Select only the fields we need
    select_fields = "userPrincipalName,displayName,givenName,surname,mail,id"
    
    url = f"{GRAPH_API_BASE}/users?$filter={filter_query}&$select={select_fields}"
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        return data.get("value", [])
    else:
        print(f"Error searching for {first_name} {last_name}: {response.status_code}")
        print(f"Response: {response.text}")
        return []

def process_user_list(access_token: str, users_data: str) -> Dict[str, List[Dict]]:
    """
    Process the user list and search for each user
    """
    results = {}
    
    # Parse the user data
    lines = users_data.strip().split('\n')
    
    for line in lines:
        # Clean and split the line
        parts = [p.strip() for p in line.split('|') if p.strip()]
        
        if len(parts) >= 2:
            first_name = parts[0]
            last_name = parts[1]
            
            print(f"Searching for: {first_name} {last_name}")
            
            found_users = search_user_by_name(access_token, first_name, last_name)
            
            key = f"{first_name} {last_name}"
            results[key] = found_users
            
            if found_users:
                for user in found_users:
                    print(f"  Found: {user.get('userPrincipalName', 'N/A')}")
            else:
                print(f"  No users found")
    
    return results

def main():
    # User list from the request
    users_data = """
    |Jack |McClean | 
    |Mike |Cartwright | 
    |Paul |Gray | 
    |William |LaPuma | 
    |Mark |Ryan | 
    |Matthew |Payne | 
    |David |Bingham | 
    |Anna |Parback | 
    |Jack |Joseph | 
    |Thomas |McKenzie | 
    |Sean |Groat | 
    |Brandon |Halvorson | 
    |Daniel |Martell | 
    |Jon |Jones | 
    |Rob |Stoves | 
    |Nuno |Figueiredo | 
    |Marcus |Hoffman | 
    |Rob |Saunders | 
    |Zachary |Coulter | 
    |Phil |Yates | 
    |Shannon |Gray | 
    |Anatoliy |Savinov | 
    |Brett |Samuels | 
    |Anna |Redmile | 
    |Aidan |Dodd | 
    |Vimi |Kaul | 
    |Mark |Wakelin | 
    |Chynna |Roberts | 
    |Tarik |Antunes | 
    |Alexandra |Van Heel | 
    |Jennifer |Lovett | 
    |Terry |McGregor | 
    |Robin |LeClerc |
    """
    
    # Get certificate path from command line argument or environment variable
    if len(sys.argv) > 1:
        cert_path = sys.argv[1]
    else:
        cert_path = os.getenv("AZURE_CERT_PATH")
        if not cert_path:
            cert_path = input("Enter the path to your certificate file (PFX or PEM): ")
    
    key_path = None
    cert_password = None
    
    # Check if separate key file exists (for PEM certificates)
    if cert_path.endswith('_cert.pem'):
        potential_key_path = cert_path.replace('_cert.pem', '_key.pem')
        if os.path.exists(potential_key_path):
            key_path = potential_key_path
            print(f"  Key file: {key_path}")
    
    # Check if it's a PFX file (which usually has a password)
    if cert_path.lower().endswith('.pfx') or cert_path.lower().endswith('.p12'):
        if len(sys.argv) > 2:
            cert_password = sys.argv[2]
        else:
            cert_password = os.getenv("AZURE_CERT_PASSWORD")
            if cert_password is None and sys.stdin.isatty():
                cert_password = input("Enter certificate password (press Enter if none): ")
                if not cert_password:
                    cert_password = None
    
    print(f"\nUsing:")
    print(f"  Tenant ID: {TENANT_ID}")
    print(f"  Client ID: {CLIENT_ID}")
    print(f"  Certificate: {cert_path}")
    
    try:
        # Load certificate
        print("\nLoading certificate...")
        cert_data = load_certificate(cert_path, cert_password, key_path)
        print(f"Certificate loaded successfully (Thumbprint: {cert_data['thumbprint']})")
        
        # Get access token
        print("\nAuthenticating with Microsoft Graph...")
        access_token = get_access_token_with_certificate(cert_data)
        
        if not access_token:
            print("Failed to obtain access token")
            return
        
        print("Authentication successful!")
        
        print("\nSearching for users...\n")
        
        # Process the user list
        results = process_user_list(access_token, users_data)
        
        # Output results in a formatted way
        print("\n" + "="*80)
        print("SUMMARY OF RESULTS")
        print("="*80)
        
        successful_lookups = []
        failed_lookups = []
        multiple_matches = []
        
        for name, users in results.items():
            if len(users) == 0:
                failed_lookups.append(name)
            elif len(users) == 1:
                successful_lookups.append({
                    "name": name,
                    "upn": users[0].get("userPrincipalName", "N/A"),
                    "email": users[0].get("mail", "N/A"),
                    "id": users[0].get("id", "N/A")
                })
            else:
                multiple_matches.append({
                    "name": name,
                    "count": len(users),
                    "users": users
                })
        
        # Output successful lookups
        print(f"\nSuccessful Lookups ({len(successful_lookups)}):")
        print("-" * 80)
        for item in successful_lookups:
            print(f"{item['name']:<30} -> {item['upn']}")
        
        # Output multiple matches
        if multiple_matches:
            print(f"\nMultiple Matches Found ({len(multiple_matches)}):")
            print("-" * 80)
            for item in multiple_matches:
                print(f"\n{item['name']} ({item['count']} matches):")
                for user in item['users']:
                    print(f"  - {user.get('userPrincipalName', 'N/A')} ({user.get('displayName', 'N/A')})")
        
        # Output failed lookups
        if failed_lookups:
            print(f"\nFailed Lookups ({len(failed_lookups)}):")
            print("-" * 80)
            for name in failed_lookups:
                print(f"  - {name}")
        
        # Export to CSV
        print("\nExporting results to CSV...")
        with open("upn_lookup_results.csv", "w") as f:
            f.write("First Name,Last Name,UPN,Email,User ID,Status\n")
            
            for item in successful_lookups:
                name_parts = item['name'].split(' ', 1)
                f.write(f"{name_parts[0]},{name_parts[1]},{item['upn']},{item['email']},{item['id']},Found\n")
            
            for item in multiple_matches:
                name_parts = item['name'].split(' ', 1)
                for user in item['users']:
                    f.write(f"{name_parts[0]},{name_parts[1]},{user.get('userPrincipalName', 'N/A')},{user.get('mail', 'N/A')},{user.get('id', 'N/A')},Multiple\n")
            
            for name in failed_lookups:
                name_parts = name.split(' ', 1)
                f.write(f"{name_parts[0]},{name_parts[1]},,,,Not Found\n")
        
        print("Results exported to: upn_lookup_results.csv")
        
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()