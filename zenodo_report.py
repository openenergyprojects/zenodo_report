#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct 12 12:08:00 2023
"""
import openpyxl
import requests
import os
import re
import json

# Global parameters
OVERWRITE_OPEN_ACCESS_LINK = False  # Set to True to overwrite existing "Open Access link"
POPULATE_DOI_COLUMN = True  # Set to False to skip populating the DOI column
CURL_UA = "Mozilla/5.0 (Linux; Android 10; SM-G996U Build/QP1A.190711.020; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Mobile Safari/537.36"

def search_zenodo_by_title(title):
    """Search Zenodo for a record by title."""
    zenodo_api_url = "https://zenodo.org/api/records"
    params = {"q": f'title:"{title}"'}
    response = requests.get(zenodo_api_url, params=params)
    if response.status_code == 200:
        data = response.json()
        if data.get("hits", {}).get("total", 0) == 1:
            recid = data["hits"]["hits"][0]["id"]
            return recid, None
        elif data.get("hits", {}).get("total", 0) > 1:
            return None, "Multiple matches found"
    return None, "No match found"

def search_zenodo_by_doi(doi):
    """Search Zenodo for a record by DOI."""
    zenodo_api_url = "https://zenodo.org/api/records"
    params = {"q": f'doi:{doi}'}
    response = requests.get(zenodo_api_url, params=params)
    if response.status_code == 200:
        data = response.json()
        if data.get("hits", {}).get("total", 0) > 0:
            recid = data["hits"]["hits"][0]["id"]
            return recid
    return None

def extract_doi_from_bibliographic_data(bibliographic_data):
    """Extract DOI from the bibliographic data field."""
    match = re.search(r"DOI\s*:\s*([\S]+)", bibliographic_data)
    if match:
        return match.group(1).replace(" ", "")
    return None

def update_xlsx_with_zenodo_links(xlsx_file, sheet_names):
    """Update the XLSX file with Zenodo open access links, DOIs, and author emails for multiple sheets."""
    try:
        workbook = openpyxl.load_workbook(xlsx_file)
    except Exception as e:
        print(f"Error loading workbook: {e}")
        return

    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            print(f"Error: Sheet '{sheet_name}' not found in the XLSX file.")
            continue

        sheet = workbook[sheet_name]

        # Find headers and their column indices (headers are at row 3)
        headers = {cell.value: idx for idx, cell in enumerate(sheet[3], start=1)}
        if "NO." not in headers or "TITLE " not in headers or "BIBLIOGRAPHIC DATA" not in headers:
            print(f"Error: Required headers ('NO.', 'TITLE ', 'BIBLIOGRAPHIC DATA') not found in sheet '{sheet_name}'.")
            continue

        if "Open Access link" not in headers:
            headers["Open Access link"] = len(headers) + 1
            sheet.cell(row=3, column=headers["Open Access link"]).value = "Open Access link"

        if "DOI" not in headers:
            headers["DOI"] = 20  # Column T corresponds to index 20
            sheet.cell(row=3, column=headers["DOI"]).value = "DOI"

        if "link_as_text" not in headers:
            headers["link_as_text"] = 21  # Column U corresponds to index 21
            sheet.cell(row=3, column=headers["link_as_text"]).value = "link_as_text"

        if "author_email" not in headers:
            headers["author_email"] = 22  # Column V corresponds to index 22
            sheet.cell(row=3, column=headers["author_email"]).value = "author_email"

        for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row):
            no_cell = row[headers["NO."] - 1]
            title_cell = row[headers["TITLE "] - 1]
            bibliographic_data_cell = row[headers["BIBLIOGRAPHIC DATA"] - 1]
            open_access_link_cell = row[headers["Open Access link"] - 1]
            doi_cell = row[headers["DOI"] - 1]
            link_as_text_cell = row[headers["link_as_text"] - 1]
            author_email_cell = row[headers["author_email"] - 1]

            # Skip rows that do not start with a number or are not the title row
            if not str(no_cell.value).strip().isdigit() and str(no_cell.value).strip() != "NO.":
                continue

            title = str(title_cell.value).strip() if title_cell.value else ""
            bibliographic_data = str(bibliographic_data_cell.value).strip() if bibliographic_data_cell.value else ""
            doi = None

            if not title:
                print(f"  Skipping row with missing title in sheet '{sheet_name}': {row}")
                continue

            print(f"Processing entry in sheet '{sheet_name}': {title}")

            # Extract link_as_text from the "LINK" column if enabled
            if link_as_text_cell.value is None and "LINK" in headers:
                links_cell = row[headers["LINK"] - 1]
                if links_cell.hyperlink:
                    link_as_text_cell.value = links_cell.hyperlink.target
                else:
                    link_as_text_cell.value = links_cell.value if links_cell.value else ""

            # Check if "Open Access link" should be updated
            if not open_access_link_cell.value or OVERWRITE_OPEN_ACCESS_LINK:
                print(f"  Searching Zenodo for title: {title}")
                recid, message = search_zenodo_by_title(title)

                if recid:
                    zenodo_link = f"https://zenodo.org/records/{recid}"
                    print(f"  Found Zenodo record for title '{title}': {zenodo_link}")
                    open_access_link_cell.value = zenodo_link

                    # Attempt to extract DOI from bibliographic data if available
                    if bibliographic_data:
                        doi = extract_doi_from_bibliographic_data(bibliographic_data)

                    # Use DOI if found, otherwise fallback to recid
                    if doi:
                        print(f"  DOI extracted from bibliographic data: {doi}")
                        doi_cell.value = f"https://doi.org/{doi}"
                    else:
                        print(f"  No DOI found in bibliographic data, using recid as fallback.")
                        doi_cell.value = f"https://doi.org/{recid}"
                else:
                    print(f"  Zenodo search by title failed for '{title}': {message}")
                    if bibliographic_data:
                        doi = extract_doi_from_bibliographic_data(bibliographic_data)
                        if doi:
                            print(f"  Extracted DOI from bibliographic data: {doi}")
                            recid = search_zenodo_by_doi(doi)
                            if recid:
                                zenodo_link = f"https://zenodo.org/records/{recid}"
                                print(f"  Found Zenodo record for DOI '{doi}': {zenodo_link}")
                                open_access_link_cell.value = zenodo_link
                                doi_cell.value = f"https://doi.org/{doi}"
                            else:
                                print(f"  No Zenodo record found for DOI: {doi}. Possible reasons: DOI not indexed in Zenodo or incorrect DOI format.")
                                doi_cell.value = f"https://doi.org/{doi}"
                        else:
                            print(f"  No DOI found in bibliographic data for '{title}'. Possible reasons: Missing DOI in bibliographic data or incorrect format.")
                    else:
                        print(f"  No bibliographic data available for '{title}'. Possible reasons: Missing bibliographic data or incorrect format.")
            else:
                print(f"  Open Access link already populated for '{title}', skipping update.")
                # Check if DOI is already populated
                if doi_cell.value is None and bibliographic_data:
                    doi = extract_doi_from_bibliographic_data(bibliographic_data)
                    if doi:
                        print(f"  Extracted DOI from bibliographic data: {doi}")
                        doi_cell.value = f"https://doi.org/{doi}"
                    else:
                        print(f"  No DOI found in bibliographic data for '{title}'. Possible reasons: Missing DOI in bibliographic data or incorrect format.")
                else:
                    print(f"  DOI already populated for '{title}', skipping update.")

            # Populate author_email column if not already populated
            if not author_email_cell.value and link_as_text_cell.value:
                author_email = extract_ucy_author_email(link_as_text_cell.value)
                if author_email:
                    author_email_cell.value = author_email

            # Save the workbook after processing each row
            try:
                workbook.save(xlsx_file)
                print(f"Workbook saved successfully after processing row: {no_cell.value}")
            except Exception as e:
                print(f"Error saving workbook after processing row {no_cell.value}: {e}")

    # Save the updated workbook
    try:
        workbook.save(xlsx_file)
        print(f"Workbook saved successfully: {xlsx_file}")
    except Exception as e:
        print(f"Error saving workbook: {e}")

def extract_ucy_author_email(link):
    """Extract the email address of the author with ucy.ac.cy domain from the given link."""
    print(f"Debug: Starting email extraction for link: {link}")

    if "ieee.org" in link:
        print("Debug: Link contains 'ieee.org', proceeding to fetch metadata.")
        json = fetch_ieee_metadata(link)
        if not json:
            print("Debug: No JSON metadata fetched from IEEE link.")
            return None

        print(f"Debug: JSON metadata fetched: {json}")
        email = extract_email_from_ieee_json(json)
        if email:
            print(f"Debug: Email extracted from JSON metadata: {email}")
            return email

        print("Debug: No email found in JSON metadata, checking for specific author exceptions.")
        # Handle exceptions for specific authors
        if "Thomas Parisini" in json:
            print("Debug: Found 'Thomas Parisini' in JSON, returning exception email.")
            return "parisini.thomas@ucy.ac.cy"
        if "Alessandro Astolfi" in json:
            print("Debug: Found 'Alessandro Astolfi' in JSON, returning exception email.")
            return "astolfi.alessandro@ucy.ac.cy"

    print("Debug: Link does not contain 'ieee.org' or no email found, returning None.")
    return None

def fetch_ieee_metadata(link):
    """Fetch metadata from an IEEE link."""
    print(f"Debug: Fetching IEEE metadata for link: {link}")
    json = None
    try:
        response = requests.get(link, headers={"User-Agent": CURL_UA})
        print(f"Debug: HTTP response status code: {response.status_code}")
        if response.status_code == 200:
            json = extract_ieee_json(response.text)
            print(f"Debug: Extracted JSON from response: {json}")
        else:
            print(f"Debug: Failed to fetch metadata, status code: {response.status_code}")
    except Exception as e:
        print(f"Error fetching IEEE metadata: {e}")
    return json

def extract_ieee_json(html):
    """Extract JSON metadata from IEEE HTML."""
    print("Debug: Extracting JSON metadata from HTML.")
    # Updated regex to match the JSON metadata structure in the provided HTML sample
    match = re.search(r'xplGlobal\.document\.metadata\s*=\s*(\{.*?\})\s*;', html)
    if match:
        print("Debug: JSON metadata found in HTML.")
        return json.loads(match.group(1))
    print("Debug: No JSON metadata found in HTML.")
    return None

def extract_email_from_ieee_json(json):
    """Extract the email address of the first author with ucy.ac.cy domain."""
    print("Debug: Extracting email from JSON metadata.")
    for author in json.get("authors", []):
        print(f"Debug: Processing author: {author}")
        for affiliation in author.get("affiliation", []):
            print(f"Debug: Checking affiliation: {affiliation}")
            if "ucy.ac.cy" in affiliation:
                email = f"{author['lastName'].lower()}.{author['firstName'].lower()}@ucy.ac.cy"
                print(f"Debug: Found email: {email}")
                return email

    print("Debug: No email found with 'ucy.ac.cy' domain. Attempting to construct email from names.")
    # Attempt to construct email from names of authors with "University of Cyprus" in their affiliation
    for author in json.get("authors", []):
        for affiliation in author.get("affiliation", []):
            if "University of Cyprus" in affiliation:
                first_name = author['firstName'].split()[0].lower()
                last_name = author['lastName'].split()[0].lower()
                email = f"{last_name}.{first_name}@ucy.ac.cy"
                print(f"Debug: Constructed email: {email}")
                return email

    print("Debug: No authors with 'University of Cyprus' affiliation found. Checking for specific exceptions.")
    # Handle exceptions for specific authors
    for author in json.get("authors", []):
        if author['firstName'] == "Thomas" and author['lastName'] == "Parisini":
            print("Debug: Found 'Thomas Parisini', returning exception email.")
            return "parisini.thomas@ucy.ac.cy"
        if author['lastName'] == "Astolfi":
            print("Debug: Found 'Astolfi', returning exception email.")
            return "astolfi.alessandro@ucy.ac.cy"

    print("Debug: No solution found for email extraction.")
    return None

if __name__ == "__main__":
    xlsx_file_path = "OS Info -  KIOS PUBLICATIONS FOR 2017-2018-2019-2020-2021_2022_2023_2024.xlsx"
    # xlsx_file_path = "test.xlsx"
    sheet_names = ["YEAR 2024", "YEAR 2023", "YEAR 2022"]  # Example: Add multiple sheets to process
    if not os.path.exists(xlsx_file_path):
        print(f"XLSX file not found: {xlsx_file_path}")
    else:
        update_xlsx_with_zenodo_links(xlsx_file_path, sheet_names)

# TODO:
# add proper logging library to allow multiple log levels as params. Ensure the printed logs are easy to be read by human, so identation and eventually colours should be added.