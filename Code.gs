/**
 * Creates Domain Shared Contacts from a Google Spreadsheet using the People API.
 * Also includes a function to list existing domain contacts.
 *
 * !! --- CRITICAL SETUP REQUIRED --- !!
 *
 * This script requires you to associate it with a Google Cloud Platform (GCP) project
 * where you can enable the People API and obtain an OAuth 2.0 Client ID for
 * Domain-Wide Delegation.
 *
 * --- DETAILED SETUP STEPS ---
 *
 * 1. PREPARE YOUR SPREADSHEET:
 * - Ensure your Google Spreadsheet has contact data with a header row (first row).
 * - Update `SPREADSHEET_ID` and `SHEET_NAME` variables below in the "CONFIGURATION" section.
 * - Adjust `COLUMN_MAPPINGS` in the "CONFIGURATION" section to match your spreadsheet's column headers.
 *
 * 2. CREATE OR SELECT A GOOGLE CLOUD PLATFORM (GCP) PROJECT:
 * - Go to the Google Cloud Console: https://console.cloud.google.com/
 * - If you don't have a suitable GCP project:
 * - Click the project selector (top of the page) > "NEW PROJECT".
 * - Name it (e.g., "Workspace-Domain-Contacts-Script"), select organization/location if needed, and click "CREATE".
 * - Note the **Project Number** or **Project ID** of this GCP project.
 *
 * 3. ASSOCIATE APPS SCRIPT WITH YOUR GCP PROJECT:
 * - In the Apps Script editor for this script:
 * - Click on "Project Settings" (the gear icon ⚙️ on the left sidebar).
 * - Under "Google Cloud Platform (GCP) Project", click "Change project".
 * - Enter the **Project Number** of the GCP project from Step 2.
 * - Click "Set project".
 *
 * 4. CONFIGURE YOUR GCP PROJECT:
 * (Perform these steps in the Google Cloud Console, ensuring your selected GCP project (from Step 2) is active)
 *
 * A. ENABLE PEOPLE API:
 * - Navigation menu (☰) > APIs & Services > Library.
 * - Search for "People API", select it, and click "Enable".
 *
 * B. CONFIGURE OAUTH CONSENT SCREEN:
 * - Navigation menu (☰) > APIs & Services > OAuth consent screen.
 * - **User Type:** Choose "Internal" (if available and script is for your organization only). Otherwise, "External".
 * - Fill in required fields: App name (e.g., "Domain Contacts Importer Script"), User support email, Developer contact.
 * - Click "SAVE AND CONTINUE" through Scopes and Test Users (default settings are usually fine here for this type of script).
 *
 * C. CREATE OAUTH 2.0 CLIENT ID:
 * - Navigation menu (☰) > APIs & Services > Credentials.
 * - Click "+ CREATE CREDENTIALS" > "OAuth client ID".
 * - **Application type:** Select "Web application".
 * - **Name:** Give it a name (e.g., "Apps Script Domain Contacts Client").
 * - **Authorized JavaScript origins:** You can leave this blank or add `https://script.google.com`.
 * - **Authorized redirect URIs:** Click "+ ADD URI" and enter:
 * `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback`
 * Replace `{SCRIPT_ID}` with your Apps Script project's Script ID. (Find this in Apps Script editor: Project Settings ⚙️ > Script ID).
 * - Click "CREATE".
 * - A dialog will show "Your Client ID". **COPY THIS CLIENT ID.** It ends with `.apps.googleusercontent.com`.
 * This is the ID needed for Domain-Wide Delegation.
 *
 * 5. GOOGLE WORKSPACE ADMIN: SET UP DOMAIN-WIDE DELEGATION:
 * A Google Workspace Super Administrator MUST perform these steps in the Admin Console (admin.google.com):
 * - Navigate to: Security > Access and data control > API Controls.
 * - Under "Domain Wide Delegation", click "MANAGE DOMAIN WIDE DELEGATION".
 * - Click "Add new".
 * - **Client ID field:** Paste the OAuth 2.0 Client ID obtained in Step 4.C.
 * - **OAuth Scopes (comma-delimited) field:** Enter EXACTLY:
 * `https://www.googleapis.com/auth/contacts,https://www.googleapis.com/auth/contacts.readonly`
 * (Adding contacts.readonly scope is good practice for listing functions)
 * - Click "Authorize".
 * (It might take some time for these settings to propagate fully).
 *
 * 6. ENABLE PEOPLE API SERVICE IN APPS SCRIPT:
 * - In the Apps Script editor, click on "Services" (+) in the left sidebar.
 * - Find "People API", click it, and then click "Add". (Identifier in code will be `People`).
 *
 * 7. RUN THE SCRIPT:
 * - Save all changes in the Apps Script editor.
 * - Reload the Spreadsheet if you have it open.
 * - A custom menu "Domain Contacts" should appear. Click it, then "Create Contacts from Sheet" or "List Domain Contacts".
 * - Or, in the Apps Script editor, select the desired function from the function dropdown and click "Run" (▶️).
 * - Authorize the script when prompted (if you haven't already or if scopes changed).
 *
 * --- SPREADSHEET STRUCTURE (for creating contacts) ---
 * - The first row of your sheet MUST be headers.
 * - Data types should be appropriate (e.g., valid email formats).
 */

// --- CONFIGURATION ---
const SPREADSHEET_ID = "Your-Key-Here"; // Replace this with your Spreadsheet ID!
const SHEET_NAME = "Contacts"; // Replace with your sheet name

const COLUMN_MAPPINGS = {
  givenName: "Given Name",
  familyName: "Family Name",
  emailWork: "Work Email",
  emailPersonal: "Personal Email",
  emailOther: "Other Email",
  phoneWork: "Work Phone",
  phoneMobile: "Mobile Phone",
  phoneHome: "Home Phone",
  company: "Company",
  jobTitle: "Job Title",
  notes: "Notes", // This key in COLUMN_MAPPINGS is fine, will be mapped to biographies
  website: "Website",
  addressStreetWork: "Work Street",
  addressCityWork: "Work City",
  addressStateWork: "Work State",
  addressZipWork: "Work ZIP",
  addressCountryWork: "Work Country"
};
// --- END CONFIGURATION ---

/**
 * Creates a custom menu in the spreadsheet to run the script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Domain Contacts')
    .addItem('Create Contacts from Sheet', 'createDomainContactsFromSpreadsheet')
    .addSeparator()
    .addItem('List Domain Contacts', 'listDomainContacts')
    .addToUi();
}

/**
 * Main function to process the spreadsheet and create domain shared contacts.
 */
function createDomainContactsFromSpreadsheet() {
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || !SPREADSHEET_ID) {
    const msg = "ERROR: SPREADSHEET_ID is not set. Please update the script configuration.";
    Logger.log(msg);
    SpreadsheetApp.getUi().alert("Configuration Error", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    if (typeof People === 'undefined' || !People.People || !People.People.createContact) {
      throw new Error("People API service not enabled or not available in Apps Script. Please check 'Services' in the editor.");
    }
  } catch (e) {
    Logger.log(e.message);
    SpreadsheetApp.getUi().alert("API Setup Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    const msg = `ERROR: Sheet "${SHEET_NAME}" not found in spreadsheet ID "${SPREADSHEET_ID}".`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert("Sheet Not Found", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) {
    Logger.log("No data found in the sheet (expected header row and at least one data row).");
    SpreadsheetApp.getUi().alert("No Data", "No data found in the sheet to process.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const headers = values[0].map(header => header.toString().trim().toLowerCase());
  const contactsData = values.slice(1);

  const headerToApiKey = {};
  for (const apiKey in COLUMN_MAPPINGS) {
    headerToApiKey[COLUMN_MAPPINGS[apiKey].toLowerCase()] = apiKey;
  }

  let contactsAdded = 0;
  let contactsFailed = 0;
  const ui = SpreadsheetApp.getUi();

  for (let i = 0; i < contactsData.length; i++) {
    const row = contactsData[i];
    const rowIndex = i + 2;
    const contactDetails = {};

    headers.forEach((header, colIndex) => {
      if (headerToApiKey[header]) {
        contactDetails[headerToApiKey[header]] = row[colIndex] ? row[colIndex].toString().trim() : "";
      }
    });

    if (!contactDetails.givenName && !contactDetails.familyName && !contactDetails.company) {
      Logger.log(`Skipping row ${rowIndex}: Insufficient data (missing Given Name, Family Name, or Company).`);
      contactsFailed++;
      continue;
    }

    const personResource = {
      names: [],
      emailAddresses: [],
      phoneNumbers: [],
      organizations: [],
      biographies: [], 
      urls: [],
      addresses: []
    };

    if (contactDetails.givenName || contactDetails.familyName) {
      personResource.names.push({
        givenName: contactDetails.givenName || "",
        familyName: contactDetails.familyName || ""
      });
    }

    // Add emails. The first one added will generally be considered primary by default.
    // Removed 'primary: true' flag.
    const addedEmails = new Set(); // To avoid duplicate email entries if columns overlap
    if (contactDetails.emailWork && !addedEmails.has(contactDetails.emailWork.toLowerCase())) {
        personResource.emailAddresses.push({ value: contactDetails.emailWork, type: "work"});
        addedEmails.add(contactDetails.emailWork.toLowerCase());
    }
    if (contactDetails.emailPersonal && !addedEmails.has(contactDetails.emailPersonal.toLowerCase())) {
        personResource.emailAddresses.push({ value: contactDetails.emailPersonal, type: "home" });
        addedEmails.add(contactDetails.emailPersonal.toLowerCase());
    }
    if (contactDetails.emailOther && !addedEmails.has(contactDetails.emailOther.toLowerCase())) {
         personResource.emailAddresses.push({ value: contactDetails.emailOther, type: "other" });
         addedEmails.add(contactDetails.emailOther.toLowerCase());
    }


    if (contactDetails.phoneWork) personResource.phoneNumbers.push({ value: contactDetails.phoneWork, type: "work" });
    if (contactDetails.phoneMobile) personResource.phoneNumbers.push({ value: contactDetails.phoneMobile, type: "mobile" });
    if (contactDetails.phoneHome) personResource.phoneNumbers.push({ value: contactDetails.phoneHome, type: "home" });

    if (contactDetails.company || contactDetails.jobTitle) {
      // For organizations, 'primary: true' might be valid or the first is primary.
      // Let's assume first is primary by order for simplicity, matching email behavior.
      personResource.organizations.push({
        name: contactDetails.company || "",
        title: contactDetails.jobTitle || "",
        type: "work"
        // primary: true // Removed for consistency, API usually takes first as primary
      });
    }
    if (personResource.names.length === 0 && contactDetails.company) {
        personResource.names.push({
            displayName: contactDetails.company
        });
    }

    if (contactDetails.notes) {
        personResource.biographies.push({
            value: contactDetails.notes,
            contentType: "TEXT_PLAIN"
        });
    }
    if (contactDetails.website) personResource.urls.push({ value: contactDetails.website, type: "website" });

    const workAddress = {};
    if (contactDetails.addressStreetWork) workAddress.streetAddress = contactDetails.addressStreetWork;
    if (contactDetails.addressCityWork) workAddress.city = contactDetails.addressCityWork;
    if (contactDetails.addressStateWork) workAddress.region = contactDetails.addressStateWork;
    if (contactDetails.addressZipWork) workAddress.postalCode = contactDetails.addressZipWork;
    if (contactDetails.addressCountryWork) workAddress.country = contactDetails.addressCountryWork;
    if (Object.keys(workAddress).length > 0) {
      workAddress.type = "work";
      personResource.addresses.push(workAddress);
    }

    for (const key in personResource) {
      if (Array.isArray(personResource[key]) && personResource[key].length === 0) {
        delete personResource[key];
      }
    }
    if (!personResource.names || personResource.names.length === 0) {
        Logger.log(`Skipping row ${rowIndex}: A name (given, family, or display via company) is required by People API.`);
        contactsFailed++;
        continue;
    }
    if (personResource.names && personResource.names.length > 0 && Object.keys(personResource).length <=1 ) {
        let hasSubstantiveName = false;
        personResource.names.forEach(name => {
            if((name.givenName && name.givenName.trim() !== "") || (name.familyName && name.familyName.trim() !== "") || (name.displayName && name.displayName.trim() !== "")) {
                hasSubstantiveName = true;
            }
        });
        if (!hasSubstantiveName) {
             Logger.log(`Skipping row ${rowIndex}: Name fields are present but empty or only whitespace.`);
             contactsFailed++;
             continue;
        }
    }

    try {
      const createdContact = People.People.createContact(personResource);
      Logger.log(`Successfully created contact for row ${rowIndex}: ${createdContact.resourceName} (Names: ${JSON.stringify(createdContact.names)})`);
      contactsAdded++;
    } catch (e) {
      Logger.log(`Error creating contact for row ${rowIndex}: ${e.message}. Details: ${JSON.stringify(contactDetails)}. Payload: ${JSON.stringify(personResource)}`);
      contactsFailed++;
    }
  }

  const summaryMessage = `Domain Contact Import Complete.\nAdded: ${contactsAdded}\nFailed/Skipped: ${contactsFailed}`;
  Logger.log(summaryMessage);
  ui.alert("Process Complete", summaryMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * Lists contacts from the domain directory using People.People.listDirectoryPeople().
 * Requires People API and relevant scopes (contacts.readonly or contacts).
 */
function listDomainContacts() {
  Logger.log("Attempting to list domain contacts using listDirectoryPeople...");
  try {
    if (typeof People === 'undefined' ||
        !People.People ||
        !People.People.listDirectoryPeople) { 
      throw new Error("People API service not enabled or listDirectoryPeople method not available. Please check 'Services' in the editor and ensure People API is correctly configured.");
    }

    const response = People.People.listDirectoryPeople({
      readMask: 'names,emailAddresses,phoneNumbers,biographies', 
      pageSize: 50,                                 
      sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_CONTACT', 'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE']
    });

    const people = response.people; 
    if (people && people.length > 0) {
      Logger.log(`Found ${people.length} domain contacts (showing up to ${50}):`);
      people.forEach((person, index) => { 
        let contactInfo = `Contact ${index + 1}:`;
        if (person.names && person.names.length > 0) {
          const nameObj = person.names[0];
          contactInfo += ` Name: ${nameObj.displayName || (nameObj.givenName + ' ' + nameObj.familyName).trim() || 'N/A'}`;
        } else {
          contactInfo += " Name: Not available";
        }
        if (person.emailAddresses && person.emailAddresses.length > 0) {
          contactInfo += `, Email: ${person.emailAddresses[0].value}`;
        }
        if (person.phoneNumbers && person.phoneNumbers.length > 0) {
          contactInfo += `, Phone: ${person.phoneNumbers[0].value}`;
        }
        if (person.biographies && person.biographies.length > 0) {
            contactInfo += `, Note: ${person.biographies[0].value}`; 
        }
        Logger.log(contactInfo);
      });
      SpreadsheetApp.getUi().alert("Domain Contacts Listed", `Found ${people.length} domain contacts. Details logged to Apps Script execution log. (View > Executions)`, SpreadsheetApp.getUi().ButtonSet.OK);

    } else {
      Logger.log("No domain contacts found or no contacts accessible via listDirectoryPeople with current permissions/setup.");
      SpreadsheetApp.getUi().alert("No Domain Contacts", "No domain contacts were found or could be accessed. Check logs for details.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
     if (response.nextPageToken) {
        Logger.log("There are more contacts. Implement pagination to see all of them.");
    }

  } catch (e) {
    Logger.log(`Error listing domain contacts: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`); 
    SpreadsheetApp.getUi().alert("Error Listing Contacts", `Failed to list domain contacts: ${e.message}. Check logs.`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Helper function to test getting headers and mappings for contact creation.
 */
function testSpreadsheetAndMappings() {
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || !SPREADSHEET_ID) {
    Logger.log("ERROR: SPREADSHEET_ID is not set.");
    return;
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Sheet "${SHEET_NAME}" not found.`);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("Detected Spreadsheet Headers: " + headers.map(h => h.toString().trim()).join(", "));

  const headerToApiKey = {};
  for (const apiKey in COLUMN_MAPPINGS) {
    headerToApiKey[COLUMN_MAPPINGS[apiKey].toLowerCase()] = apiKey;
  }
  Logger.log("Mapped Headers (Spreadsheet Header -> API Key):");
  headers.forEach(header => {
    const lcHeader = header.toString().trim().toLowerCase();
    if (headerToApiKey[lcHeader]) {
      Logger.log(`'${header}' -> '${headerToApiKey[lcHeader]}'`);
    } else {
      Logger.log(`'${header}' -> No mapping found (Consider adding to COLUMN_MAPPINGS if needed)`);
    }
  });
}

