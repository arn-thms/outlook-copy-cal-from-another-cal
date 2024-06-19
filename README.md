# outlook-copy-cal-from-another-cal
## VBA Code to Copy Meeting Requests Based on Specific Email Address

This VBA script is designed for Microsoft Outlook and should be saved in `ThisOutlookSession`. The script operates offline and performs the following tasks:

1. Monitors incoming meeting requests from all accounts.
2. Checks if a specific email address is present in the `To` field (including recursively checking distribution lists).
3. If the specified email address is found, the meeting request is copied to another account.

### Prerequisites

- Microsoft Outlook
- VBA enabled in Outlook (Developer Mode)

### Installation

1. Open Microsoft Outlook.
2. Press `ALT + F11` to open the VBA editor.
3. In the Project Explorer, locate `ThisOutlookSession`.
4. Copy and paste the provided VBA code into `ThisOutlookSession`.

### Usage

- Ensure the script is saved and Outlook is running in offline mode.
- The script will automatically monitor and copy relevant meeting requests based on the specified conditions.
