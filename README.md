# token-sheet
A Google Apps Script that generates row UUIDs for any Column A with "UUID" as the header.  Accepts values to corresponding rows based on their column headers.

## Shared Secret

Requires a long, strong, and complex shared secret stored as a script property named **SECRET_TOKEN**.

## Setup

Attach the scripts to a sheet, create a **SECRET_TOKEN** script property, and run the createOnChangeTrigger function once.

## Example

Command updates columns with headers "First Name" and "Last Name" if they exist in a row with `<ROW_UUID>`.

```
curl -L -H "Content-Type: application/json" \
  -d '{
    "authToken": "<SECRET TOKEN>",
    "UUID": "<ROW UUID>",
    "First Name": "Joey",
    "Last Name": "JoeJoe"
  }' \
  '<ENDPOINT URL>'
```
