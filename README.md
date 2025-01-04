# Jira DC-to-Cloud Automations Migration

This script automates the process of transforming Jira Server/DC automation rules into a format compatible with Jira Cloud. It also generates a spreadsheet (`mapping_result.xlsx`) that shows the mapping between server IDs and cloud IDs for statuses, priorities, projects, etc.

## Features

- **Removes Inactive Rules**: Filters out rules marked as `DISABLED`
- **Replaces Fixed Custom Fields**: Changes references like `"type": "ID", "value": "customfield_10000"` to the appropriate **Cloud** ID, or uses the field name if needed
- **Maps**: Status, Priority, Issue Type, Project, Users (`JIRAUSER` references), etc.
- **Generates a `mapping_result.xlsx`** file listing all mappings. Missing items in the Cloud are marked in **red** (with "YES" in the "Missing?" column)

## Requirements

- Python 3.7+ (recommended)
- Dependencies listed in `requirements.txt`:
  ```
  requests==2.31.0
  openpyxl==3.1.1
  tqdm==4.66.1
  inquirer==2.10.1
  jsonpath-ng==1.5.3
  ```

Install them with:

```bash
pip install -r requirements.txt
```

## How It Works

1. **mapping.xlsx**:
   - Contains sheets like `users`, `customFields`, `status`, `priority`, `issuetype`, and `projects`
   - Each sheet must have **server/DC IDs** in the first column (e.g., `A`) and **names** in the second column (e.g., `B`)

2. **Running the Script**:
   - Place this script and `mapping.xlsx` in the same folder
   - Have your **JSON** file(s) (automation export from Jira Server/DC) in the same folder as well
   - Run in the terminal:
     ```bash
     python app.py
     ```
     Replace `app.py` with the actual name of your script

3. **CLI Prompts**:
   - The script will list all `.json` files in the directory and ask you to pick one
   - It will ask if you want to split the final JSON into multiple files (one file per rule)
   - Once you confirm, it starts the migration process

4. **Output Files**:
   - `<yourfile>-original-pretty.json`: A prettified version of your original JSON (for backup/reference)
   - `<yourfile>-modified-for-cloud.json`: Intermediate JSON with transformations applied
   - `<yourfile>-modified-for-cloud-pretty.json`: Final, prettified JSON for Jira Cloud
   - `mapping_result.xlsx`: A new spreadsheet summarizing all the mappings (server ID -> cloud ID). Missing items in Cloud are highlighted in red
   - If you chose to split rules, additional `.json` files are created, one per rule

5. **No Logs File**:
   - By default, the script only outputs logs/progress bars to the console (no file logs)
   - If you need a file log, you can add a `FileHandler` in the logging section of the script

## Example Flow

1. **Export** automations from Jira Server/DC to JSON

2. **Ensure** your `mapping.xlsx` is filled with the correct server/DC IDs in column A and corresponding names/keys in column B for each object type

3. **Run** the script:
   ```bash
   python main.py
   ```

4. **Choose** your JSON file when prompted

5. **Check** the console output to see the progress bars for each replacement (status, priority, etc.)

6. **Open** the final `<yourfile>-modified-for-cloud-pretty.json` to see your updated rules, now referencing Jira Cloud IDs

7. **Check** `mapping_result.xlsx` for a summary of which objects were successfully mapped or missing in the Cloud
