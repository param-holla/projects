# Apex Home Lending data model as deployable metadata — Project 03 step 0.3

The model lives in **`Apex-Data-Model.xlsx`**. `datamodel.py` reads it and generates both these folders and the field tables in `03-salesforce-platform-security.md`, so the spreadsheet, the guide and the org can't drift apart. Deploy the four folders in order: roll-ups need the child objects, and the LTV formula needs the roll-ups.

| Order | Folder | What it creates | Permission set |
|---|---|---|---|
| 1 | `1-standard-object-fields` | Account `Core_Household_Id__c`; Contact `Core_Borrower_Id__c`, `Originator_Id__c`, `Customer_Since__c`, `SIN__c` (encrypted); 25 Opportunity fields | `Apex_Data_Model_1` |
| 2 | `2-custom-objects` | 7 custom objects (Approval Request, Property, Applicant Income, Liability, Condition, Loan, Loan Party) and 3 platform events (`Borrower_Address_Changed__e`, `Approved_Action__e`, `Renewal_Due__e`) | `Apex_Data_Model_2` |
| 3 | `3-rollups` | Opportunity `Purchase_Price__c` (MAX), `Appraised_Value__c` (MIN), `Outstanding_Conditions__c` (COUNT of Outstanding or Received) | `Apex_Data_Model_3` |
| 4 | `4-ltv-formula` | Opportunity `LTV__c` | `Apex_Data_Model_4` |

```bash
cd 1-standard-object-fields
sf project deploy start --source-dir force-app --target-org apex-dev --dry-run   # validate first
sf project deploy start --source-dir force-app --target-org apex-dev
sf org assign permset --name Apex_Data_Model_1 --target-org apex-dev
cd ../2-custom-objects   # then 3-rollups, then 4-ltv-formula: the same three commands, permission sets _2, _3, _4
```

If you've already created some fields by hand, deploying the same field again just updates it to this definition. Check the dry-run output first.

After deploying:

- **Field access:** deployed fields and objects are invisible until a permission set grants access, hence the four sets above. They're admin conveniences for the build; the real access model comes from `PS_Specialist_Core` and the permission set groups (step 0.4).
- **Page layouts and tabs:** metadata deploys don't place fields on layouts or create tabs. Add the fields to each layout, and create tabs for Loan, Approval Request, Property and Condition (Setup → Tabs → New).
- **Source control:** copy everything into your main `force-app` and commit it.
- **To change the model:** edit `Apex-Data-Model.xlsx` (sheets *Objects* and *Fields*; the *Types* sheet says which columns each field type needs), then:
  ```bash
  pip3 install openpyxl                 # once
  python3 datamodel.py check            # finds mistakes before Salesforce does
  python3 datamodel.py package out      # writes out/1-… to out/4-…
  python3 datamodel.py objects          # guide tables (also: guide, events)
  ```
  `check` catches the slips seen in a real build: `__c` in a label, `_c__c` API names, lengths over the limit, a picklist default that isn't a value, a roll-up pointing at a field that doesn't exist, duplicate child relationship names, and sharing that doesn't match a master-detail field. `package` refuses to run until `check` passes.

Still built by hand, as the guide describes: the `Mortgage Deal` sales process, record types, the support process, validation rules and page layouts.

## org-as-code/

Files for keeping the org as source and rebuilding it (Project 03 §0.3a):

- `manifest/package.xml` — everything Project 03 builds (objects with their fields, record types and processes; stage and Case status values; custom-object layouts; the four permission sets).
- `scripts/build-org.sh <alias> [--scratch]` — optionally creates a scratch org from the `apex-dev` Dev Hub, deploys the manifest, assigns the permission sets and imports `data/sample-plan.json` if present.
- `config/project-scratch-def.json` — Developer edition with `MultiCurrency` (the Developer Edition has it on).

Copy the three folders into your Salesforce project, retrieve with the manifest, commit, then run the script against a scratch org to prove the source is complete.
