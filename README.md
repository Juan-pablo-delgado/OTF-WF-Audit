# OTF-WF-Audit

## About this code
This code is designed to audit HubSpot workflows and provide users with recommendations based on workflow status, activity, and other factors. It is ideal for businesses using HubSpot to automate processes and seek efficient management and evaluation of these workflows. The code interacts with Excel data, performing transformations and analyses specifically tailored to HubSpot workflows. It includes functions for automating audit tasks, such as generating links, evaluating workflow status (active or inactive), and recommending actions based on various metrics.
## General workflow description:
### 1. Loading the Excel file:

* The code starts by importing essential libraries like datetime, pandas, re, and tkinter for file loading.
* It opens a dialog box using the tkinter library, allowing the user to select the Excel file that contains the HubSpot workflow data.

### 2. Credential input:

*The code prompts the user to enter the HubSpot account number and an authorization token (bearer token) required to make requests to the HubSpot API.

### 3. Reading the Excel file:

* The selected Excel file is read into a pandas DataFrame, which simplifies data manipulation and allows for calculations or transformations to be performed.

### 4. Transformations and link generation:

* Several functions are defined to add additional information to the DataFrame, such as:
  * wf_apply_link: Adds links to workflows with the appropriate URL.
  * folder_apply_link: Adds links to the folders associated with the workflows.
  * on_off: Defines whether the workflow is active or inactive, represented by 'On' or 'Off'.
  * is_issue: Indicates if there are any issues associated with the workflow.
  * recommended_action: Recommends actions (keep, review, delete) based on the workflow's activity conditions.
  * recommendation: Provides a more detailed recommendation based on the recommended action and identified issues.
  * re_enrollment: Checks if the workflow allows re-enrollment via the HubSpot API.

### 5. Processing and adding columns:

* The code applies these functions to the columns of the DataFrame, adding relevant information and generating new columns such as "Workflow Name," "Folder Name," "ON/OFF," "Issues?", "Recommended Action," and more.

### 6. Creating the audit file:

* Finally, it selects the relevant columns for the audit and saves them into a new Excel file named "Template HubSpot Audit.xlsx."

### 7. Error handling:

* The code has an exception handling block to ensure that if any error occurs during the creation of the audit file, it is captured, and a failure message is printed.

### Main functions:
* wf_apply_link: Generates a direct link to a workflow in HubSpot.
* folder_apply_link: Generates a link to the folder where the workflow resides in HubSpot.
* on_off: Indicates whether a workflow is active or inactive.
* is_issue: Detects if there are any issues associated with the workflow.
* recommended_action: Generates a recommendation for the workflow based on its status.
* recommendation: Provide a detailed recommendation based on the action recommendation and workflow issues.
* re_enrollment: Verifies if the workflow allows re-enrollment through the HubSpot API.
  
## Purpose:
This code is designed to audit workflows in HubSpot, offering users recommendations based on factors such as status, activity, and other variables. It is ideal for businesses using HubSpot to automate processes and seek more efficient workflow management and evaluation.

## Requirements
* Download the workflow report from the HubSpot account in .xlsx format.
* Ensure Python is installed.
* Account number for HubSpot
* Develop a private app in HubSpot with [**Automation**] scope.
* API key from the private application.
  
## Get Started
### Install local
1. Clone this repository with the following command:

   ```
   git clone https://github.com/Juan-pablo-delgado/OTF-WF-Audit
   ```
   
2. Install all dependencies with the following command:

   ```
   pip install -r requirements.txt
   ```

### How to use

1. Execute the following command:

   ```
    py workflow_audit_generator.py
   ```
   Or
   ```
   python workflow_audit_generator.py
   ```

2. Load the report workflow file in .xlsx format.
3. Fill the HubSpot account number
4. Fill the API key or bearer token
5. Wait for the process to complete.

* **Note:** A file named **'Template HubSpot Audit.xlsx'** will be automatically generated in the same folder.
  Please copy the content from the generated file into the template uploaded to Drive.

6. Complete the missing fields and initiate the workflow analysis.
7. After completion, access the presentation file, update the links, verify the information's accuracy, download a copy, and make necessary adjustments based on the client’s requirements.

### Cloud
1. The file is uploaded in the following [link](https://colab.research.google.com/drive/1asxBCYoybntURk7HCu9kas6rLlqa0eUW?usp=drive_link). To run it, use **(Ctrl + F9)**.
2. Load the report workflow file in .xlsx format.
3. Fill the HubSpot account number
4. Fill the API key or bearer token
5. Wait for the process to complete.

* **Note:** A file named '**Template HubSpot Audit.xlsx**' will be automatically downloaded.
  Please copy the content from the generated file into the template uploaded to Drive.
6. Complete the missing fields and initiate the workflow analysis.
7. After completion, access the presentation file, update the links, verify the accuracy of the information, download a copy, and make necessary adjustments based on the client’s requirements.

## Limitations
The HubSpot workflow API is currently in beta, which may result in certain data being unavailable. In such cases, manual entry of the field is required.
