**About**

Gets specific data from Azure DevOps using Azure Devops services API and stores in a google sheet.


**Create Google API Credentials**

1.  Go to the Google Cloud Console.
2.  Create a new project and enable the Google Sheets API.
3.  Create Service Account Credentials and download the credentials.json file.
4.  Share your Google Sheet with the service account's email (e.g., your-service-account@your-project.iam.gserviceaccount.com).

Example env is provided

to run `deno run -A main.ts`
the -A switch is to allow permissions without user interaction.