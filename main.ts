import "jsr:@std/dotenv/load";
import { google } from "npm:googleapis";
import * as azdev from "npm:azure-devops-node-api";
import * as WorkItemTrackingInterfaces from "npm:azure-devops-node-api/interfaces/WorkItemTrackingInterfaces.js";
import * as GitInterfaces from "npm:azure-devops-node-api/interfaces/GitInterfaces.js";

const AZURE_ORG = Deno.env.get("AZURE_ORG")!;
const AZURE_PROJECT = Deno.env.get("AZURE_PROJECT")!;
const AZURE_PERSONAL_ACCESS_TOKEN = Deno.env.get(
  "AZURE_PERSONAL_ACCESS_TOKEN"
)!;
const REPOSITORY_ID_FE = Deno.env.get("REPOSITORY_ID_FE")!;
const REPOSITORY_ID_BE = Deno.env.get("REPOSITORY_ID_BE")!;
const SHEET_ID = Deno.env.get("SHEET_ID")!;
const QUERY_ID = Deno.env.get("QUERY_ID")!;

const AZURE_BASE_URL = `https://dev.azure.com/${AZURE_ORG}`;

// Function to fetch commits
async function fetchCommits(): Promise<GitInterfaces.GitCommit[]> {
  const authHandler = azdev.getPersonalAccessTokenHandler(
    AZURE_PERSONAL_ACCESS_TOKEN
  );
  const connection = new azdev.WebApi(AZURE_BASE_URL, authHandler);
  const MIN_DATE = new Date();
  MIN_DATE.setDate(MIN_DATE.getDate() - 1);

  const gitApi = await connection.getGitApi();

  const frontEndCommits = await gitApi.getCommits(
    REPOSITORY_ID_FE,
    {
      fromDate: MIN_DATE.toISOString(),
    },
    AZURE_PROJECT
  );

  const backEndCommits = await gitApi.getCommits(
    REPOSITORY_ID_BE,
    {
      fromDate: MIN_DATE.toISOString(),
    },
    AZURE_PROJECT
  );

  const allCommits = [...frontEndCommits, ...backEndCommits];

  return allCommits;
}

// Function to fetch pull requests
async function fetchPullRequests(): Promise<GitInterfaces.GitPullRequest[]> {
  const authHandler = azdev.getPersonalAccessTokenHandler(
    AZURE_PERSONAL_ACCESS_TOKEN
  );
  const connection = new azdev.WebApi(AZURE_BASE_URL, authHandler);
  const MIN_DATE = new Date();
  MIN_DATE.setDate(MIN_DATE.getDate() - 1);

  const gitApi = await connection.getGitApi();
  const frontEndPRs = await gitApi.getPullRequestsByProject(AZURE_PROJECT, {
    repositoryId: REPOSITORY_ID_FE,
    minTime: MIN_DATE,
  });
  const backEndPRs = await gitApi.getPullRequestsByProject(AZURE_PROJECT, {
    repositoryId: REPOSITORY_ID_BE,
    minTime: MIN_DATE,
  });

  const allPRs = [...frontEndPRs, ...backEndPRs];

  return allPRs;
}

// Function to fetch pull requests
async function fetchWorkItems(): Promise<
  WorkItemTrackingInterfaces.WorkItem[]
> {
  const authHandler = azdev.getPersonalAccessTokenHandler(
    AZURE_PERSONAL_ACCESS_TOKEN
  );
  const connection = new azdev.WebApi(AZURE_BASE_URL, authHandler);

  const workItemTrackingApi = await connection.getWorkItemTrackingApi();

  const workItems = await workItemTrackingApi.queryById(QUERY_ID);
  const workItemIds = workItems?.workItems?.map(
    (workItem: WorkItemTrackingInterfaces.WorkItemReference) => workItem.id
  );

  const workItemFields = [
    "System.Id",
    "System.Title",
    "System.WorkItemType",
    "System.Reason",
    "System.AssignedTo",
    "System.ChangedDate",
    "System.ChangedBy",
    "System.State",
  ];
  if (workItemIds) {
    const workItemData = await workItemTrackingApi.getWorkItemsBatch({
      // $expand: WorkItemTrackingInterfaces.WorkItemExpand.All,
      ids: workItemIds.filter((x) => typeof x === "number") as number[],
      fields: workItemFields,
    });

    return workItemData;
  }
  return [];
}

// Function to get Google Sheets client
async function getGoogleSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// Function to update Google Sheets
async function updateGoogleSheet(
  commits: GitInterfaces.GitCommit[],
  pullRequests: GitInterfaces.GitPullRequest[],
  workItems: WorkItemTrackingInterfaces.WorkItem[]
) {
  const sheets = await getGoogleSheetsClient();
  const today = new Date().toISOString().split("T")[0]; // Format date as YYYY-MM-DD

  // Helper function to add a merged row
  async function addMergedRow(
    sheetName: string,
    message: string,
    rowIndex: number,
    columns: number = 5
  ) {
    const sheet = await sheets.spreadsheets.get({
      spreadsheetId: SHEET_ID,
    });

    const sheetId = sheet.data.sheets?.find(
      (s) => s.properties?.title === sheetName
    )?.properties?.sheetId;

    if (!sheetId && sheetId !== 0) {
      throw new Error(`Sheet ${sheetName} not found.`);
    }
    const rangeResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!A:A`, // Assuming column A always has data
    });

    const lastRowIndex = (rangeResponse.data.values?.length || 0) + 1; // Get next available row index

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          {
            mergeCells: {
              range: {
                sheetId,
                startRowIndex: lastRowIndex - 1, // Convert to 0-indexed
                endRowIndex: lastRowIndex, // One row below
                startColumnIndex: 0,
                endColumnIndex: columns,
              },
              mergeType: "MERGE_ALL",
            },
          },
          {
            updateCells: {
              range: {
                sheetId,
                startRowIndex: lastRowIndex - 1, // Same row as merge
                endRowIndex: lastRowIndex,
                startColumnIndex: 0,
              },
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: { stringValue: message },
                      userEnteredFormat: {
                        horizontalAlignment: "CENTER",
                        textFormat: { bold: true },
                      },
                    },
                  ],
                },
              ],
              fields: "userEnteredValue,userEnteredFormat",
            },
          },
        ],
      },
    });

    return lastRowIndex; // Return the new header row index
  }

  // Add commits section
  await addMergedRow("Commits", `Commits Data for ${today}`, 0);
  console.log("Adding commits to Google Sheet...");
  const commitData = commits.map((commit) => [
    commit.commitId,
    commit.author?.name,
    commit.comment,
    commit.author?.date,
  ]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "Commits!A2",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [...commitData],
    },
  });
  console.log("Google Sheet updated with commits!");

  // Add pull requests section
  await addMergedRow("PullRequests", `Pull Requests Data for ${today}`, 0);
  console.log("Adding pull requests to Google Sheet...");
  const prData = pullRequests.map((pr) => [
    pr.pullRequestId,
    pr.title,
    pr.createdBy?.displayName,
    pr.creationDate,
    pr.workItemRefs?.map((workItem) => workItem.id).join(", "),
  ]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "PullRequests!A2",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [...prData],
    },
  });
  console.log("Google Sheet updated with pull requests!");

  // Add work items section
  await addMergedRow("WorkItems", `Work Items Data for ${today}`, 0, 8);
  console.log("Adding work items to Google Sheet...");
  const workItemData = workItems.map((workItem) => [
    workItem.id,
    workItem.fields?.["System.Title"],
    workItem.fields?.["System.WorkItemType"],
    workItem.fields?.["System.Reason"],
    workItem.fields?.["System.AssignedTo"]?.displayName,
    workItem.fields?.["System.ChangedDate"],
    workItem.fields?.["System.ChangedBy"]?.displayName,
    workItem.fields?.["System.State"],
  ]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "WorkItems!A2",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [...workItemData],
    },
  });
  console.log("Google Sheet updated with work items!");

  console.log("Google Sheet updated successfully with merged date rows!");
}

// Main function
async function main() {
  try {
    console.log("Fetching commits...");
    const commits = await fetchCommits();
    console.log("Retrieved Commits:", commits.length);
    console.log("Fetching pull requests...");
    const pullRequests = await fetchPullRequests();
    console.log("Retrieved Pull Requests:", pullRequests.length);
    console.log("Fetching work items...");
    const workItems = await fetchWorkItems();
    console.log("Retrieved Work Items:", workItems.length);
    await updateGoogleSheet(commits, pullRequests, workItems);
  } catch (error) {
    console.error("Error:", error);
  }
}

main();
