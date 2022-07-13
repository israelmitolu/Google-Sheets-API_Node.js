const express = require("express");

const { google } = require("googleapis");

// Initialise express and listen for the server on port 8080
const app = express();
const port = 8080;

//This allows us to parse the incoming request body as JSON
app.use(express.json());

// With this, we'll listen for the server on port 8080
app.listen(port, () => console.log(`Listening on port ${port}`));

// Spreadsheet ID
const id = "12fSSjNTpDtTJ8vF4T3LRYe36WjPEfn1_sr_lSSpZJUo";

//Authorise the Google client
async function authSheets() {
  //Function for authentication object
  const auth = new google.auth.GoogleAuth({
    keyFile: "keys.json",
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  //Create client instance for auth
  const authClient = await auth.getClient();

  //Instance of the Sheets API
  const sheets = google.sheets({ version: "v4", auth: authClient });

  return {
    auth,
    authClient,
    sheets,
  };
}

app.get("/", async (req, res) => {
  const { sheets } = await authSheets();

  // Read rows from spreadsheet
  const getRows = await sheets.spreadsheets.values.get({
    spreadsheetId: id,
    range: "Sheet1",
  });

  // Write rows to spreadsheet
  await sheets.spreadsheets.values.append({
    spreadsheetId: id,
    range: "Sheet1",
    valueInputOption: "USER_ENTERED",
    resource: {
      values: [["Gabriella", "Female", "4. Senior"]],
    },
  });

  // Updating a specified range in the sheet
  await sheets.spreadsheets.values.update({
    spreadsheetId: id,
    range: "Sheet1!C2",
    valueInputOption: "USER_ENTERED",
    resource: {
      values: [["2. Sophomore"]],
    },
  });

  // Delete data from spreadsheet
  await sheets.spreadsheets.values.clear({
    spreadsheetId: id,
    range: "Sheet1!A6:C6",
  });

  // Update spreadsheet formatting
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    resource: {
      requests: [
        {
          updateBorders: {
            range: {
              sheetId: 0,
              startRowIndex: 0,
              endRowIndex: 6,
              startColumnIndex: 0,
              endColumnIndex: 3,
            },
            top: {
              style: "DASHED",
              width: 1,
              color: {
                red: 1.0,
              },
            },
            bottom: {
              style: "DASHED",
              width: 1,
              color: {
                red: 1.0,
              },
            },
            innerHorizontal: {
              style: "DASHED",
              width: 1,
              color: {
                red: 1.0,
              },
            },
          },
        },
      ],
    },
  });

  //Output
  res.send(getRows.data);
});
