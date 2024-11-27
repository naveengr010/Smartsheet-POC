let smartsheetClient;  // Smartsheet JS client object

// Required libraries
const express = require("express");
const bodyParser = require("body-parser");
const smartsheetSdk = require("smartsheet");
const config = require("./config.json");

const app = express();
app.use(bodyParser.json());

// Initialize Smartsheet SDK client
function initializeSmartsheetClient(accessToken, logLevel) {
    smartsheetClient = smartsheetSdk.createClient({
        accessToken: accessToken,  // If token is falsy, the SDK will use the environment variable SMARTSHEET_ACCESS_TOKEN
        logLevel: logLevel
    });
}

// Test access to the Smartsheet
async function checkSheetAccessibility(sourceSheetId) {
    console.log(`Checking sheet with ID: ${sourceSheetId}`);
    const getSheetOptions = {
        id: sourceSheetId,
        queryParameters: { pageSize: 1 } // Limit response to 1 row for efficiency
    };
    try {
        const sheetResponse = await smartsheetClient.sheets.getSheet(getSheetOptions);
        console.log(`Successfully accessed sheet: "${sheetResponse.name}" at ${sheetResponse.permalink}`);
    } catch (error) {
        console.error(`Error accessing sheet with ID: ${sourceSheetId}`, error);
    }
}

/*
* Initialize or verify an existing webhook for a target sheet.
* This ensures the webhook callback URL is set up correctly.
*/
async function setupWebhookForSheet(sourceSheetId, webhookName, callbackUrl) {
    try {
        let existingWebhook = null;

        // Retrieve all webhooks
        const webhookListResponse = await smartsheetClient.webhooks.listWebhooks({
            includeAll: true
        });
        console.log(`Found ${webhookListResponse.totalCount} webhooks for the user`);

        // Check for an existing webhook for the given sheet
        existingWebhook = webhookListResponse.data.find(hook =>
            hook.scopeObjectId === sourceSheetId && hook.name === webhookName
        );

        if (!existingWebhook) {
            // Create a new webhook if not found
            const createWebhookOptions = {
                body: {
                    name: webhookName,
                    callbackUrl: callbackUrl,
                    scope: "sheet",
                    scopeObjectId: sourceSheetId,
                    events: ["*.*"],
                    version: 1
                }
            };

            const createResponse = await smartsheetClient.webhooks.createWebhook(createWebhookOptions);
            existingWebhook = createResponse.result;
            console.log(`Created new webhook with ID: ${existingWebhook.id}`);
        }

        // Ensure webhook is enabled and set to the correct callback URL
        const updateWebhookOptions = {
            webhookId: existingWebhook.id,
            callbackUrl: callbackUrl,
            body: { enabled: true }
        };

        const updateResponse = await smartsheetClient.webhooks.updateWebhook(updateWebhookOptions);
        console.log(`Webhook updated: Enabled = ${updateResponse.result.enabled}, Status = ${updateResponse.result.status}`);

    } catch (error) {
        console.error("Error setting up webhook:", error.stack);  // Log full stack trace for debugging
    }
}

// Handle webhook callbacks
app.post("/", async (req, res) => {
    try {
        const body = req.body;

        if (body.challenge) {
            // Verification callback
            console.log("Received verification callback from Smartsheet");
            res.status(200).json({ smartsheetHookResponse: body.challenge });
        } else if (body.events) {
            // Event callback (sheet updates)
            console.log(`Received ${body.events.length} event(s) at ${new Date().toLocaleString()}`);
            await processEvents(body);
            res.sendStatus(200);
        } else if (body.newWebHookStatus) {
            // Webhook status update callback
            console.log(`Received status update: ${body.newWebHookStatus}`);
            res.sendStatus(200);
        } else {
            console.log(`Received unknown callback: ${JSON.stringify(body)}`);
            res.sendStatus(200);
        }
    } catch (error) {
        console.error("Error processing webhook event:", error.stack);
        res.status(500).send(`Error: ${error}`);
    }
});

/*
* Process the events sent by Smartsheet.
* For this example, we only handle cell changes, but you can extend it as needed.
*/
async function processEvents(callbackData) {
    if (callbackData.scope !== "sheet") return;

    for (const event of callbackData.events) {
        if (event.objectType === "cell") {
            console.log(`Cell changed, Row ID: ${event.rowId}, Column ID: ${event.columnId}`);

            // Fetch row and column details
            const options = {
                id: callbackData.scopeObjectId, // Source sheet ID from callback
                queryParameters: {
                    rowIds: event.rowId.toString(),
                    columnIds: event.columnId.toString()
                }
            };

            try {
                const sheetResponse = await smartsheetClient.sheets.getSheet(options);
                const row = sheetResponse.rows[0];
                const cell = row.cells[0];
                const column = sheetResponse.columns.find(c => c.id === cell.columnId);

                const cellValue = cell.displayValue || null;

                console.log(`Updated value: "${cellValue}" in column "${column.title}", row number ${row.rowNumber}`);

                await syncWithDestinationSheet(cellValue, column, row);
            } catch (error) {
                console.error("Error processing event:", error.stack);
            }
        }
    }
}

/*
* Sync the updated value to the corresponding row and column in the destination sheet.
*/
async function syncWithDestinationSheet(cellValue, column, row) {
    const destinationSheetId = config.destinationSheetId;  // ID of the destination sheet

    try {
        const sheetResponse = await smartsheetClient.sheets.getSheet({ id: destinationSheetId });
        const columnInDestinationSheet = sheetResponse.columns.find(c => c.title === column.title);
        if (!columnInDestinationSheet) {
            console.log(`Column "${column.title}" not found in destination sheet.`);
            return;
        }

        const rowInDestinationSheet = sheetResponse.rows.find(r => r.rowNumber === row.rowNumber);
        if (!rowInDestinationSheet) {
            console.log(`Row ${row.rowNumber} not found in destination sheet.`);
            return;
        }

        // Update destination sheet with the new cell value
        await updateDestinationSheet(destinationSheetId, rowInDestinationSheet, columnInDestinationSheet, cellValue);

    } catch (error) {
        console.error("Error syncing data to destination sheet:", error.stack);
    }
}

/*
* Update a row in the destination sheet with the new cell value.
*/
async function updateDestinationSheet(destinationSheetId, row, column, cellValue) {
    const rowToUpdate = {
        id: row.id,
        cells: [{
            columnId: column.id,
            value: cellValue || null // Use null or empty string if value is missing
        }]
    };

    const updateOptions = {
        sheetId: destinationSheetId,
        body: rowToUpdate
    };

    try {
        const updateResponse = await smartsheetClient.sheets.updateRow(updateOptions);
        console.log(`Updated row in destination sheet with value: ${cellValue}`);
    } catch (error) {
        console.error("Error updating row in destination sheet:", error.stack);
    }
}

// Main execution
(async () => {
    try {
        initializeSmartsheetClient(config.smartsheetAccessToken, config.logLevel);
        await checkSheetAccessibility(config.sourceSheetId);  // Sanity check to ensure access to the source sheet

        app.listen(3000, () => {
            console.log("Server listening on port 3000");
        });

        // Set up webhook
        await setupWebhookForSheet(config.sourceSheetId, config.webhookName, config.callbackUrl);

    } catch (error) {
        console.error("Error during application initialization:", error.stack);
    }
})();
