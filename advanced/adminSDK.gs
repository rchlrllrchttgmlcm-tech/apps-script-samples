  let page;
  do {
    page = AdminReports.Activities.list("all", "login", {
      startTime: startTime,
      endTime: endTime,
      maxResults: 500,
      pageToken: pageToken,
    });
    const items = page.items;
    if (items) {
      for (const item of items) {
        const row = [
          new Date(item.id.time),
          item.actor.email,
          item.events[0].name,
        ];
        rows.push(row);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);

  if (rows.length === 0) {
    console.log("No results returned.");
    return;
  }
  const spreadsheet = SpreadsheetApp.create("Google Workspace Login Report");
  const sheet = spreadsheet.getActiveSheet();

  // Append the headers.
  const headers = ["Time", "User", "Login Result"];
  sheet.appendRow(headers);

  // Append the results.
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  console.log("Report spreadsheet created: %s", spreadsheet.getUrl());
}
// [END apps_script_admin_sdk_generate_login_activity_report]

// [START apps_script_admin_sdk_generate_user_usage_report]
/**
 * Generates a user usage report for this day last week as a spreadsheet. The
 * report includes the date, user, last login time, number of emails received,
 * and number of drive files created.
 * @see https://developers.google.com/admin-sdk/reports/reference/rest/v1/userUsageReport/get
 */
function generateUserUsageReport() {
  const today = new Date();
  const oneWeekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  const timezone = Session.getScriptTimeZone();
  const date = Utilities.formatDate(oneWeekAgo, timezone, "yyyy-MM-dd");

  const parameters = [
    "accounts:last_login_time",
    "gmail:num_emails_received",
    "drive:num_items_created",
  ];
  const rows = [];
  let pageToken;
  let page;
  do {
    page = AdminReports.UserUsageReport.get("all", date, {
      parameters: parameters.join(","),
      maxResults: 500,
      pageToken: pageToken,
    });
    if (page.warnings) {
      for (const warning of page.warnings) {
        console.log(warning.message);
      }
    }
    const reports = page.usageReports;
    if (reports) {
      for (const report of reports) {
        const parameterValues = getParameterValues(report.parameters);
        const row = [
          report.date,
          report.entity.userEmail,
          parameterValues["accounts:last_login_time"],
          parameterValues["gmail:num_emails_received"],
          parameterValues["drive:num_items_created"],
        ];
        rows.push(row);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);

  if (rows.length === 0) {
    console.log("No results returned.");
    return;
  }
  const spreadsheet = SpreadsheetApp.create(
    "Google Workspace User Usage Report",
  );
  const sheet = spreadsheet.getActiveSheet();

  // Append the headers.
  const headers = [
    "Date",
    "User",
    "Last Login",
    "Num Emails Received",
    "Num Drive Files Created",
  ];
  sheet.appendRow(headers);

  // Append the results.
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  console.log("Report spreadsheet created: %s", spreadsheet.getUrl());
}

/**
 * Gets a map of parameter names to values from an array of parameter objects.
 * @param {Array} parameters An array of parameter objects.
 * @return {Object} A map from parameter names to their values.
 */
function getParameterValues(parameters) {
  return parameters.reduce((result, parameter) => {
    const name = parameter.name;
    let value;
    if (parameter.intValue !== undefined) {
      value = parameter.intValue;
    } else if (parameter.stringValue !== undefined) {
      value = parameter.stringValue;
    } else if (parameter.datetimeValue !== undefined) {
      value = new Date(parameter.datetimeValue);
    } else if (parameter.boolValue !== undefined) {
      value = parameter.boolValue;
    }
    result[name] = value;
    return result;
  }, {});
}
// [END apps_script_admin_sdk_generate_user_usage_report]

// [START apps_script_admin_sdk_get_subscriptions]
/**
 * Logs the list of subscriptions, including the customer ID, date created, plan
 * name, and the sku ID. Notice the use of page tokens to access the full list
 * of results.
 * @see https://developers.google.com/admin-sdk/reseller/reference/rest/v1/subscriptions/list
 */
function getSubscriptions() {
  let result;
  let pageToken;
  do {
    result = AdminReseller.Subscriptions.list({
      pageToken: pageToken,
    });
    for (const sub of result.subscriptions) {
      const creationDate = new Date();
      creationDate.setUTCSeconds(sub.creationTime);
      console.log(
        "customer ID: %s, date created: %s, plan name: %s, sku id: %s",
        sub.customerId,
        creationDate.toDateString(),
        sub.plan.planName,
        sub.skuId,
      );
    }
    pageToken = result.nextPageToken;
  } while (pageToken);
}
// [END apps_script_admin_sdk_get_subscriptions]
