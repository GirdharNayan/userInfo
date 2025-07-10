API_KEY = '';
API_THROTTLE_TIME = '';
SNIPE_IT_URL = '';

function getData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("sheet1");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 2, lastRow - 1, 9).getValues();

  SpreadsheetApp.getUi().alert(
    `‼️ This script will now establish connection with snipe to pull requested data ‼️\n\n` +
    `you can see data being populated in sheet in real time 🏋️\n` +
    `Please be patient while script continues🤓\n` +
    `Press "OK" to continue with this script😎\n`
  );

  Logger.log("Starting to process rows...");

  data.forEach((row, index) => {
    const email = row[0];
    const columnIValue = row[8]; 

    Logger.log("🏋️ Processing row " + (index + 2) + ": Email = " + email + ", Column I value = " + columnIValue);

    if (!columnIValue) {
      const userInfo = fetchUserInfoByEmail(email);
      let firstName = '';

      if (userInfo && userInfo.first_name) {
        firstName = userInfo.first_name;
      }

      const assetsAndCategories = fetchAssetsByEmail(email);
      if (assetsAndCategories.assets.length > 0) {
        sheet.getRange(index + 2, 1).setValue(firstName);
        sheet.getRange(index + 2, 3).setValue(assetsAndCategories.assets.join('\n'));
        sheet.getRange(index + 2, 4).setValue(assetsAndCategories.serialNumbers.join('\n'));
        sheet.getRange(index + 2, 6).setValue(assetsAndCategories.categories.join('\n'));
        sheet.getRange(index + 2, 8).setValue(assetsAndCategories.assetNames.join('\n'));
        Logger.log("Assets processed for " + email);
      }

      SpreadsheetApp.flush();
      Utilities.sleep(API_THROTTLE_TIME);

      const accessoriesAndCategories = fetchAccessoriesByEmail(email);
      if (accessoriesAndCategories.accessories.length > 0) {
        sheet.getRange(index + 2, 5).setValue(accessoriesAndCategories.accessories.join('\n'));
        sheet.getRange(index + 2, 7).setValue(accessoriesAndCategories.categories.join('\n'));
        Logger.log("Accessories processed for " + email);
      }

      SpreadsheetApp.flush();
      Utilities.sleep(API_THROTTLE_TIME);
    } else {
      Logger.log("Skipping row " + (index + 2) + " as column I is not empty.");
    }
  });

  Logger.log("Finished processing all rows.");
}


function fetchUserInfoByEmail(email) {
  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + API_KEY,
      'Content-Type': 'application/json'
    }
  };

  const url = SNIPE_IT_URL + '/users?email=' + encodeURIComponent(email);
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data.total > 0 && data.rows.length > 0) {
    return data.rows[0]; 
  }

  return null;
}
