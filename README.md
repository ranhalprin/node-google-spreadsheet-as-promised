# Basic Google Spreadsheet Access with Promises and Simple Cell Notation (node.js)

[![NPM version](https://badge.fury.io/js/google-spreadsheet-as-promised.png)](http://badge.fury.io/js/google-spreadsheet-as-promised)

This is a very basic node.js module that allows connecting to a Google Spreadsheet, changing cell values and reading cell values. It is useful in the context where a spreadsheet has complicated logic that you do not want to move into code, or if you want to validate that your code behaves the same as the spreadsheet for different inputs.

## Installation
``` npm install google-spreadsheet-as-promised ```

## Usage

The following usage example is complete, meaning it covers all provided functionality. It shows how to change some parameter cells in the spreadsheet and then read a value from a different cell.

```javascript
ֿvar GoogleSpreadsheetAsPromised = require('google-spreadsheet-as-promised');
var CREDS = require('./google-api-creds.json');

var SPREADSHEET_KEY = '<spreadsheet key>';

var WORKSHEET_ID = 0;
var PARAMETER_RANGE = 'B4:C12'; // Matrix range covers all cells where parameters should be changed
var PARAMETER_CELLS = {
    parameter_1: 'B4',
    parameter_2: 'B5',
    parameter_3: 'C4',
    parameter_4: 'C10'
    parameter_5: 'C12'
};
var RESULT_CELL = 'D14';

var parameters = {
    parameter_1: '100',
    parameter_2: '0.5',
    parameter_3: '17',
    parameter_4: '55'
    parameter_5: 'NONE'
};

getResultWithParameter(parameters);

// Given parameters, resolves to the result cell after
function getResultWithParameter(parameters) {
    return new Promise(function(resolve, reject) {
        var sheet = new GoogleSpreadsheetAsPromised();
        var worksheet;
        sheet.load(SPREADSHEET_KEY, CREDS).then(function() {
            return sheet.getWorksheet(WORKSHEET_ID);
        }).then(function(resultWorksheet) {
            worksheet = resultWorksheet;
            return worksheet.getCells(PARAMETER_RANGE); 
        }).then(function(cells) {
            var promises = [];
            var parameters = Object.keys(parameters);
            for (var i = 0; i < parameters.length; i++) {
                var name = parameters[i];
                var value = parameters[name];
                var cell = PARAMETER_CELLS.parameters[name];
                if (!cell) {
                    return reject("Unknown Parameter: " + name);
                }
                promises.push(cells.setValue(cell, value));
            }
            return Promise.all(promises); // This makes all values change in parallel
        }).then(function() {
            return worksheet.getCell(RESULT_CELL); // We must load the result cell only after parameter values are set
        }).then(function(cell) {
            return resolve(cell.getValue());
        });    
    });
});
```

### `GoogleSpreadsheetAsPromised`

The main class that represents an entire spreadsheet.


#### `new GoogleSpreadsheetAsPromised()`

Create a new google spreadsheet object.

#### 'GoogleSpreadsheetAsPromised.load(sheet_id, creds)'

Load must be called before any other functionality is available.
Resolves to undefined.

- `sheet_id` -- the ID of the spreadsheet (from its URL)
- `creds` -- the credential needed to access the spreadsheet (see the "Authentication" section)

#### 'GoogleSpreadsheetAsPromised.getWorksheet(worksheet_id)'

Resolves to a WorksheetAsPromised object that represents one of the worksheets (tabs) in the spreadsheet.

- `worksheet_id` -- the numeric ID of the worksheet you want to work with, usually 0 for the default tab.

#### 'WorksheetAsPromised.getCells(range)'

Resolves to a CellsAsPromised object that represents a groups of cells.

-- `range` -- a string representation of the cell range to load in the form ```'<min-col><min-row>:<max-col><max-row>'``` for example 'A1:B4' will load the eight cells 'A1','A2','A3','A4','B1','B2','B3','B4'

#### 'WorksheetAsPromised.getCell(cell)'

Resolves to a CellAsPromised object that represents a single cells.

-- `cell` -- a string representation of the cell in the form ```'<col><row>'```, for example 'B2'

#### 'CellsAsPromised.getValue(cell)'

Resolves to the value in a given cell.

-- `cell` -- a string representation of the cell in the form ```'<col><row>'```, for example 'B2'

#### 'CellsAsPromised.setValue(cell, value)'

Sets a value to a cell in the spreadsheet.
Resolves to undefined.

-- `cell` -- a string representation of the cell in the form ```'<col><row>'```, for example 'B2'
-- `value` -- a string value to write into the cell

#### 'CellAsPromised.getValue()'

Resolves to the value in the cell.

-- `cell` -- a string representation of the cell in the form ```'<col><row>'```, for example 'B2'

#### 'CellsAsPromised.setValue(value)'

Sets a value to the cell.
Resolves to undefined.

-- `value` -- a string value to write into the cell

## Authentication
 
The recommended authentication method is to give a google api service access to the spreadsheet. There is no need to create a new user for the module to use it.

([More about Service Authentication](https://developers.google.com/identity/protocols/OAuth2ServiceAccount))

__Setup Instructions__

1. Go to the [Google Developers Console](https://console.developers.google.com/project)
2. Select your project or create a new one (and then select it)
3. Enable the Drive API for your project
  - In the sidebar on the left, expand __APIs & auth__ > __APIs__
  - Search for "drive"
  - Click on "Drive API"
  - click the blue "Enable API" button
4. Create a service account for your project
  - In the sidebar on the left, expand __APIs & auth__ > __Credentials__
  - Click blue "Add credentials" button
  - Select the "Service account" option
  - Select the "JSON" key type option
  - Click blue "Create" button
  - your JSON key file is generated and downloaded to your machine (__it is the only copy!__)
  - note your service account's email address (also available in the JSON key file)
5. Share the doc (or docs) with your service account using the email noted above

If you are using heroku or another environment where you cannot save a local file, the recommended method is to save the two required fields into enviornment variables:

- `client_email` -- your service account's email address
- `private_key` -- the private key found in the JSON file

And then load them in runtime like so:

```
creds.private_key = process.env.GOOGLE_API_KEY.replace(/\\n/g,"\n"); // Newlines are escaped in env variables
creds.client_email = process.env.GOOGLE_API_EMAIL;
```

## Acknowledgement

This module wraps the [google-spreadsheet nodejs module](https://www.npmjs.org/package/google-spreadsheet). It provides two changes in API:
- Instead of callbacks, it uses bluebird promises for easier flow control
- Instead of supplying the cell ranges as numbered parameters, they are supplied as strings in standard spreadsheet notation

NOTE: The wrapper is very limited and supports only a limited set of functionality the exists in the original module.

### Contributions

If you need more features or find bugs, please contribute your improvements to [our github](https://github.com/ranhalprin/node-google-spreadsheet-as-promised) for the benefit of the community.

### Updates to underlying module

In order to maintain compatability this wrapper is locked to a specific version of the google-spreadsheet module.

## Future features

Support bulk setting of multiple cells in one API call to reduce network overhead and response time.

## Links
- <https://github.com/ranhalprin/node-google-spreadsheet-as-promised>
- <https://www.npmjs.com/package/google-spreadsheet/>
- <https://developers.google.com/google-apps/spreadsheets/>

## License
google-spreadsheets-as-promised is free and unencumbered public domain software. For more information, see the accompanying UNLICENSE file.