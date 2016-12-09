var Promise = require('bluebird');
var GoogleSpreadsheetAsPromised = require('./index.js'); // require('google-spreadsheet-as-promised') externally

var CREDS = require('./google-api-creds.json');

var SPREADSHEET_KEY = '<spreadsheet key>'; 

var WORKSHEET_ID = 0;
var PARAMETER_RANGE = 'B2:C8'; // Matrix range covers all cells where parameters should be changed
var PARAMETER_CELLS = {
    parameter1: 'B2',
    parameter2: 'B4',
    parameter3: 'C7',
    parameter4: 'C8'
};
var TOTAL_CELL = 'D20';
var INTERMEDIATE_CALCULATIONS_RANGE = 'D10:D14';

var parameters = {
    parameter1: '100',
    parameter2: '0.5',
    parameter3: '17',
    parameter4: '55'
};

getResultWithParameter(parameters).then(function(result) {
    console.log(result);
});

// Given parameters, resolves to the result cell after
function getResultWithParameter(parameters) {
    return new Promise(function(resolve, reject) {
        var sheet = new GoogleSpreadsheetAsPromised();
        var worksheet;
        var result = {};
        sheet.load(SPREADSHEET_KEY, CREDS).then(function() {
            return sheet.getWorksheet(WORKSHEET_ID);
        }).then(function(resultWorksheet) {
            worksheet = resultWorksheet;
            return worksheet.getCells(PARAMETER_RANGE); 
        }).then(function(cells) {
            var promises = [];
            var names = Object.keys(parameters);
            for (var i = 0; i < names.length; i++) {
                var name = names[i];
                var value = parameters[name];
                var cell = PARAMETER_CELLS[name];
                if (!cell) {
                    return reject("Unknown Parameter: " + name);
                }
                promises.push(cells.setValue(cell, value));
            }
            return Promise.all(promises); // This makes all values change in parallel 
        }).then(function() {
            return worksheet.getCell(TOTAL_CELL); // We must load the result cell only after parameter values are all set
        }).then(function(cell) {
            result.total = cell.getValue();
            return worksheet.getCells(INTERMEDIATE_CALCULATIONS_RANGE);
        }).then(function(cells) {
            result.intermediate_values_array = cells.getAllValues();
            return resolve(result);
        });
    });
};