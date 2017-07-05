var Promise = require('bluebird');
var GoogleSpreadsheet = require('google-spreadsheet');

var parseColumn = function(string) {
    var result = 0;
    for (var i = 0; i < string.length; i++) {
        result *= 26;
        result += parseInt(string.charCodeAt(i) - 65) + 1; // Excel is 1-based
    }
    return result;
};

function CellAsPromised(cell) {
    this.getValue = function() {
        return cell.value;
    };

    this.setValue = function(value) {
        return new Promise(function(resolve, reject) {
            cell.setValue(value, function(err) {
                if (err) {
                    return reject(err);
                }
                return resolve();
            });
        });
    };
}

function CellsAsPromised(cells, range) {

    var cellToIndex = function(cellString) {
        var col = parseColumn(cellString.replace(/\d/g,''));
        var row = parseInt(cellString.replace(/\D/g,''));
        return (col - range['min-col']) + (row - range['min-row']) * (range['max-col'] - range['min-col'] + 1);
    };

    this.getValue = function(cellString) {
        var index = cellToIndex(cellString);
        return cells[index].value;
    };

    this.getAllValues = function() {
        var result = [];
        for (var i = 0; i < cells.length; i++) {
            result.push(cells[i].value);
        }
        return result;
    };

    this.getWidth = function() {
        return range['max-col'] - range['min-col'] + 1;
    };

    this.getHeight = function() {
        return range['max-row'] - range['min-row'] + 1;
    };

    this.setValue = function(cellString, value) {
        var index = cellToIndex(cellString);
        return new Promise(function(resolve, reject) {
            cells[index].setValue(value, function(err) {
                if (err) {
                    return reject(err);
                }
                return resolve();
            });
        });
    };
}

function WorksheetAsPromised(worksheet) {

    var parseRange = function(range) {
        var parts = range.split(':');
        return {
            'min-col': parseColumn(parts[0].replace(/\d/g,'')),
            'max-col': parseColumn(parts[1].replace(/\d/g,'')),
            'min-row': parseInt(parts[0].replace(/\D/g,'')),
            'max-row': parseInt(parts[1].replace(/\D/g,''))
        };
    };

    this.getCells = function(rangeString) {
        var range = parseRange(rangeString);
        params = range;
        params['return-empty'] = 'true';
        return new Promise(function(resolve, reject) {
            worksheet.getCells(params, function(err, cells) {
                if (err) {
                    return reject(err);
                }
                return resolve(new CellsAsPromised(cells, range));
            });
        });
    };

    this.getCell = function(cellString) {
        var range = parseRange(cellString + ":" + cellString);
        params = range;
        params['return-empty'] = 'true';
        return new Promise(function(resolve, reject) {
            worksheet.getCells(params, function(err, cells) {
                if (err) {
                    return reject(err);
                }
                return resolve(new CellAsPromised(cells[0]));
            });
        });
    };
}

function GoogleSpreadsheetAsPromised() {

    var self = this;

    this.load = function(spreadsheet_key, creds) {
        return new Promise(function(resolve, reject) {
            var doc = new GoogleSpreadsheet(spreadsheet_key);
            doc.useServiceAccountAuth(creds, function(err) {
                if (err) {
                    return reject(err);
                }
                doc.getInfo(function(err, info) {
                    if (err) {
                        return reject(err);
                    }
                    self.worksheets = info.worksheets;
                    self.worksheetNames = {};
                    for (var i = 0; i < self.worksheets.length; i++) {
                        self.worksheetNames[self.worksheets[i].title] = i;
                    }
                    return resolve();
                });
            });
        });
    };

    this.getWorksheet = function(index) {
        return new Promise(function(resolve, reject) {
            if (!self.worksheets[index]) {
                return reject("Cannot find worksheet index: " + index);
            }
            return resolve(new WorksheetAsPromised(self.worksheets[index]));
        });
    };

    this.getWorksheetByName = function(name) {
        return new Promise(function(resolve, reject) {
            var index = self.worksheetNames[name];
            if (index === undefined) {
                return reject("Cannot find worksheet name: `" + name + "`");
            }
            self.getWorksheet(index).then(function(worksheetAsPromised) {
                return resolve(worksheetAsPromised);
            });
        });
    };
}

module.exports = GoogleSpreadsheetAsPromised;
