var _ = require('lodash'),
    Spreadsheet = require('edit-google-spreadsheet');

module.exports = {
    checkAuthOptions: function (step, dexter) {
        if(!step.input('range').first())
            return 'A [range] inputs variable is required for this module';

        if(!dexter.environment('google_spreadsheet'))
            return 'A [google_spreadsheet] environment variable is required for this module';

        return false;
    },

    convertColumnLetter: function(val) {
        var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
            result = 0;

        var i, j;

        for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {

            result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
        }

        return result;
    },

    pickSheetData: function (rows, startRow, startColumn, numRows, numColumns) {
        var pickSheetData = {};
        // pick data range
        for (var rowKey = startRow; rowKey <= numRows; rowKey++) {

            for (var columnKey = startColumn; columnKey <= numColumns; columnKey++) {

                if (rows[rowKey] !== undefined && rows[rowKey][columnKey] !== undefined) {

                    pickSheetData[rowKey] = pickSheetData[rowKey] || {};
                    pickSheetData[rowKey][columnKey] = rows[rowKey][columnKey];
                }
            }
        }
        // transform range to array
        return _.map(pickSheetData, function (columns) {

            return _.toArray(columns);
        });
    },

    parseRange: function (range) {
        var rgx = new RegExp('^([A-Z]+)(\\d+)(:)([A-Z]+)(\\d+)');

        if (!rgx.test(range)) {

            return false;
        } else {

            var oneRow = range.split(':')[0].match(/\d+/g)[0],
                oneColumn = this.convertColumnLetter(range.split(':')[0].match(/[A-Z]+/g)[0]),
                twoRow = range.split(':')[1].match(/\d+/g)[0],
                twoColumn = this.convertColumnLetter(range.split(':')[1].match(/[A-Z]+/g)[0]);

            return {
                startRow: Math.min(oneRow, twoRow),
                startColumn: Math.min(oneColumn, twoColumn),

                lastRow: Math.max(oneRow, twoRow),
                lastColumn: Math.max(oneColumn, twoColumn)
            };
        }
    },

    /**
     * The main entry point for the Dexter module
     *
     * @param {AppStep} step Accessor for the configuration for the step using this module.  Use step.input('{key}') to retrieve input data.
     * @param {AppData} dexter Container for all data used in this workflow.
     */
    run: function(step, dexter) {
        var credentials = dexter.provider('google').credentials(),
            error = this.checkAuthOptions(step, dexter);

        var spreadsheetId = dexter.environment('google_spreadsheet'),
            worksheetId = step.input('worksheet', 1).first(),
            range = step.input('range').first(),
            parseRange = this.parseRange(range);

        if (error) return this.fail(error);

        if (parseRange)
            Spreadsheet.load({
                //debug: true,
                spreadsheetId: spreadsheetId,
                worksheetId: worksheetId,
                accessToken: {
                    type: 'Bearer',
                    token: _.get(credentials, 'access_token')
                }
            }, function (err, spreadsheet) {

                if (err)
                    this.fail(err);
                else
                    spreadsheet.receive({getValues: true}, function (error, rows) {
                        error?
                            this.fail(error) :
                            this.complete({value: this.pickSheetData(rows, parseRange.startRow, parseRange.startColumn, parseRange.lastRow, parseRange.lastColumn)});
                    }.bind(this));
            }.bind(this));
        else
            this.fail('Range must be as "A2:D5"');

    }
};
