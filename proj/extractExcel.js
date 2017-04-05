/**
 * Created by pengl on 3/27/2017.
 * Description: This script is used to extract URLs from Propel Org Excel
 * Parameters : xlsxFile
 * Return     : Array of JSON
 */

const xlsxFile = 'PROD_ORG_Info_Excel.xlsx';  //----- Needs modification in the future

var xlsx = require('node-xlsx').default;
const workSheetsFromFile = xlsx.parse( xlsxFile );
var propelAccountsArray = [];
var propelAccountsObj = {};
/************************************************
 * Main Processes
 * 1) Import Excel, and find Account Sheet
 * 2) Get columns ID for required Attributes
 * 3) Extract info, fill it in one Array
 ************************************************/
var activeSheet = getAccountSheet();
var dataInActiveSheet = workSheetsFromFile[activeSheet].data;

getColumnsForAttributes();
extractPropelAccounts();

function getAccountSheet() {
    for ( var k=0; k < workSheetsFromFile.length; k++ ) {
        if ('Account' == workSheetsFromFile[k].name) {
            return k;
        }
    }
}

function getColumnsForAttributes() {
    var accountTitles = workSheetsFromFile[activeSheet].data[0];

    for(var k=0; k< accountTitles.length; k++) {
        if('QRS CUSTOMER NAME' === accountTitles[k]) {
            checkPoint = k;
        }

        if('URL' === accountTitles[k]) {
            urlCol = k;
        }

        if('Propel Account' === accountTitles[k]) {
            accountCol = k;
        }

        if('Propel Password' === accountTitles[k]) {
            passwordCol = k;
        }
    }

    if( checkPoint == 0 || urlCol == 0 || accountCol == 0 || passwordCol == 0){

        alert("I can't find all required countOf...");
    }
}

function extractPropelAccounts( ) {

    for (var k=1; k< dataInActiveSheet.length ; k++) {
        if( dataInActiveSheet[k].length < 10 && dataInActiveSheet[k][checkPoint] == undefined ){
            //console.log("Not valid customer...");
        } else {
            var customerRecord = {url: '', account: 'migration', password: ''};
            customerRecord.url = dataInActiveSheet[k][urlCol];
            customerRecord.account = dataInActiveSheet[k][accountCol];
            customerRecord.password = dataInActiveSheet[k][passwordCol];
            propelAccountsArray.push( customerRecord );
            propelAccountsObj[customerRecord.url] = customerRecord;
        }
    }
    //console.log( propelAccountsArray.length );
    //console.log( propelAccountsArray );
    //console.log(propelAccountsObj);
}

module.exports = {
    propelAccountList : propelAccountsArray,
    propelAccountObj  : propelAccountsObj

};