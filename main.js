const allowedFileTypes = ['application/csv',
  'application/x-csv',
  'text/csv',
  'text/comma-separated-values',
  'text/x-comma-separated-values',
  'text/tab-separated-values',
  'text/plain',
  'application/vnd.ms-excel',
  'text/x-csv']

const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

//Changes integer into a letter corresponding to a column
function numToCol(int) {
  if (int < 26) {
    return alphabet.charAt(int)
  } else {
    return alphabet.charAt(Math.floor(int/26) - 1) + alphabet.charAt(int%26)
  }
}

let createNewSheet = function(token, title, addTitle) {
  //Create new spreadsheet and return id of it
  let init = {
    method: 'POST',
    async: true,
    headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'text/csv'
    },
    'contentType': 'json'
  };  
  if (addTitle) {
    init['body'] = JSON.stringify({
      'properties': {
        'title': title
      }
    })
  }

  return fetch(
    'https://sheets.googleapis.com/v4/spreadsheets',
    init)
    .then((response) => response.json())
    .then(function(data) { 
        return data['spreadsheetId'];
        })
    .catch(function(error) {
      console.log('Error with creating sheet: ' + error.message);
      document.getElementById("imageHolder").src = "images/error.png";
      document.getElementById("message").textContent= "There was an unexpected error. Please try again.";
    })
};

let writeValuesToSheet = function(token,spreadsheetId,data,start_row) {
  //Writes a simple list of values to the sheet
  if (document.getElementById('convertInput').checked) {
    var inputOption = "USER_ENTERED"
  } else {
    var inputOption = "RAW"
  }

  let request = {
    method: 'POST',
    async: true,
    headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
    },
    'contentType': 'json',
    body:  JSON.stringify({
        "valueInputOption": inputOption,
        "data": [
           {
            "range": "Sheet1!A" + start_row + ":" + numToCol(data[0].length -1),
            "values": data
          }
        ]
      })
  };
  return fetch(
    'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values:batchUpdate',
    request)
    .then((response) => response.json())
    .then(function() {
      let sheet_url = 'https://docs.google.com/spreadsheets/d/'+spreadsheetId;
      return sheet_url;
      }
    )
    .catch(function(error) {
      console.log('Error: ' + error.message);
      document.getElementById("imageHolder").src = "images/error.png";
      document.getElementById("message").textContent= "There was an issue with your file. Please try again or choose a different file.";
      }
    );
}

let formatSheet = function(token, spreadsheetId,start_row, end_column) {
  console.log('Formatting' + spreadsheetId);
  let formatting = {
    method: 'POST',
    async: true,
    headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
    },
    'contentType': 'json',

    body: JSON.stringify({
        "requests": [
          {
            "updateSheetProperties": {
              "fields": "gridProperties.frozenRowCount",
              "properties": {
                "gridProperties": {
                  "frozenRowCount": start_row
                },
                "sheetId": 0
              }
            }
          },{
             "updateCells": {
                  "rows": [
                    {
                      "values": Array(end_column).fill(
                        {
                          "userEnteredFormat": {
                            "wrapStrategy": "WRAP"
                          }
                        })
                    }   
                  ],
            "fields": "userEnteredFormat.wrapStrategy",
            "start": {
              "sheetId": 0,
                "rowIndex": start_row - 1,
                "columnIndex": 0
                  }
                }
        }
          
        ]
    })
  };  

  try {
    response = fetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + ':batchUpdate',
      formatting);
  } catch (error) {
    console.log('Error with formatting sheet: ' + error.message);
  }
}

let profileDataRequests = function(token, spreadsheetId, end_row, end_column) {
  //Sends formating and formula values to Google Sheets API in order to highlight nulls and repeats in the data
  console.log('Profiling Data ' + spreadsheetId);
  uniqueFormula = "=IFERROR(FLOOR(COUNTUNIQUE(A4:A" + end_row + ")/COUNTA(A4:A" + end_row + "),0.01),0)"
  notNullFormula = "=IFERROR(FLOOR(COUNTA(A4:A" + end_row + ")/ROWS(A4:A" + end_row + "),0.01),0)"
  requests = [
    {
      "repeatCell": {
            "cell": 
              {
                "userEnteredValue": {
                  "formulaValue": uniqueFormula
                },
                "userEnteredFormat": {
                  "numberFormat": {
                    "pattern": "#0.#%",
                    "type": "NUMBER"
                  }
                }
              },
            "range": {
              "sheetId": 0,
              "startRowIndex": 0,
              "endRowIndex": 1,
              "startColumnIndex": 0,
              "endColumnIndex": end_column
            }, "fields":"*"
            
          }
    },
    {
            "repeatCell": {
                  "cell":
              {
                "userEnteredValue": {
                  "formulaValue": notNullFormula
                },
                "userEnteredFormat": {
                  "numberFormat": {
                    "pattern": "#0.#%",
                    "type": "NUMBER"
                  }
                }
              },
              "range": {
                "sheetId": 0,
                "startRowIndex": 1,
                "endRowIndex": 2,
                "startColumnIndex": 0,
                "endColumnIndex": end_column
              }, "fields":"*" 
          }
    },
    {
      "addConditionalFormatRule": {
        "rule": {
          "booleanRule": {
            "condition": {
              "type": "BLANK"
            },
            "format": {
              "backgroundColor": {
                "red": 0.867,
                "green": 1,
                "blue": 0.976,
                "alpha": 1
              }
            }
          },
          "ranges": [
            {
              "sheetId": 0,
              "startColumnIndex": 0,
              "endColumnIndex": end_column,
              "startRowIndex": 3,
              "endRowIndex": end_row
            }
          ]
        }
      }},
      {
        "addConditionalFormatRule": {
          "rule": {
            "booleanRule": {
              "condition": {
                "type": "CUSTOM_FORMULA",
                "values": [
                  {
                    "userEnteredValue": "=COUNTIFS(A$4:A$" + end_row + ",A4)>1"
                  }]
              },
              "format": {
                "backgroundColor": {
                "red": 1,
                "green": 0.855,
                "blue": 0.78,
                "alpha": 1
                }
              }
            },
            "ranges": [
              {
                "sheetId": 0,
                "startColumnIndex": 0,
                "endColumnIndex": end_column,
                "startRowIndex": 3,
                "endRowIndex": end_row
              }
            ]
          }
        } }     
  ]
  if (end_column < 26) {
    requests.push({
      "updateCells": {
        "range": {
          "sheetId": 0,
          "startRowIndex": 0,
          "endRowIndex": 2,
          "startColumnIndex": end_column,
          "endColumnIndex": end_column+2
        },
        "rows": [
          {
            "values": [
              {
                "userEnteredValue": {
                  "stringValue": "Percent Unique"
                },
              },
              {
                "userEnteredValue": {
                  "stringValue": "Indicates Repeated Value in Column"
                },
                "userEnteredFormat":{
                  "backgroundColor": {
                  "red": 1,
                  "green": 0.855,
                  "blue": 0.78,
                  "alpha": 1
                  }}
              }
            ]
          },
          {
            "values": [
              {
                "userEnteredValue": {
                  "stringValue": "Percent Filled"
                }
              },
              {
                "userEnteredValue": {
                  "stringValue": "Indicates Null Value in Column"
                },
                "userEnteredFormat":{"backgroundColor": {
                "red": 0.867,
                "green": 1,
                "blue": 0.976,
                "alpha": 1
                }}
              }
            ]
          }
        ],
        "fields": "*"
      }
    })
  }
  let formatting = {
    method: 'POST',
    async: true,
    headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
    },
    'contentType': 'json',
    body: JSON.stringify(
      {
        "requests": requests
      }
    )
  };  

  try {
    response = fetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + ':batchUpdate',
      formatting);
  } catch (error) {
    console.log('Error with formatting sheet: ' + error.message);
  }
}

let csvToSheet = function(data, title, prettifyColumns, profileData, addTitle) {
    chrome.identity.getAuthToken({interactive: true}, function(token) {
    
      createNewSheet(token, title, addTitle).then(function(spreadsheetId) {
        end_column = data[0].length
        if (profileData) {
          start_row = 3
          end_row = data.length + 2
          console.log(data[2])
          console.log(data[1])
          console.log(data[0])
          console.log(end_row)
        } else {
          start_row = 1
        };
        if (prettifyColumns) {
        formatSheet(token, spreadsheetId, start_row, end_column)
        for (let i = 0; i < data[0].length; i++) {
          if (typeof(data[0][i]) == 'string'){
            data[0][i] = data[0][i].replaceAll('_', ' ')
          
          split_into_words = data[0][i].split(' ')
          new_title = ''
          for (let j = 0; j < split_into_words.length; j++) {
            new_title = new_title + split_into_words[j].charAt(0).toUpperCase() + split_into_words[j].slice(1).toLowerCase()
            if (j + 1 < split_into_words.length) {
              new_title = new_title + ' '
            }
          }
          data[0][i] = new_title
          }
        }
        };

      writeValuesToSheet(token,spreadsheetId,data,start_row)
      .then(function(sheet_url){
        if (sheet_url) {
          console.log(sheet_url);
          console.log('complete');
          window.location.href = sheet_url;
          if (profileData) {profileDataRequests(token,spreadsheetId,end_row,end_column)};
        }
        else {
          document.getElementById("imageHolder").src = "images/error.png";
          document.getElementById('dropContainer').style.background = '#ffeded';
          document.getElementById("message").textContent= "An error occurred. Please try again or choose a different file.";
        }
      });
    })
    });
};

dropContainer.ondragover = dropContainer.ondragenter = function(evt) {
  evt.stopPropagation();
  evt.preventDefault();
  document.getElementById('dropContainer').style.background = '#eeeeee';
};


dropContainer.ondrop = function(evt) {
  evt.stopPropagation();
  evt.preventDefault();
  fileInput.files = evt.dataTransfer.files;
  fileInput.onchange();
};

fileInput.onchange = function() {
  var validType = false;
  if (fileInput.files){
    for (var i = 0; i < fileInput.files.length; i++) {
      if (fileInput.files[i]) {
        file = fileInput.files[i]
        if (allowedFileTypes.includes(file.type)) {
          document.getElementById('dropContainer').style.background = '#e5ffb0';
          validType = true;
          document.getElementById("imageHolder").src = "images/checkmark.png";
          document.getElementById("message").textContent= "A sheet will open in a few moments";
          break;
        } else {
          document.getElementById('dropContainer').style.background = '#ffeded';
          document.getElementById("imageHolder").src = "images/error.png";
          document.getElementById("message").textContent= "Please choose a *.csv file";
        }
      }
    }
    if (!validType) {
      return;
    }
    let fieldDelimiterInput = String(document.getElementById('fieldDelimiterInput').value);
    let recordDelimiterInput = String(document.getElementById('recordDelimiterInput').value);
    let nullsInput = String(document.getElementById('nullsInput').value);
    let fieldEnclosedInput = String(document.getElementById('fieldEnclosedInput').value);
    let escapeInput = String(document.getElementById('escapeInput').value);
    let convertInput = document.getElementById('convertInput').checked;
    let addTitle = document.getElementById('addTitle').checked;
    let prettifyColumns = document.getElementById('prettifyColumns').checked;
    let profileData = document.getElementById('profileData').checked;
    let skipEmpty = document.getElementById('skipEmpty').checked;
    getFileData(file, fieldDelimiterInput, 
                recordDelimiterInput, nullsInput, 
                fieldEnclosedInput, escapeInput, convertInput,
                prettifyColumns, profileData, addTitle, skipEmpty)
  };
  
};

function convertNulls(nullVal){
  return function(val){
    if(val==nullVal){
      return ''
    }else{
      return val
    }
  }
}

let getFileData = function(file, fieldDelimiterInput, recordDelimiterInput,
                           nullsInput, fieldEnclosedInput, escapeInput, convertInput,
                           prettifyColumns, profileData, addTitle, skipEmpty) {
  //calls this function when done
  let title = file.name
  let params = {complete: function (parsed) {csvToSheet(parsed.data, title, prettifyColumns, profileData, addTitle)}}
  
  //other params for PapaParse
  if (fieldDelimiterInput != 'auto-detect'){
    params['delimiter'] = fieldDelimiterInput
  }
  if (recordDelimiterInput != 'auto-detect'){
    params['newline'] = recordDelimiterInput
  }
  if (nullsInput != ''){
    params['transform'] = convertNulls(nullsInput)
  }
  if (fieldEnclosedInput != '"'){
    params['quoteChar'] = fieldEnclosedInput
  }
  if (escapeInput != '"'){
    params['escapeChar'] = escapeInput
  }
  if (convertInput){
    params['dynamicTyping'] = true
  }
  if (skipEmpty){
    params['skipEmptyLines'] = true
  }
  Papa.parse(file, params)
}
