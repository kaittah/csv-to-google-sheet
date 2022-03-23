// Saves options to chrome.storage
function save_options() {
    var fieldDelimiterInput = String(document.getElementById('fieldDelimiterInput').value);
    var recordDelimiterInput = String(document.getElementById('recordDelimiterInput').value);
    var nullsInput = String(document.getElementById('nullsInput').value);
    var fieldEnclosedInput = String(document.getElementById('fieldEnclosedInput').value);
    var escapeInput = String(document.getElementById('escapeInput').value);
    var convertInput = document.getElementById('convertInput').checked;
    var addTitle = document.getElementById('addTitle').checked;
    var prettifyColumns = document.getElementById('prettifyColumns').checked;
    var profileData = document.getElementById('profileData').checked;
    var skipEmpty = document.getElementById('skipEmpty').checked;
    chrome.storage.sync.set({
        fieldDelimiterInput : fieldDelimiterInput,
        recordDelimiterInput : recordDelimiterInput,
        nullsInput : nullsInput,
        fieldEnclosedInput : fieldEnclosedInput,
        escapeInput : escapeInput,
        convertInput : convertInput,
        addTitle: addTitle,
        prettifyColumns: prettifyColumns,
        profileData: profileData,
        skipEmpty: skipEmpty
    }, function() {
      // Update status to let user know options were saved.
      var status = document.getElementById('status');
      status.textContent = 'Options saved.';
      setTimeout(function() {
        status.textContent = '';
      }, 750);
    });
  }
  
// Restores select box and checkbox state using the preferences
// stored in chrome.storage.
function restore_options() {
chrome.storage.sync.get({
    fieldDelimiterInput : 'auto-detect',
    recordDelimiterInput : 'auto-detect',
    nullsInput : '',
    fieldEnclosedInput : '"',
    escapeInput : '"',
    convertInput :true,
    addTitle: false,
    prettifyColumns: false,
    profileData: false,
    skipEmpty: true
}, function(items) {
    document.getElementById('fieldDelimiterInput').value = items.fieldDelimiterInput;
    document.getElementById('recordDelimiterInput').value = items.recordDelimiterInput;
    document.getElementById('nullsInput').value = items.nullsInput;
    document.getElementById('fieldEnclosedInput').value = items.fieldEnclosedInput;
    document.getElementById('escapeInput').value = items.escapeInput;
    document.getElementById('convertInput').checked = items.convertInput;
    document.getElementById('addTitle').checked = items.addTitle;
    document.getElementById('prettifyColumns').checked = items.prettifyColumns;
    document.getElementById('profileData').checked = items.profileData;
    document.getElementById('skipEmpty').checked = items.skipEmpty

});
}

document.addEventListener('DOMContentLoaded', restore_options);
document.getElementById('save').addEventListener('click',
    save_options);