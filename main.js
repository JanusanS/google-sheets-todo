const app = SpreadsheetApp.getActiveSpreadsheet(); // app
const lock = LockService.getScriptLock();
try {
  var ui = SpreadsheetApp.getUi();

} catch { }
var today = new Date();
today.setHours(0, 0, 0, 0);
const sheetslist = ['Today', 'Weekly', 'Monthly', 'All', 'Remove', 'Edit', 'Completed'];
const colorslist1 = [['Red', '#FFCFC9'], ['Orange', '#ffc8aa'], ['Yellow', '#ffe5a0'], ['Green', '#d4edbc'], ['Blue', '#bfe1f6'], ['Cyan', '#c6dbe1'], ['Purple', '#e6cff2'], ['Silver', '#e8eaed']];
const colorslist2 = [['Gray', '#3d3d3d'], ['Red2', '#b10202'], ['Brown', '#753800'], ['Blue3', '#473822'], ['Green2', '#11734b'], ['Blue2', '#0a53a8'], ['BlueGreen', '#215a6c'], ['Purple', '#5a3286']];
const week = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const user = PropertiesService.getUserProperties();
//setprop('silenced',false);

function open() { // triggered when sheet is opened
  menu();
  if (!getprop("Setup")) { // checks if the spreadsheet is new
    setup();
  }
}

function resetUser() {
  try {
    var confirmation = ui.alert("Are you sure you want to reset all data", "Reset User", ui.ButtonSet.YES_NO);
    if (confirmation === ui.Button.YES) {
      user.deleteAllProperties();
      toast("All user properties have been deleted");
      setup();
    }
  } catch {
    user.deleteAllProperties();
    toast("All user properties have been deleted");
    setup();
  }
}
function reloadApp() {
  copy('Setting Up')
  var sheets = app.getSheets();
  for (let x = 0; x < sheets.length - 1; x++) {
    app.deleteSheet(sheets[x]);
  }
  copy('Add');
  copy('Remove');
  copy('Edit');
  copy('Today');
  copy('Weekly');
  copy('Monthly');
  copy('All');
  copy('Completed')
  try {
    for (let topic of getprop('topics')) {
      copy('Class', topic);
      fillSheet(topic, getprop(topic));
    }
  }
  catch { }
  copy('Raw').hideSheet();
  copy('Raw Settings').hideSheet();
  updateTopicDropdown();
  updateKeywordDropdown();
  addColors();
  updateColors();
  updateAll();
  menu();
  app.deleteSheet(getSheet('Setting Up'));
  alert("App has been reloaded");
}
function menu() {
  try {
    var functions = ui.createMenu("Functions");
    var settings = ui.createMenu("Settings");
    functions.addItem("Reload App", 'reloadApp');
    functions.addItem("Reset User Data", 'resetUser');
    functions.addToUi();
    settings.addToUi();
  }
  catch { }
}

function setup() {
  toast("Setting Up SpreadSheet");
  copy('Setting Up');
  var sheets = app.getSheets();
  for (let x = 0; x < sheets.length - 1; x++) {
    app.deleteSheet(sheets[x]);
  }
  copy('Add');
  copy('Remove');
  copy('Edit');
  copy('Today');
  copy('Weekly');
  copy('Monthly');
  copy('Completed');
  copy('All');
    copy('Raw').hideSheet();
  copy('Raw Settings').hideSheet();

  app.deleteSheet(getSheet("Setting Up"));
  setprop('Setup', true);
  setprop('items');
  setprop('topics');
  setprop('topicsC');
  setprop('keywords');
  setprop('keywordsC');
  setprop('completed');
  user.setProperty('id', 0);
  setprop('silenced', false);

  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger("open")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()) // Bind to the current spreadsheet
    .onOpen()
    .create();

  ScriptApp.newTrigger("edit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()) // Bind to the current spreadsheet
    .onEdit()
    .create();

  ScriptApp.newTrigger("update")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
  alert("Welcome!\nStart by adding a topic for your items")
}
function readSetup() {
  user.deleteAllProperties();
  toast("Reseting and Reading");
  copy('Setting Up');
  var sheets = app.getSheets();
  for (let x = 0; x < sheets.length - 1; x++) {
    if (!sheets[x].getName().startsWith("Raw")) {
      app.deleteSheet(sheets[x]);
    }
  }
  copy('Add');
  copy('Remove');
  copy('Edit');
  copy('Today');
  copy('Weekly');
  copy('Monthly');
  copy('Completed');
  copy('All');

  app.deleteSheet(getSheet("Setting Up"));
  setprop('Setup', true);
  setprop('items');
  setprop('topics');
  setprop('topicsC');
  setprop('keywords');
  setprop('keywordsC');
  setprop('completed');
  user.setProperty('id', 0);
  setprop('silenced', false);

  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger("open")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()) // Bind to the current spreadsheet
    .onOpen()
    .create();

  ScriptApp.newTrigger("edit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()) // Bind to the current spreadsheet
    .onEdit()
    .create();

  ScriptApp.newTrigger("update")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
}
function readData() { // makes an array of all items in data // figure out if this is needed
  readSetup();
  setprop('silenced', true);
  var settings = getSheet('Raw Settings');
  var data = getSheet('Raw');
  var topics = settings.getRange("A3:B").getValues();
  var keywords = settings.getRange("C3:D").getValues();
  var temptopics = getprop('topics');
  var topcolors = getprop('topicsC');
  for (let x of topics) {
    if (x[0] == "" || x[1] == ""){
      break;
    }
    temptopics = temptopics.concat(x[0]);
    topcolors = topcolors.concat(x[1]);
    Logger.log(x)
    copy('Class', x[0]);
    setprop(x[0]);
  }
  setprop('topics', temptopics);
  setprop('topicsC', topcolors);
  updateTopicDropdown();
  var tempkeywords = getprop('keywords');
  var keycolors = getprop('keywordsC');
  for (let y of keywords) {
        if (y[0] == "" || y[1] == ""){
      break;
    }
    tempkeywords = tempkeywords.concat(y[0]);
    keycolors = keycolors.concat(y[1]);

  }
  setprop('keywords', tempkeywords);
  setprop('keywordsC', keycolors);
  updateKeywordDropdown();
  updateColors();
  var headers = data.getRange(2, 1, 1, 7).getValues()[0];
  var row;
  var count = 3;
  var tempitems = getprop('items');
  while (data.getRange( count,2).getValue() !== "") {
    var item = {}; // empty object
    row = data.getRange(count, 1, 1, 7).getValues()[0];
    for (let y = 0; y < headers.length; y++) {
      if (y == 3 && !(typeof row[3] === "string")){
            var date = new Date(row[3]);
    item.date = date.toLocaleDateString('en-CA');
      }else{
        item[headers[y].toLowerCase()] = row[y]; // fills object
      }
      
    }
    var id = parseInt(user.getProperty('id')) + 1;
    user.setProperty('id', id);
    item.id = id;
    tempitems.push(item); // adds object to list
    count = count +1;
    Logger.log(item);
  }
  setprop('items', tempitems);
  var filteredItems = tempitems.filter(function (item) {
    return item.completion == true;
  });
  setprop('completed',filteredItems);
  groupTopics();
  debounce("updateAll");
}

function groupTopics() {
  var items = getprop('items');
  var topics = getprop('topics');
  for (let x of topics) {
    var filteredItems = items.filter(function (item) {
      return item.topic == x;
    });
    setprop(x, filteredItems);
    fillSheet(x, getprop(x));
  }
  
}

function update() { // update weekly and monthly depending on date
  fillSheet("Today", dateFilter(1))
  fillSheet("Weekly", dateFilter(2));
  fillSheet("Monthly", dateFilter(3));
  clearCompletion("Daily");
  if (today.getDay() == 1) {
    clearCompletion("Week");
  }
  if (today.getDate() == 1) {
    clearCompletion("Month")
  }
  toast("Daily Update");
}

function dateFilter(range, items = getprop("items")) {
  var month = today.getUTCMonth();
  var year = today.getFullYear();
  var day = today.getDay();
  if (range == 1) { //today
    var filteredItems = items.filter(function (item) {
      var itemDate = new Date(item.date);
      return itemDate.getUTCDate() == today.getUTCDate() || item.date == "Daily" || endsWithAny(item.date, week);
    });
  }
  if (range == 2) { // week
    var weekstart = new Date(today);
    weekstart.setDate(today.getDate() - (day - 1));
    var weekend = new Date(weekstart);
    weekend.setDate(weekstart.getDate() + 7);
    var filteredItems = items.filter(function (item) {
      var itemDate = new Date(item.date);
      return (itemDate.getUTCDate() >= weekstart.getUTCDate() && itemDate <= weekend) || item.date.startsWith("Weekly") || item.date == "Daily";
    });
  }
  if (range == 3) { // month
    var filteredItems = items.filter(function (item) {
      var itemDate = new Date(item.date);
      return (itemDate.getUTCMonth() === month && itemDate.getFullYear() === year) || item.date == "Monthly" || item.date.startsWith("Weekly") || item.date == "Daily";
    });
  }
  return filteredItems;
}


function fillSheet(sheetname, fillItems) { // fills a sheet with items
  sheet = getSheet(sheetname);
  //sheet.activate();
  var xplus = sheet.getName() === "Edit" || sheet.getName() === "Remove" ? 11 : 3;

  if (fillItems === null || fillItems.length === 0) { // empty array
    clearSheet(sheet, xplus);
    return;
  }

  var columns = sheet.getName() === "Edit" || sheet.getName() === "Remove" ? 8 : 7;
  var fill = []
  fillItems = sortitems(fillItems);
  clearSheet(sheet, xplus);


  for (let x = 0; x < fillItems.length; x++) {
    let fillObject = Object.values(fillItems[x]);
    if (sheet.getName() === "Edit" || sheet.getName() === "Remove") {
      let beforeJ = fillObject.slice(0, 6);
      let afterJ = fillObject.slice(6);
      fillObject = beforeJ.concat([false]).concat(afterJ);
    }
    fill.push(fillObject);
  }
  sheet.getRange(xplus, 1, fillItems.length, columns).setValues(fill);
  toast("Updated " + sheetname);
}

function updateAll(msg = "") {
  fillSheet("Today", dateFilter(1));
  fillSheet("Weekly", dateFilter(2));
  fillSheet("Monthly", dateFilter(3));
  fillSheet("All", getprop('items'));
  fillSheet("Completed", getprop('completed'));
  fillSheet("Remove", getprop('items'));
  fillSheet("Edit", getprop('items'));

  for (let topic of getprop('topics')) {
    fillSheet(topic, getprop(topic));
  }
  toast("Updated All");
  if (msg) {
    alert(msg);
  }
}
function updateAllItem(topic, id, items = getprop("items")) {
  updateItem("Today", dateFilter(1, items), id);
  updateItem("Weekly", dateFilter(2, items), id);
  updateItem("Monthly", dateFilter(3, items), id);
  updateItem("All", items, id);
  updateItem("Remove", items, id);
  updateItem("Edit", items, id);
  updateItem(topic, getprop(topic), id);
  toast("Updated All");
}


function addItem() {
  var add = getSheet('Add');
  add.getRange('C12').setValue(false);
  var headers = ["completion", "item", "topic", "date", "location", "priority", "id"];
  var item = {};

  var column = getSheet('Add').getRange(4, 3, 7, 1).getValues();
  var reoccuring = column[4].toString();
  if (column[0] === null || column[0] === "" || column[0] == []) {

    alert("Item Field Empty");
    return;
  }
  if (column[1][0] === null || column[1][0] === "") {
    alert("Class/Topic Field Empty");
    return;
  }

  item[headers[0]] = false;
  item[headers[1]] = column[0][0];
  //add.getRange(4, 3).setValue("");
  item[headers[2]] = column[1];
  //add.getRange(5, 3).setValue("");
  if (reoccuring.startsWith("Indefinitely")) {
    var weekday = "";
    if (reoccuring.endsWith("Weekly") && (!(column[3] === null) || !(column[3] === ""))) {
      weekday = " - " + week[(new Date(column[3])).getDay()];
    }
    item.date = reoccuring.slice(13) + weekday;
  }
  else if (column[4] === null || column[4] === "") {
    item.date = "";
  }
  else {
    var date = new Date(column[3]);
    item.date = date.toLocaleDateString('en-CA');
  }
  add.getRange(7, 3).setValue(today);
  item[headers[4]] = column[2];
  //add.getRange(6, 3).setValue("");
  item[headers[5]] = column[6];
  //add.getRange(10, 3).setValue("");
  if (column[4] != "" && !reoccuring.startsWith("Indefinitely")) {

    addReoccuringItem(item, column[3], column[5], column[4]);
    return;
  }
  var id = parseInt(user.getProperty('id')) + 1;
  user.setProperty('id', id);
  item[headers[6]] = id;
  var tempitems = getprop('items').concat(item);
  setprop('items', tempitems);
  toast(item);
  //setprop('items',getprop('items').concat(item));
  debounce("updateAll");
  updateTopic(item);
  fillSheet(item.topic, getprop(item.topic));
  alert("Item has been added");
}


function addReoccuringItem(item, start, end, frequency) { // adds reoccuring items
  var startd = new Date(start);
  //startd = startd.toLocaleDateString('en-CA');
  var endd = new Date(end);
  var day = startd.getDate();
  //endd = endd.toLocaleDateString('en-CA');
  var name = item.item;
  var count = 1;
  var regex = /(.*?)(\d+)$/; // checks to see if string name ends with number ex. quiz 4
  var match = name.match(regex);
  if (match) {
    name = match[1].trim(); // Everything before the number
    count = parseInt(match[2], 10); // The number at the end
  }

  while (startd <= endd) {

    item.item = name + " " + count;
    item.date = startd.toLocaleDateString('en-CA');
    var id = parseInt(user.getProperty('id')) + 1;
    user.setProperty("id", id);
    item.id = id;
    var tempitems = getprop('items').concat(item);
    setprop('items', tempitems);
    updateTopic(item);
    toast(item);
    startd = nextDate(startd, frequency[0], day);

    count++;

  }

  fillSheet(item.topic, getprop(item.topic));
  debounce("updateAll");
  alert("Items have been added")
}

function isCellHyperlink(cellAddress) { // think about it
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange(cellAddress); // Specify the cell by its address, e.g., "A1"

  // Get the rich text value of the cell
  var richTextValue = cell.getRichTextValue();

  if (richTextValue) {
    var runs = richTextValue.getRuns(); // Get the text runs in the cell
    for (var i = 0; i < runs.length; i++) {
      var link = runs[i].getLinkUrl(); // Get the link URL for the text run
      if (link) {
        Logger.log("Cell " + cellAddress + " contains a hyperlink: " + link);
        return true; // Return true if at least one link is found
      }
    }
  }

  Logger.log("Cell " + cellAddress + " does not contain any hyperlinks.");
  return false; // Return false if no links are found
}

// Example usage
function testHyperlinkCheck() {
  var isLink = isCellHyperlink("A1"); // Check if cell A1 contains a hyperlink
  Logger.log("Is cell A1 a hyperlink? " + isLink);
}


function makeItem(sheetName, rowNum) {
  sheet = getSheet(sheetName);
  var headers = ["completion", "item", "topic", "date", "location", "priority", "id"];
  var item = {};
  var row = sheet.getRange(rowNum, 1, 1, 6).getValues()[0];
  if (row[1] === null || row[1] === "" || row[1] == []) {

    alert("Item Field Empty");
    return -1;
  }
  if (row[2] === null || row[2] === "") {
    alert("Class/Topic Field Empty");
    return -1;
  }


  item[headers[0]] = row[0];
  item[headers[1]] = row[1];
  item[headers[2]] = row[2];
  if (typeof row[3] === "string") {
    item.date = row[3]
  }
  else if (row[3] === null || row[3] === "") {
    item.date = "";
  }
  else {
    var date = new Date(row[3]);
    item.date = date.toLocaleDateString('en-CA');
  }

  item[headers[4]] = row[4];
  item[headers[5]] = row[5];
  return item;

}

function editItem(row) {
  var edit = getSheet('Edit');
  //edit.getRange('C7').setValue(false);
  var topicchange = new Set();
  var datechange = false;
  var id = edit.getRange(row, 8).getValue();
  var editedItem = makeItem('Edit', row);
  if (editedItem === -1) {
    return;
  }
  editedItem.id = id;
  for (let x = 0; x < 5; x++) {
    if (lock.tryLock(5000)) { // lock to make sure property isn't accessed multiple times at once
      try {
        var items = getprop('items');
        var originalItem = items.find(item => item.id === id); // Finding item with id


        if (originalItem.date !== editedItem.date) {
          datechange = true;
          topicchange.add(originalItem.topic); // maybe redundatn
        }


        Object.assign(originalItem, editedItem); // reassigning item properties with new properties
        setprop('items', items);
        if (originalItem.topic !== editedItem.topic) { // moving item to new topic
          var filteredTopic = getprop(originalItem.topic).filter(function (item) {
            return item.id !== id;
          });
          setprop(originalItem.topic, filteredTopic);
          updateTopic(editedItem);
          topicchange.add(originalItem.topic);
          topicchange.add(editedItem.topic);
        }
        else { //
          var topics = getprop(originalItem.topic);
          var topicItem = topics.find(topic => topic.id === id);
          Object.assign(topicItem, editedItem);
          setprop(originalItem.topic, topics);

        }
      }
      finally {
        lock.releaseLock();


        if (!datechange) {
          updateAllItem(editedItem.topic, id); // maybe redundant
        }
        for (let topic of topicchange) { // redundant maybe
          fillSheet(topic, getprop(topic));
        }
        if (datechange) {
          debounce("updateAll");
        }
        setprop("toedit");
        alert("Item(s) Edited");
        break;
      }
    }
    Utilities.sleep(2000);
  }


}

function remItem(row) {
  var remove = getSheet('Remove');
  var itemname = remove.getRange(row, 2).getValue();
  var confirmation = alert("Are you sure you want to delete " + itemname, "Confirm Delete", ui.ButtonSet.YES_NO);
  if (confirmation === ui.Button.NO) {
    return;
  }
  var id = remove.getRange(row, 8).getValue();
  //alert("id: " + id);
  var topic = remove.getRange(row, 3).getValue();
  //remove.getRange('C7').setValue(false);
  for (let x = 0; x < 5; x++) {
    if (lock.tryLock(5000)) { // lock to make sure property isn't accessed multiple times at once
      try {
        var filteredItems = getprop("items").filter(function (item) {
          return item.id !== id;
        });

        var filteredTopic = getprop(topic).filter(function (item) {
          return item.id !== id;
        });

        var filteredCompleted = getprop("completed").filter(function (item) {
          return item.id !== id;
        });

        setprop("items", filteredItems);
        setprop("completed", filteredCompleted);
        setprop(topic, filteredTopic);
      }
      finally {
        lock.releaseLock();
        //scheduleTrigger("update", "updateAll");
        debounce("updateAll", 5000);
        alert("Item(s) Removed");
        break;
      }
    }
    Utilities.sleep(2000);
  }
}

function searchItems(sheetname) {
  var sheet = getSheet(sheetname)
  sheet.getRange('C7').setValue(false);
  var items = getprop('items');
  var item = sheet.getRange('C4').getValue().toString();
  var topic = sheet.getRange('C5').getValue().toString();
  if (topic !== null && topic !== "") {
    items = getprop(topic);
  }
  if (item !== null && item !== "") {

    item = item.toLowerCase();
    for (let i = items.length - 1; i >= 0; i--) {

      if (!(items[i].item).toLowerCase().includes(item)) {
        items.splice(i, 1); // Remove the item from the array
      }
    }
  }

  updateColors();
  fillSheet(sheetname, items);

  alert('SEARCHED');
}



function nextDate(date, frequency, day) {

  switch (frequency) {
    case "Daily":

      date.setDate(date.getDate() + 1);

      return date;
    case "Weekly":
      date.setDate(date.getDate() + 7);
      return date;
    case "Monthly":

      date.setMonth(date.getMonth() + 1);
      if (date.getDate() < day) { // if month has less days then previous
        date.setDate(day);
      }
      return date;
    case "Yearly":
      date.setYear(date.getYear() + 1);
      return date;
  }
}


function clearSheet(sheet, xplus = 3) {
  var start = xplus;

  var end = sheet.getLastRow();
  rows = end - start;
  var range = sheet.getRange("B" + xplus + ":F" + end);
  range.clearContent();
  //range.clearDataValidations();
  if (rows > 0) {
    sheet.deleteRows(start, rows);
    //sheet.getRange(start, 1, 1, sheet.getLastColumn()).clear();
  }


}



function updateTopic(item) {
  var topic = item.topic;

  var temptopic = getprop(topic).concat(item);
  setprop(topic, temptopic);

}

function findTopic(item) {
  for (let x = 0; x < getprop('topics').length; x++) {
    if (getprop('topics')[x] === item.topic) {

      return getprop('topics')[x];
    }
  }
}

function addTopic() {

  getSheet('Add').getRange('C19').setValue(false);
  var topic = getSheet('Add').getRange('C16').getValue();
  var color = getSheet('Add').getRange('C17').getValue();
  getSheet('Add').getRange('C16').setValue("");
  getSheet('Add').getRange('C17').setValue("Red");
  if (getSheet(topic)) {
    alert("Class/Topic Exists");
    return;
  }
  if (topic === null || topic === "") {
    alert("Field Empty");
    return;
  }
  var temptopics = getprop('topics').concat(topic);
  setprop('topics', temptopics);
  var colors = getprop('topicsC').concat(getColor(color));
  setprop('topicsC', colors);
  updateTopicDropdown();
  copy('Class', topic);
  setprop(topic);
  updateColors()
  alert("Class/Topic has been added");
}

function remTopic() {
  getSheet('Remove').getRange('F6').setValue(false);
  var topic = getSheet('Remove').getRange('F4').getValue();
  getSheet('Remove').getRange('F4').setValue("");

  if (!search(getprop('topics'), topic)) {
    alert("Topic does not exist");
    return;
  }
  if (topic === null || topic === "") {
    alert("Field Empty");
    return;
  }

  var filteredItems = getprop('items').filter(function (item) {
    return String(item.topic).trim() !== String(topic).trim();
  });
  setprop('items', filteredItems);

  var filteredCompleted = getprop('completed').filter(function (item) {
    return String(item.topic).trim() !== String(topic).trim();
  });
  setprop('completed', filteredCompleted);
  remprop(topic);
  app.deleteSheet(getSheet(topic));
  var index = getprop('topics').indexOf(topic);
  var temptopics = getprop('topics')
  temptopics.splice(index, 1);
  setprop('topics', temptopics);
  var colors = getprop('topicsC');
  colors.splice(index, 1)
  setprop('topicsC', colors);
  updateColors();
  updateTopicDropdown();
  debounce("updateAll");
  alert("Topic/Class has been removed");
}

function editTopic() {
  var sheet = getSheet('Edit')
  sheet.getRange('F8').setValue(false);
  var topic = sheet.getRange('F4').getValue();
  var newtopic = sheet.getRange('F5').getValue();
  var newcolor = sheet.getRange('F6').getValue();
  sheet.getRange('F4').setValue("");
  //sheet.getRange('F5').setValue("");

  if (!search(getprop('topics'), topic)) {
    alert("Topic does not exist");
    return;
  }

  if (search(getprop('topics'), newtopic)) {
    alert("New Topic Exists");
    return;
  }
  if (topic === null || topic === "") {
    alert("Field Empty");
    return;
  }

  var index = getprop('topics').indexOf(topic);
  var temptopics = getprop('topics')

  //temptopics.splice(index, 1);
  var colors = getprop('topicsC');
  //colors.splice(index, 1)
  var temptopic = getprop(topic);


  if (newtopic !== null && newtopic !== "") { // Case replacing old keyword with new keyword
    temptopics[index] = newtopic;

    for (let x = 0; x < temptopic.length; x++) {
      temptopic[x].topic = newtopic;

    }
    remprop(topic);
    setprop(newtopic, temptopic);
    setprop('topics', temptopics);

    var tempitems = [];
    for (let x = 0; x < temptopics.length; x++) {

      tempitems = tempitems.concat(getprop(temptopics[x]));
    }
    setprop('items', tempitems);

    if (newcolor !== null && newcolor !== "") { // Case color change

      colors[index] = getColor(newcolor);
    }
    var topicsheet = getSheet(topic);
    topicsheet.setName(newtopic);

    fillSheet(newtopic, getprop(newtopic));
  }
  else { // Case same keyword new color
    if (newcolor !== null && newcolor !== "") { // if new color exists
      colors[index] = getColor(newcolor);
    }

  }


  setprop('topics', temptopics);
  setprop('topicsC', colors);
  updateColors();
  updateTopicDropdown();
  debounce("updateAll");
  alert("Topic/Class has been updated");
}

function addKeyword() {
  getSheet('Add').getRange('C26').setValue(false);
  var keyword = getSheet('Add').getRange('C23').getValue();
  var color = getSheet('Add').getRange('C24').getValue();
  getSheet('Add').getRange('C23').setValue("");
  getSheet('Add').getRange('C24').setValue("Red");
  if (search(getprop('keywords'), keyword)) {
    alert("Keyword already exists");
    return;
  }
  if (keyword === null || keyword === "") {
    alert("Field Empty");
    return;
  }
  var tempkeywords = getprop('keywords').concat(keyword);
  setprop('keywords', tempkeywords);
  var colors = getprop('keywordsC').concat(getColor(color));
  setprop('keywordsC', colors);
  updateColors();
  updateKeywordDropdown();
  alert("Keyword has been added");
}


function remKeyword() {
  getSheet('Remove').getRange('K6').setValue(false);
  var keyword = getSheet('Remove').getRange('J4').getValue();
  getSheet('Remove').getRange('J4').setValue("");

  if (!search(getprop('keywords'), keyword)) {
    alert("Keyword does not exist");
    return;
  }
  if (keyword === null || keyword === "") {
    alert("Field Empty");
    return;
  }
  var index = getprop('keywords').indexOf(keyword);
  var tempkeywords = getprop('keywords')
  tempkeywords.splice(index, 1);
  setprop('keywords', tempkeywords);
  var colors = getprop('keywordsC');
  colors.splice(index, 1)
  setprop('keywordsC', colors);
  updateColors();
  updateKeywordDropdown();
  alert("Keyword has been removed");
}

function editKeyword() {
  getSheet('Edit').getRange('K8').setValue(false);

  var keyword = getSheet('Edit').getRange('K4').getValue();
  var newkeyword = getSheet('Edit').getRange('K5').getValue();
  var newcolor = getSheet('Edit').getRange('K6').getValue();
  getSheet('Edit').getRange('K4').setValue("");
  getSheet('Edit').getRange('K5').setValue("");

  if (!search(getprop('keywords'), keyword)) {
    alert("Keyword does not exist");
    return;
  }
  if (keyword === null || keyword === "") {
    alert("Field Empty");
    return;
  }
  var tempkeywords = getprop('keywords');
  var index = tempkeywords.indexOf(keyword);
  var colors = getprop('keywordsC');

  if (newkeyword !== null && newkeyword !== "") { // Case replacing old keyword with new keyword

    tempkeywords[index] = newkeyword;
    if (newcolor !== null && newcolor !== "") { // Case color change

      colors[index] = getColor(newcolor);
    }

  }
  else { // Case same keyword new color
    if (newcolor !== null && newcolor !== "") { // if new color exists
      colors[index] = getColor(newcolor);
    }

  }
  setprop('keywords', tempkeywords);
  setprop('keywordsC', colors);

  updateColors();
  updateKeywordDropdown();
  alert("Keyword has been updated");
}

function getColor(color) {
  for (let x of colorslist1.concat(colorslist2)) {
    if (x[0] === color) {
      return x[1];
    }
  }
}

function updateDropdown(list, sheet, cell) {
  var dropdown = SpreadsheetApp.newDataValidation().requireValueInList(getprop(list));
  Logger.log(getprop(list));
  getSheet(sheet).getRange(cell).setDataValidation(dropdown);
}

function updateTopicDropdown() {

  updateDropdown('topics', 'Add', 'C5');
  updateDropdown('topics', 'Remove', 'C5');
  updateDropdown('topics', 'Remove', 'F4');
  updateDropdown('topics', 'Edit', 'C5');
  updateDropdown('topics', 'Edit', 'F4');
  updateDropdown('topics', 'Edit', 'C11:C');
  toast("Topic Dropdown's Updated");
}

function updateKeywordDropdown() {
  updateDropdown('keywords', 'Remove', 'J4');
  updateDropdown('keywords', 'Edit', 'J4');
  toast("Keyword Dropdown's Updated");
}

function markComplete(sheet, row) {
  var id = getId(sheet, row);
  var completion = sheet.getRange(row, 1).getValue();

  var items = getprop('items');
  var item = items.find(item => item.id === id);
  item.completion = completion;
  setprop("items", items);


  var topics = getprop(item.topic);
  var topicItem = topics.find(topic => topic.id === id);
  topicItem.completion = completion;
  setprop(item.topic, topics);

  updateAllItem(item.topic, id);
  //alert("Marked Complete");
  debounce("fillComplete", 4000);
}

function fillComplete() {
  var items = getprop("items");
  var completed = items.filter(function (itemx) {
    return itemx.completion === true;
  });
  setprop('completed', completed);
  Logger.log(completed);

  fillSheet("Completed", getprop('completed'));
  //ScriptApp.deleteTrigger(getprop("completeTrigger")); // delete triggers later
  remprop("completeTrigger");
}

function clearCompletion(keyword) {
  var completed = getprop("completed");
  var items = getprop('items');

  for (let x of completed) {
    if (x.date == keyword && x.completion == true) {
      x.completion = false;
      let topic = getprop(x.topic);
      let id = x.id;
      var item = items.find(item => item.id === id);
      item.completion = false;
      var topicItem = topic.find(topic => topic.id === id);
      topicItem.completion = false;
      setprop(x.topic, topic);
      updateAllItem(x.topic, id, items);
    }
  }
  setprop("items", items);
  setprop("completed", completed);
  fillComplete();
}

function updateItem(sheetName, fillItems, id) {
  var sheet = getSheet(sheetName);
  fillItems = sortitems(fillItems);
  var index = fillItems.findIndex(item => item.id === id);
  if (sheetName === "Completed") {
    toast(index);
  }
  if (index === -1) {
    return;
  }
  if (sheetName === "Edit" || sheetName === "Remove") {
    let fillObject = Object.values(fillItems[index]);
    let beforeJ = fillObject.slice(0, 6);

    let afterJ = fillObject.slice(6);
    fillObject = beforeJ.concat([false]).concat(afterJ);
    sheet.getRange(index + 11, 1, 1, 8).setValues([fillObject]);
  }
  else {
    let fillObject = Object.values(fillItems[index]);
    sheet.getRange(index + 3, 1, 1, 7).setValues([fillObject]);
  }
}

function sortitems(items) {
  items.sort(function (a, b) {
    var aDate = new Date(a.date);
    var bDate = new Date(b.date);

    var aIsDate = !isNaN(aDate.getTime()); //check if string
    var bIsDate = !isNaN(bDate.getTime());

    if (aIsDate && !bIsDate) return -1; // if one is date order it first
    if (!aIsDate && bIsDate) return 1;

    if (aIsDate && bIsDate) { // if both dates
      var dateComparison = aDate - bDate;
      if (dateComparison !== 0) {
        return dateComparison;
      }
    }
    if (!aIsDate && !bIsDate) { // if both strings
      var stringComparison = a.date.localeCompare(b.date);
      if (stringComparison !== 0) {
        return stringComparison;
      }
    }
    return a.id - b.id;
  });
  return items
}

function getId(sheet, row) {
  var column = 7;
  if (sheet.getName() === "Edit" || sheet.getName() === "Remove") {
    column = 8;
  }
  return sheet.getRange(row, column).getValue();
}
function addColors() {
  var all = sheetslist.concat(getprop('topics'));
  for (let x = 0; x < all.length; x++) {
    var sheet = getSheet(all[x]);
    Logger.log(all[x]);
    var checkboxrange = sheet.getRange("A3:A");
    try {
      var rules = sheet.getConditionalFormatRules();
    } catch {
      var rules = [];
    }
    if (sheet.getName() === "Remove" || sheet.getName() === "Edit") {
      var checkboxrange = sheet.getRange("G11:G");
      rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B11<>""')
        .setFontColor('#808080')
        .setRanges([checkboxrange])
        .build();
      rules.push(rule);
    }
    else {
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B3<>""')
        .setFontColor('#808080')
        .setRanges([checkboxrange])
        .build();
      rules.push(rule);
    }
    sheet.setConditionalFormatRules(rules);

  }
}

function updateColors() {
  var col = ["C", "B"];
  var colors = [getprop('topicsC'), getprop('keywordsC')];
  var words = [getprop('topics'), getprop('keywords')];
  var all = sheetslist.concat(getprop('topics'));
  var fontColor;
  for (let x = 0; x < all.length; x++) {
    var sheet = getSheet(all[x]);
    var rules = sheet.getConditionalFormatRules();
    for (let w = 0; w < 2; w++) {
      var range = sheet.getRange(col[w] + "1:" + col[w] + sheet.getMaxRows());

      rules = rules.filter(function (rule) { // deletes old rules in the range
        return !rule.getRanges().some(function (r) {
          return r.getA1Notation() === range.getA1Notation(); // keeps ones not in range
        });
      });

      for (let y = 0; y < colors[w].length; y++) { // makes rules
        fontColor = isdark(colors[w][y])
        var rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains(words[w][y]) // FIX FOR TOPICS whentextequals
          .setFontColor(fontColor)
          .setBackground(colors[w][y])
          .setRanges([range])
          .build();
        rules.push(rule);
      }

      sheet.setConditionalFormatRules(rules);
    }
  }
}
function isdark(color) {
  for (let x of colorslist2) {
    if (x[1] === color) {
      return "#efefef"; // Return this if the value is found
    }
  }
  return "#000000";
}

function isInRange(range, targetRange) {
  var sheet = range.getSheet();
  var target = sheet.getRange(targetRange);
  return (
    range.getRow() >= target.getRow() &&
    range.getLastRow() <= target.getLastRow() &&
    range.getColumn() >= target.getColumn() &&
    range.getLastColumn() <= target.getLastColumn()
  );
}

function edit(e) {
  //if (isRunning) { // LATER FOR ERRORS
  //}
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var formula = range.getFormula();
  if ((range.getValue() !== true && range.getValue() !== false) || formula) {
    //alert("worked");
    return;
  }

  if (isInRange(range, 'A3:A')) {
    if ((sheet.getName() === "Edit" || sheet.getName() === "Remove") && !isInRange(range, 'A10:A')) {
      return;
    }
    markComplete(sheet, range.getRow());
  }

  var newItem = getSheet('Add').getRange("C12");
  var newTopic = getSheet('Add').getRange("C19");
  var newKeyword = getSheet('Add').getRange("C26");


  var searchItem1 = getSheet('Remove').getRange("C7");
  var remTopic1 = getSheet('Remove').getRange("F6");
  var remKeyword1 = getSheet('Remove').getRange("J6");

  var searchItem2 = getSheet('Edit').getRange("C7");
  var editTopic1 = getSheet('Edit').getRange("F8");
  var editKeyword1 = getSheet('Edit').getRange("K8");

  switch (sheet.getName()) {
    case 'Add':
      switch (range.getA1Notation()) {
        case newItem.getA1Notation():
          addItem();
          break;
        case newTopic.getA1Notation():
          Utilities.sleep(100);
          addTopic();
          break;
        case newKeyword.getA1Notation():
          addKeyword();
          break;
      }
      break;
    case 'Remove':
      switch (range.getA1Notation()) {
        case searchItem1.getA1Notation():
          searchItems('Remove');
          break;
        case remTopic1.getA1Notation():
          remTopic();
          break;
        case remKeyword1.getA1Notation():

          remKeyword();

          break;
        default:

          if (isInRange(range, 'G11:G')) {
            row = range.getRow();
            remItem(row);
          }
      }
      break;
    case 'Edit':
      switch (range.getA1Notation()) {
        case searchItem2.getA1Notation():
          searchItems('Edit');
          break;
        case editTopic1.getA1Notation():
          editTopic();
          break;
        case editKeyword1.getA1Notation():

          editKeyword();

          break;
        default:

          if (isInRange(range, 'G11:G')) {
            row = range.getRow();
            editItem(row);
          }
      }
      break;
  }

}

function debounce(func, delay = 2000) {
  for (let x = 0; x < 5; x++) {
    if (lock.tryLock(5000)) { // lock to make sure property isn't accessed multiple times at once
      try {
        var calls = getprop(func + "calls");
        if (!calls) {
          calls = 1;
        } else {
          calls = calls + 1;
        }
        setprop(func + "calls", calls);
      }
      finally {
        lock.releaseLock();
        break;
      }
    }
    Utilities.sleep(2000);
  }
  Utilities.sleep(delay);
  callfunc(func, calls);
}

function callfunc(func, call) {
  for (let x = 0; x < 5; x++) {
    if (lock.tryLock(5000)) { // lock to make sure property isn't accessed multiple times at once
      try {
        var calls = getprop(func + "calls");
        //alert(call);
        //alert(calls);
        if (calls > call || calls == 0) {
          lock.releaseLock();
          return;
        }
        if (calls == call) {
          setprop(func + "calls", 0);
          this[func]();
          lock.releaseLock();
          return;
        }
      }
      finally {
        lock.releaseLock();
        break;
      }
    }
    Utilities.sleep(2000);
  }
}

function alert(message, title = "", button = 0) {
  try {
    if (!button) {
      button = ui.ButtonSet.OK;
    }
    if (getprop('silenced')) {
      toast(message);
    } else {


      ui.alert(title, String(message), button);
    }
  }
  catch {
    Logger.log(message);
  }
}

function toast(message) {
  try {
    app.toast(message);
  }
  catch {
    try {
      app.toast(JSON.stringify(message));
    }
    catch {

      Logger.log(message);
    }
  }
}

function getprop(prop) {
  return JSON.parse(user.getProperty(prop));
}

function setprop(name, prop = []) {
  user.setProperty(name, JSON.stringify(prop));
}

function remprop(prop) {
  user.deleteProperty(prop);
}


function copy(sheet, name) {
  var pages = SpreadsheetApp.openById('1Zp4hf1TSKAYE9xOvRplMBQ1Pc8PDvcSk1pPkLFef4UE');
  name = name || sheet;
  try {
    pages.getSheetByName(sheet).copyTo(app).setName(name);
  }
  catch {
    app.deleteSheet(getSheet('Setting Up'));
    pages.getSheetByName(sheet).copyTo(app).setName(name);
  }
  return getSheet(name);

}

function getSheet(sheet) {
  return app.getSheetByName(sheet);
}

function exists(sheet) { // checks if sheet exists, if not create it
  if (!getSheet(sheet)) {
    return copy(sheet);
  }
  return getSheet(sheet); // existing sheet
}

function search(list, item) {
  for (let x of list) {
    if (x == item) {
      return true;
    }
  }
  return false;
}

function endsWithAny(string, values) {
  return values.some(function (value) {
    return string.endsWith(value);
  });
}
