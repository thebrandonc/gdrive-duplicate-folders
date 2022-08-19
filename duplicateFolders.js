/** ------------------------------------------------------------------------------
/** MAKE SURE THE TAB NAMES BELOW MATCH YOUR SPREADSHEET
/** ------------------------------------------------------------------------------*/

const foldersToCopy = 'Sheet1';  // tab name containing list of folders to duplicate

/** ------------------------------------------------------------------------------
/** BEWARE: EDITING BELOW THIS LINE MAY BREAK THE SCRIPT
/** ------------------------------------------------------------------------------*/


class TemplateDir {
  getTemplate(templateDirId) {
    try {
      this.folder = DriveApp.getFolderById(templateDirId);
      this.parentDir = this.getParentFolder(this.folder);
    } catch(err) {
      const msg = {
          type: '😞 Something went wrong',
          msg: `There was a problem finding the folder to copy. ${err}`
        };
        new Notification().send(msg);
    }
  };

  getParentFolder(folder) {
    try {  
      const parents = folder.getParents();
      while (parents.hasNext()) {
        if (!this.parent) {
          return parents.next();
        };
      };
    } catch(err) {
      const msg = {
          type: '😞 Something went wrong',
          msg: `There was a problem finding the parent folder. ${err}`
        };
        new Notification().send(msg);
    };
  };
};


class DestinationDir {
  create(parentDir, newDirName) {
    this.parentDir = DriveApp.getFolderById(parentDir.getId());
    this.folder = this.parentDir.createFolder(newDirName);
    return this.folder;
  };
};


class CopyManager {
  constructor() {
    this.numFolders = 1;
    this.foldersInspected = 0;
    this.queue = [];
    this.processComplete = false;
  };

  create(templateDir, destinationDir) {
    try {
      this.templateDir = templateDir;
      this.destinationDir = destinationDir;
      while (this.foldersInspected < this.numFolders) {
        const parentFolder = this.foldersInspected === 0 ? true : false;
        const source = parentFolder ? this.templateDir : this.queue[0].source;
        const destination = parentFolder ? this.destinationDir : this.queue[0].destination;
        this.copyFolders(source, destination);
        this.copyFiles(source, destination);
        if (!parentFolder) this.queue.shift();
        this.foldersInspected += 1;
      };
      if (this.numFolders === this.foldersInspected) {
        this.processComplete = true;
      };
    } catch(err) {
        const msg = {
          type: '😞 Something went wrong',
          msg: `There was a problem copying a folder. ${err}`
        };
        new Notification().send(msg);
    };
  };

  copyFolders(source, destination) {
    const folders = source.getFolders();
    while (folders.hasNext()) {
      const subDir = folders.next();
      const newSubDir = destination.createFolder(subDir.getName());
      this.queue.push({ source: subDir, destination: newSubDir });
      this.numFolders += 1;
    };
  };

  copyFiles(source, destination) {
    const files = source.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      file.makeCopy(file.getName(), destination);
    };
  };
};


class Sheet {
  constructor() {
    this.doc = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.doc.getSheetByName(foldersToCopy);
    this.sheetRange = this.sheet.getDataRange();
    this.sheetValues = this.sheetRange.getValues();
    this.foldersToDupe = [];
  };

  getFoldersToDupe() {
    try {
      this.sheetValues.forEach((folder, i) => {
        folder.push(i+1);
        if (folder[0] === '') this.foldersToDupe.push(folder);
      });
      return this.foldersToDupe;
    } catch(err) {
        const msg = {
          type: '😞 Something went wrong',
          msg: `There was a problem finding folders to copy. ${err}`
        };
        new Notification().send(msg);
    };
  };

  markComplete(folder) {
    const cellLocation = `A${folder[folder.length-1]}`;
    const cell = this.sheet.getRange(cellLocation);
    const timestamp = `${new Date()}`;
    cell.setValue(timestamp);
  };
};


class App {
  constructor() {
    this.sheet = new Sheet();
    this.foldersToDupe = this.sheet.getFoldersToDupe();
    this.completedCopies = 0;
  };

  run() {
    try {
      this.foldersToDupe.forEach(folder => {
        const [_, templateDirId, newDirName] = folder;
        const templateDir = new TemplateDir();
        const destinationDir = new DestinationDir()
        const copy = new CopyManager()
        templateDir.getTemplate(templateDirId);
        destinationDir.create(templateDir.parentDir, newDirName);
        copy.create(templateDir.folder, destinationDir.folder);
        if (copy.processComplete) {
          this.sheet.markComplete(folder);
          this.completedCopies += 1;
        }
      });
      let msg;
      if (this.completedCopies > 0) {
        msg = { 
          type: '🏆 Success', 
          msg: `${this.completedCopies} folder(s) successfully copied!` 
        };
      } else {
        msg = { 
          type: 'Copying Complete', 
          msg: `No folders were copied.` 
        };
      };
      new Notification().send(msg);
    } catch(err) {
      const msg = {
        type: '😞 Something went wrong ',
        msg: err
      };
      new Notification().send(msg);
    };
  };
};


class Notification {
  constructor() {
    this.ui = SpreadsheetApp.getUi();
  };

  send(event) {
    this.ui.alert(event.type, event.msg, this.ui.ButtonSet.OK);
  };
};


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂 Duplicator')
      .addItem('📂 Duplicate Folders', 'main')
      .addToUi();
};

function main() {
  new App().run();
};
