/** ------------------------------------------------------------------------------
/** MAKE SURE THE TAB NAMES BELOW MATCH YOUR SPREADSHEET
/** ------------------------------------------------------------------------------*/

const foldersToCopy = 'Sheet1';  // tab name containing list folders to duplicate

/** ------------------------------------------------------------------------------
/** BEWARE: EDITING BELOW THIS LINE MAY BREAK THE SCRIPT
/** ------------------------------------------------------------------------------*/


class TemplateDir {
  find(rootDirId) {
    try {
      this.rootDirId = rootDirId;
      const rootDir = DriveApp.getFolderById(rootDirId);
      const folders = rootDir.getFolders();
      
      while (folders.hasNext()) {
        const nextFolder = folders.next();
        if (nextFolder.getName().includes('TEMPLATE')) {
          this.folder = nextFolder;
          return nextFolder;
        };
      };
    } catch(err) {
        const msg = {
          type: '😞 Something went wrong',
          msg: `There was a problem finding a template to copy. ${err}`
        };
        new Notification().send(msg);
    };
  };
};


class DestinationDir {
  create(rootDirId, newDirName) {
    this.rootDir = DriveApp.getFolderById(rootDirId);
    this.folder = this.rootDir.createFolder(newDirName);
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
        const rootFolder = this.foldersInspected === 0 ? true : false;
        const source = rootFolder ? this.templateDir : this.queue[0].source;
        const destination = rootFolder ? this.destinationDir : this.queue[0].destination;
        this.copyFolders(source, destination);
        this.copyFiles(source, destination);
        if (!rootFolder) this.queue.shift();
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
        const [_, rootDirId, newDirName] = folder;
        const templateDir = new TemplateDir().find(rootDirId);
        const destinationDir = new DestinationDir().create(rootDirId, newDirName);
        const copy = new CopyManager()
        copy.create(templateDir, destinationDir);
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
