// Add this function to any .gs file for testing
function simpleDriveTest() {
  try {
    const specificFolderId = "1nzIgO28PtutJgbu4skYB2a-MoGPGkjt_"; // YOUR FOLDER ID
    let folder;

    console.log("simpleDriveTest: Attempting to get folder by ID: " + specificFolderId);
    folder = DriveApp.getFolderById(specificFolderId);
    console.log("simpleDriveTest: Successfully accessed specific folder: " + folder.getName());

    const testBlob = Utilities.newBlob("Hello Specific Folder", MimeType.PLAIN_TEXT, "driveTestSpecific.txt");
    const fileInSpecific = folder.createFile(testBlob);
    console.log("simpleDriveTest: Test file created in specific folder: " + fileInSpecific.getName() + ", ID: " + fileInSpecific.getId());
    // fileInSpecific.setTrashed(true); // Optional cleanup

    console.log("simpleDriveTest: Attempting to get root folder...");
    const rootFolder = DriveApp.getRootFolder();
    console.log("simpleDriveTest: Successfully accessed root folder: " + rootFolder.getName());
    const testBlobRoot = Utilities.newBlob("Hello Root Folder", MimeType.PLAIN_TEXT, "driveTestRoot.txt");
    const fileInRoot = rootFolder.createFile(testBlobRoot);
    console.log("simpleDriveTest: Test file created in root folder: " + fileInRoot.getName() + ", ID: " + fileInRoot.getId());
    // fileInRoot.setTrashed(true); // Optional cleanup

  } catch (e) {
    console.error("simpleDriveTest Error: " + e.toString() + "\nStack: " + e.stack);
  }
}