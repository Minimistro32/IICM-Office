function uploadFileToDrive(fileData, fileNameInfo, parentDir, newFileName) {

    try {

        // Checks for existing directories with the name parentDir
        var folder, folders = DriveApp.getFoldersByName(parentDir);
        if (folders.hasNext()) {
            folder = folders.next();
        } else {
            // If no directory is found, create it
            folder = DriveApp.createFolder(parentDir);
        };

        // Deconstruct the file for reconstruction, conversion, and uploading
        var contentType = fileData.substring(5, fileData.indexOf(";")),
            bytes = Utilities.base64Decode(fileData.substr(fileData.indexOf('base64,') + 7)),
            blob = Utilities.newBlob(bytes, contentType, fileNameInfo);

        // Setup parameters for rebuilding the file
        var convertedFile = {
            'title': newFileName + ' - ' + Date.now(),
            'parents': [{
                'id': folder.getId()
            }]
        };

        // Inserts the File into Drive, converts it, and then returns the file
        convertedFile = Drive.Files.insert(convertedFile, blob, {
            'convert': true
        });

        return convertedFile;

    } catch (err) {
        // Implement better error handling
        throw err;
    };
};
