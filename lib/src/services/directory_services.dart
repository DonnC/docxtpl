import 'dart:io';
import 'package:path/path.dart' as path;

/// loop through temp dir and search for file
/// if file present return its File object
Future<File?> searchDir(Directory dir, String fileSegment) async {
  // file seperator differs
  // adopt to system's seperator
  var seperator = path.separator;

  var _dirPath = dir.path;

  var allFiles = await dir.list(recursive: true).toList();

  for (var file in allFiles) {
    var tempDirFile = file.path.replaceFirst(_dirPath + seperator, '').trim();

    // replace with inner word xml file format seperator, word xml uses `/` file seperator
    var seg = tempDirFile.replaceAll(seperator, '/');

    if (seg == fileSegment) {
      // return its file
      return File(file.path);
    }
  }

  return null;
}

// delete temp after process done
Future<void> deleteTempDir(var dir) async {
  if (dir.existsSync()) {
    await dir.delete(recursive: true);
  }
}

/// grab system temp directory
Directory tempDir() {
  Directory tempDir = Directory.systemTemp.createTempSync('.docxtpl');
  return tempDir;
}
