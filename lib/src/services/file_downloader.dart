import 'dart:io';

import 'package:http/http.dart' as http;
import 'package:path/path.dart' as path;

/// download remote .docx and save to temp folder
Future docxRemoteFileDownloader(String directory, String docxUrl) async {
  final String _templateFile = 'docxtpl.docx';

  try {
    final docxPath = path.join(directory, _templateFile);

    final http.Client _client = http.Client();

    var req = await _client.get(Uri.parse(docxUrl));

    var bytes = req.bodyBytes;

    File tplFile = File(docxPath);

    await tplFile.writeAsBytes(bytes);

    return tplFile;
  }

  // catch error
  catch (e) {
    return e.toString();
  }
}

Future saveAssetTpl(String directory, List<int> fileBytes) async {
  final String _templateFile = 'docxtpl.docx';

  try {
    final docxPath = path.join(directory, _templateFile);

    // save to temp dir temporarily
    File tplFile = File(docxPath);

    await tplFile.writeAsBytes(fileBytes);

    return tplFile;
  }

  // catch error
  catch (e) {
    return e.toString();
  }
}
