import 'dart:io';

import 'package:archive/archive_io.dart';
import 'package:collection/collection.dart' show IterableExtension;
import 'package:path/path.dart' as path;
import 'package:xml/xml.dart';

import '../docxtpl.dart';
import 'docx_constants.dart';
import 'services/index.dart';

/// main [DocxTpl] class
class DocxTpl {
  /// .docx file path, can be file path or url to the .docx file
  final String? docxTemplate;

  /// internal zip file object of the read .docx file
  late Archive _zip;

  /// hold docxFile obj
  late File _docxFile;

  /// hold temp directory path
  late Directory _dir;

  /// hold xml document
  Map<dynamic, XmlDocument> _parts = Map<dynamic, XmlDocument>();

  List<XmlElement> _instrTextChildren = [];

  /// hold merge fields extracted from the document
  List<String> mergedFields = [];

  XmlDocument? _settings;

  ArchiveFile? _settingsInfo;

  DocxTpl({
    required this.docxTemplate,
  }) {
    // get temp dir and save it once
    _dir = tempDir();
  }

  // TODO: isAssetFile != isRemoteFile

  // ignore: always_declare_return_types
  Future<List> __getTreeOfFile(XmlElement file) async {
    var type = file.getAttribute('PartName')!;

    var innerFile = type.replaceFirst('/', '');

    var zi = _zip.findFile(innerFile)!;

    final ziFileData = zi.content as List<int>;

    // save file temporarily
    var ziFilename = path.join(_dir.path, zi.name);
    var ziFile = await File(ziFilename).create(recursive: true);

    await ziFile.writeAsBytes(ziFileData);

    var tempZiStringData = await ziFile.readAsString();

    var parsedZi = XmlDocument.parse(tempZiStringData);

    return [zi, parsedZi];
  }

  /// Retrieves a list of unique merge fields extracted from the .docx template.
  ///
  /// This method returns all merge fields found in the document template,
  /// eliminating any duplicates. Merge fields are placeholders in the document
  /// that can be filled with actual data.
  ///
  /// Returns:
  ///   A [List<String>] containing unique merge field names.
  ///
  /// Example:
  ///   If the document contains fields like ${name}, ${age}, ${name}, ${address},
  ///   this method will return ['name', 'age', 'address'].
  ///
  /// Note:
  ///   The returned list is created from a [Set], so the order of fields
  ///   may not match their order of appearance in the document.
  List<String> getMergeFields() {
    return mergedFields.toSet().toList();
  }

  /// Saves the modified document to a specified file path.
  ///
  /// This method creates a new .docx file (which is essentially a zip file)
  /// at the specified [filepath], copying over all the files from the original
  /// document template, including any modifications made during the merge process.
  ///
  /// Parameters:
  ///   [filepath] - The full path where the new document should be saved.
  ///
  /// Returns:
  ///   A [Future<String>] that completes with the path of the saved file.
  ///
  /// Throws:
  ///   May throw I/O related exceptions if file operations fail.
  ///
  /// Note:
  ///   This method performs several file system operations and should be
  ///   called asynchronously.
  Future<String> save(String filepath) async {
    File(filepath)..createSync(recursive: true);

    var zipF = ZipFileEncoder();

    zipF.create(filepath);

    for (var zi in _zip.files) {
      // Write back all zi.name == tempDir files to new word doc
      var ziName = zi.name;

      // Search for file and return its file object
      var _file = await searchDir(_dir, ziName);

      if (_file != null) {
        zipF.addFile(_file, ziName);
      }
    }

    // Close zip object
    zipF.close();

    // Clean up
    await deleteTempDir(_dir);

    return zipF.zipPath;
  }

  /// Writes the provided data to the merge fields in the document.
  ///
  /// This method replaces all merge fields in the document with the corresponding
  /// values provided in the [data] map.
  ///
  /// Parameters:
  ///   [data] - A map where keys are merge field names and values are the data
  ///            to be inserted into those fields.
  ///
  /// Throws:
  ///   May throw exceptions if file operations fail.
  ///
  /// Note:
  ///   This method performs file I/O operations and should be called asynchronously.
  Future<void> writeMergeFields({required Map<String, dynamic> data}) async {
    // Get all merge fields extracted and replace with user data
    var fields = getMergeFields();

    // Check if all fields are in data
    if (!fields.same(data.keys.toList())) {
      throw ArgumentError('Data should contain all merge fields');
    }

    // Remove any duplicates if any
    var elementTagSet = _instrTextChildren.toSet();

    var elementTags = elementTagSet.toList();

    for (var field in fields) {
      // Replace field with proper data in elText
      for (var element in elementTags) {
        // Grab the text to check templating {{..}} and change field
        var elText = element.innerText;

        // Only change proper templated fields  {{..}} and leave the rest as is
        if (elText.contains(RegExp(
          '{{\\w*}}',
          caseSensitive: true,
          multiLine: true,
        ))) {
          // Replace field with data passed by user
          var rep = elText.replaceAll(RegExp('{{$field}}'), data[field]);
          element.innerText = rep;
        }
      }
    }

    // Grab any element's root note and save to disk
    // Grab the root document already changed by calling [element.innerText] = '<new-data>' while replacing fields above
    var documentXmlRootDoc = elementTags.first.root.root.document!;

    // Grab xml as is without pretty printed
    var xml = documentXmlRootDoc.toXmlString();

    // Write document to temp dir file
    var docXml = path.join(_dir.path, 'word', 'document.xml');

    File _docXmlFile = await File(docXml).create(recursive: true);
    await _docXmlFile.writeAsString(xml);
  }

  /// Parses the DOCX template file and prepares it for merge field replacement.
  ///
  /// This method handles different sources of the DOCX file (remote, asset, or local),
  /// extracts its contents, and identifies merge fields within the document.
  ///
  /// Returns:
  ///   A [Future<MergeResponse>] indicating the success or failure of the parsing operation.
  ///
  /// Throws:
  ///   May throw exceptions for various reasons, including file download failures,
  ///   file reading errors, or invalid document structure.
  Future<MergeResponse> parseDocxTpl() async {
    try {
      if (isRemoteFile) {
        // Download file first
        var result = await docxRemoteFileDownloader(_dir.path, docxTemplate!);

        if (result is File) {
          // Template downloaded successfuly
          _docxFile = result;
        }

        // Error downloading remote tpl file
        else {
          throw Exception(
              'error downloading remote .docx template file: ' + result);
        }
      }

      if (isAssetFile) {
        var result = await saveAssetTpl(_dir.path, docxTemplate!);

        if (result is File) {
          // Template loaded successfully
          _docxFile = result;
        }

        // Error loading asset tpl file
        else {
          throw Exception('error loading asset .docx template file: ' + result);
        }
      }

      if (isLocalFile) {
        // TODO: Validate file | check file extension
        if (File(docxTemplate!).existsSync()) {
          _docxFile = File(docxTemplate!);
        } else {
          throw Exception('file does not exist');
        }
      }

      final bytes = _docxFile.readAsBytesSync();

      _zip = ZipDecoder().decodeBytes(bytes);

      // Save all files to temp dir
      for (var zipInnerFile in _zip.files) {
        if (zipInnerFile.isFile) {
          // Write file to temp dir
          final innerFname = zipInnerFile.name;

          var fname = innerFname.replaceAll('/', path.separator);

          var filename = path.join(_dir.path, fname);

          final innerFData = zipInnerFile.content as List<int>;

          var _tempFile = await File(filename).create(recursive: true);
          await _tempFile.writeAsBytes(innerFData);
        }
      }

      // ignore: omit_local_variable_types
      ArchiveFile? zippedFile = _zip.files.firstWhereOrNull(
        (zippedElement) => zippedElement.name == '[Content_Types].xml',
      );

      if (zippedFile == null) {
        throw Exception('failed to read .docx template file passed');
      }

      if (zippedFile.isFile) {
        // ignore: omit_local_variable_types
        final String filename = zippedFile.name;

        final fileData = zippedFile.content as List<int>;

        var contentFname = path.join(_dir.path, filename);

        // Save file temporarily
        var _tempFile = await File(contentFname).create(recursive: true);
        await _tempFile.writeAsBytes(fileData);

        // Begin parsing xml document
        final contentTypes = XmlDocument.parse(_tempFile.readAsStringSync());

        // Loop through xml document to check required data
        for (var file in contentTypes.findAllElements(
          'Override',
          namespace: "${NAMESPACES['ct']}",
        )) {
          var type = file.getAttribute(
            'ContentType',
            namespace: "${NAMESPACES['ct']}",
          );

          for (var contentTypePart in CONTENT_TYPES_PARTS) {
            if (type == contentTypePart) {
              // Checking
              var chunkResp = await __getTreeOfFile(file);
              _parts[chunkResp.first] = chunkResp.last;
            }
          }

          // Check in another
          if (type == CONTENT_TYPE_SETTINGS) {
            var chunkResp = await __getTreeOfFile(file);
            _settingsInfo = chunkResp.first;
            _settings = chunkResp.last;
          }

          for (var part in _parts.values) {
            // Hunt for w:t text and check for simple templating {{<name>}}
            // TODO: Add more checking as word xml structure changes
            for (var parent in part.findAllElements(
              'w:t',
            )) {
              _instrTextChildren.add(parent);
            }

            // Use unique fields
            _instrTextChildren.toSet().toList();

            // Loop through the _instrTextChildren
            for (var instrChild in _instrTextChildren) {
              // Extract merge-field
              var chunkResult = templateParse(instrChild.innerText);
              mergedFields..addAll(chunkResult);
            }
          }
        }
      }

      return MergeResponse(
        mergeStatus: MergeResponseStatus.Success,
        message: 'success',
      );
    } catch (e) {
      // TODO: Make detailed custom exceptions to return
      return MergeResponse(
        mergeStatus: MergeResponseStatus.Error,
        message: e.toString(),
      );
    }
  }

  DocxTemplateSource get templateSource {
    if (docxTemplate == null) {
      throw ArgumentError('docxTemplate cannot be null');
    }

    if (docxTemplate!.startsWith('http://') ||
        docxTemplate!.startsWith('https://')) {
      return DocxTemplateSource.remote;
    } else if (!docxTemplate!.startsWith('/')) {
      return DocxTemplateSource.asset;
    } else {
      return DocxTemplateSource.local;
    }
  }

  bool get isRemoteFile => templateSource == DocxTemplateSource.remote;
  bool get isAssetFile => templateSource == DocxTemplateSource.asset;
  bool get isLocalFile => templateSource == DocxTemplateSource.local;
}
