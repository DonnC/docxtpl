import 'dart:io';

import 'package:collection/collection.dart' show IterableExtension;
import 'package:flutter/foundation.dart';
import 'package:archive/archive.dart';
import 'package:archive/archive_io.dart';
import 'package:xml/xml.dart';
import 'package:path/path.dart' as path;

import 'docx_constants.dart';
import 'models/tpl_response.dart';
import 'services/index.dart';
import 'tpl_utils.dart';

/// main [DocxTpl] class
class DocxTpl {
  /// .docx file path, can be file path or url to the .docx file
  final  String? docxTemplate;

  /// indicate whether [docxTemplate] is a remote url file path or a local file
  final bool isRemoteFile;

  /// indicate whether [docxTemplate] is from flutter assets folder
  final bool isAssetFile;

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

  var _settings;

  var _settingsInfo;

  DocxTpl({
    this.docxTemplate,
    this.isAssetFile: false,
    this.isRemoteFile: false,
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

  /// get all extracted merge fields from .docx to use to fill data in
  List<String> getMergeFields() {
    return mergedFields.toSet().toList();
  }

  /// write all changes and save, returns String path to file
  Future<String> save(String filepath) async {
    File(filepath)..createSync(recursive: true);

    var zipF = ZipFileEncoder();

    zipF.create(filepath);

    for (var zi in _zip.files) {
      // write back all zi.name == tempDir files to new word doc
      var ziName = zi.name;

      // search for file and return its file object
      //print('===== directory: ${_dir.path} | ziName: $ziName');
      var _file = await searchDir(_dir, ziName);

      if (_file != null) {
        zipF.addFile(_file, ziName);
      }
    }

    // close zip object
    zipF.close();

    // clean up
    await deleteTempDir(_dir);

    return zipF.zipPath;
  }

  /// write all fields with provided data
  Future<void> writeMergeFields({required Map<String, dynamic> data}) async {
    // TODO: Validate data keys must be equal and same with [mergedFields] values

    // get all merge fields extracted and replace with user data
    var fields = getMergeFields();

    // remove any duplicates if any
    var elementTagSet = _instrTextChildren.toSet();

    var elementTags = elementTagSet.toList();

    for (var field in fields) {
      // replace field with proper data in elText
      for (var element in elementTags) {
        // grab the text to check templating {{..}} and change field
        var elText = element.text;

        // only change proper templated fields  {{..}} and leave the rest as is
        if (elText.contains(RegExp(
          '{{\\w*}}',
          caseSensitive: true,
          multiLine: true,
        ))) {
          // replace field with data passed by user
          var rep = elText.replaceAll(RegExp('{{$field}}'), data[field]);
          element.innerText = rep;
        }
      }
    }

    // grab any element's root note and save to disk

    // TODO: First check if file exists in dir and override it
    // grab the root document already changed by calling [element.innerText] = '<new-data>' while replacing fields above
    var documentXmlRootDoc = elementTags.first.root.root.document!;

    // grab xml as is without pretty printed
    var xml = documentXmlRootDoc.toXmlString();

    // write document to temp dir file
    var docXml = path.join(_dir.path, 'word', 'document.xml');

    File _docXmlFile = await File(docXml).create(recursive: true);
    await _docXmlFile.writeAsString(xml);
  }

  /// process word document template passed, if [isRemoteFile] is true, it downloads to temp dir and processes the file
  Future<MergeResponse> parseDocxTpl() async {
    try {
      if (isRemoteFile) {
        // download file first
        var result = await docxRemoteFileDownloader(_dir.path, docxTemplate!);

        if (result is File) {
          // template downloaded successfuly
          _docxFile = result;
        }

        // error downloading remote tpl file
        else {
          throw Exception(
              'error downloading remote .docx template file: ' + result);
        }
      }

      if (isAssetFile) {
        var result = await saveAssetTpl(_dir.path, docxTemplate!);

        if (result is File) {
          // template loaded successfuly
          _docxFile = result;
        }

        // error loading asset tpl file
        else {
          throw Exception('error loading asset .docx template file: ' + result);
        }
      }

      // else take file path passed as is
      // TODO: Validate file | check file extension | check if file exist
      if (!isAssetFile && !isRemoteFile) {
        _docxFile = File(docxTemplate!);
      }

      final bytes = _docxFile.readAsBytesSync();

      _zip = ZipDecoder().decodeBytes(bytes);

      // save all files to temp dir
      for (var zipInnerFile in _zip.files) {
        if (zipInnerFile.isFile) {
          // write file to temp dir
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

        // save file temporarily
        var _tempFile = await File(contentFname).create(recursive: true);
        await _tempFile.writeAsBytes(fileData);

        // begin parsing xml document
        final contentTypes = XmlDocument.parse(_tempFile.readAsStringSync());

        // loop through xml document to check required data
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
              // checking
              var chunkResp = await __getTreeOfFile(file);
              _parts[chunkResp.first] = chunkResp.last;
            }
          }

          // check in another
          if (type == CONTENT_TYPE_SETTINGS) {
            var chunkResp = await __getTreeOfFile(file);
            _settingsInfo = chunkResp.first;
            _settings = chunkResp.last;
          }

          for (var part in _parts.values) {
            // hunt for w:t text and check for simple templating {{<name>}}
            // TODO: Add more checking as word xml structure changes
            for (var parent in part.findAllElements(
              'w:t',
            )) {
              _instrTextChildren.add(parent);
            }

            // use unique fields
            _instrTextChildren.toSet().toList();

            // loop through the _instrTextChildren
            for (var instrChild in _instrTextChildren) {
              // extract merge-field
              var chunkResult = templateParse(instrChild.text);

              /// add merge fields to list
              mergedFields..addAll(chunkResult);
            }
          }
        }
      }

      return MergeResponse(
        mergeStatus: MergeResponseStatus.Success,
        message: 'success',
      );
    }

    // catch any errors
    catch (e) {
      // TODO: Make detailed custom exeptions to return
      return MergeResponse(
        mergeStatus: MergeResponseStatus.Error,
        message: e.toString(),
      );
    }
  }
}
