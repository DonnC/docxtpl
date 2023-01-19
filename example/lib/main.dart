import 'dart:io';

import 'package:ext_storage/ext_storage.dart';
import 'package:filesystem_picker/filesystem_picker.dart';
import 'package:flutter/material.dart';
import 'package:open_file/open_file.dart';
import 'package:path/path.dart' as path;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';

import 'package:docxtpl/docxtpl.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'DocxTpl Demo',
      theme: ThemeData(
        primarySwatch: Colors.blue,
        visualDensity: VisualDensity.adaptivePlatformDensity,
      ),
      home: MyHomePage(title: 'DocxTpl Demo Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  MyHomePage({Key key, this.title}) : super(key: key);
  final String title;

  @override
  _MyHomePageState createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  // word documents templates
  final assetTpl = 'assets/invite.docx';
  final String remoteTpl =
      'https://gitlab.com/DonnChins/dart-mail-merge/-/raw/master/invite.docx';

  // keys should correspond to fields obtained from [docxTpl.getMergeFields()]
  var templateData = {
    'name': 'Dart | Flutter Developer',
    'event': 'DocxTpl Plugin Contribution',
    'venue': 'Github',
    'time': 'Anytime',
    'row': '1',
    'guest': 'Flutter',
    'sender': '@DonnC',
  };

  bool loading = false;

  String savedFile = '';

  List<String> mergeFields = [];

  @override
  void initState() {
    askPermissions();
    super.initState();
  }

  Future<void> _pickTplFile() async {
    // android only
    final directory = await ExtStorage.getExternalStorageDirectory();

    String path = await FilesystemPicker.open(
      title: 'Select .docx template',
      context: context,
      rootDirectory: Directory(directory),
      fsType: FilesystemType.file,
      allowedExtensions: ['.docx'],
    );

    await generateDocumentFromTpl(path);
  }

  Future<void> askPermissions() async {
    await [
      Permission.storage,
    ].request();
  }

  Future openFile() async {
    try {
      await OpenFile.open(savedFile);
    } catch (e) {
      // error
      print('[ERROR] Failed to open file: $savedFile');
    }
  }

  Future<void> generateDocumentFromTpl(String file) async {
    setState(() {
      loading = true;
    });

    final directory = await ExtStorage.getExternalStorageDirectory();
    var filename = path.join(directory, 'generated_tpl_local.docx');

    var saveTo = await File(filename).create(recursive: true);

    final DocxTpl docxTpl = DocxTpl(
      docxTemplate: file,
    );

    var r = await docxTpl.parseDocxTpl();
    print(r.mergeStatus);
    print(r.message);

    var fields = docxTpl.getMergeFields();

    print('[INFO] local template file fields found: ');
    print(fields);

    await docxTpl.writeMergeFields(data: templateData);

    var savedLocal = await docxTpl.save(saveTo.path);

    print('[INFO] Generated document [local] saved to: $savedLocal');

    setState(() {
      mergeFields = fields;
      savedFile = savedLocal;
      loading = false;
    });
  }

  Future<void> generateDocumentFromAssetTpl() async {
    setState(() {
      loading = true;
    });

    final directory = await getTemporaryDirectory();
    var filename = path.join(directory.path, 'generated_tpl_asset.docx');

    var saveTo = await File(filename).create(recursive: true);

    final DocxTpl docxTpl = DocxTpl(
      docxTemplate: 'assets/invite.docx',
      isAssetFile: true,
    );

    var r = await docxTpl.parseDocxTpl();
    print(r.mergeStatus == MergeResponseStatus.Success);
    print(r.message);

    var fields = docxTpl.getMergeFields();

    print('[INFO] asset template file fields found: ');
    print(fields);

    await docxTpl.writeMergeFields(data: templateData);

    var savedAsset = await docxTpl.save(saveTo.path);

    print('[INFO] Generated document [asset] saved to: $savedAsset');

    setState(() {
      mergeFields = fields;
      loading = false;
      savedFile = savedAsset;
    });
  }

  Future<void> generateDocumentFromRemoteTpl() async {
    setState(() {
      loading = true;
    });

    final directory = await getTemporaryDirectory();
    var filename = path.join(directory.path, 'generated_tpl_remote.docx');

    var saveTo = await File(filename).create(recursive: true);

    final DocxTpl docxTpl = DocxTpl(
      docxTemplate: remoteTpl,
      isRemoteFile: true,
    );

    var r = await docxTpl.parseDocxTpl();
    print(r.mergeStatus);
    print(r.message);

    var fields = docxTpl.getMergeFields();

    print('[INFO] remote template file fields found: ');
    print(fields);

    await docxTpl.writeMergeFields(data: templateData);

    var savedRemote = await docxTpl.save(saveTo.path);

    print('[INFO] Generated document [remote] saved to: $savedRemote');

    setState(() {
      mergeFields = fields;
      savedFile = savedRemote;
      loading = false;
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.title),
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(10),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            loading
                ? Center(
                    child: Padding(
                    padding: const EdgeInsets.all(20),
                    child: CircularProgressIndicator(),
                  ))
                : SizedBox.shrink(),
            Text(
              'Generate document from asset .docx template',
            ),
            TextButton(
              style: TextButton.styleFrom(
                foregroundColor: Colors.grey,
              ),
              onPressed: () async => await generateDocumentFromAssetTpl(),
              child: Text('Generate from asset tpl'),
            ),
            SizedBox(height: 30),
            Divider(),
            Text(
              'Generate document from remote .docx template',
            ),
            TextButton(
              style: TextButton.styleFrom(
                foregroundColor: Colors.grey,
              ),
              onPressed: () async => await generateDocumentFromRemoteTpl(),
              child: Text('Generate from remote tpl'),
            ),
            SizedBox(height: 30),
            Divider(),
            TextButton(
              style: TextButton.styleFrom(
                foregroundColor: Colors.grey,
              ),
              onPressed: () async => await _pickTplFile(),
              child: Text('Generate from local .docx file'),
            ),
            SizedBox(height: 30),
            Divider(),
            Text('Merge fields found'),
            Wrap(
              children: mergeFields
                  .map(
                    (field) => Padding(
                      padding: const EdgeInsets.all(8.0),
                      child: Container(
                        decoration: BoxDecoration(
                          borderRadius: BorderRadius.circular(13),
                          color: Colors.blue,
                        ),
                        child: Padding(
                          padding: const EdgeInsets.all(5.0),
                          child: Text(
                            field,
                            style: TextStyle(color: Colors.white),
                          ),
                        ),
                      ),
                    ),
                  )
                  .toList(),
            ),
            Text('Generated word document saved to:'),
            Text(
              savedFile,
              style: TextStyle(
                color: Colors.blue,
                fontSize: 17,
                fontWeight: FontWeight.w600,
              ),
            ),
            SizedBox(height: 15),
            TextButton(
              style: TextButton.styleFrom(
                foregroundColor: Colors.grey,
              ),
              onPressed: () async => await openFile(),
              child: Text('Open generated file'),
            ),
          ],
        ),
      ),
    );
  }
}
