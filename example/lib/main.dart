import 'dart:io';

import 'package:docxtpl/docxtpl.dart';
import 'package:filesystem_picker/filesystem_picker.dart';
import 'package:flutter/material.dart';
import 'package:open_filex/open_filex.dart';
// ignore: depend_on_referenced_packages
import 'package:path/path.dart' as path;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';

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
      home: const MyHomePage(title: 'DocxTpl Demo Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key, required this.title}) : super(key: key);
  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
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
    final directory = await getExternalStorageDirectory();

    String? path = await FilesystemPicker.open(
      title: 'Select .docx template',
      context: context,
      rootDirectory: directory!,
      fsType: FilesystemType.file,
      allowedExtensions: ['.docx'],
    );

    // await generateDocumentFromTpl(path);
  }

  Future<void> askPermissions() async {
    await [
      Permission.storage,
    ].request();
  }

  Future openFile() async {
    try {
      await OpenFilex.open(savedFile);
    } catch (e) {
      // error
      print('[ERROR] Failed to open file: $savedFile');
    }
  }

  Future<void> generateDocumentFromTpl(String file) async {
    setState(() {
      loading = true;
    });

    final directory = await getExternalStorageDirectory();
    var filename = path.join(directory!.path, 'generated_tpl_local.docx');

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
                ? const Center(
                    child: Padding(
                    padding: EdgeInsets.all(20),
                    child: CircularProgressIndicator(),
                  ))
                : const SizedBox.shrink(),
            const Text(
              'Generate document from asset .docx template',
            ),
            OutlinedButton(
              style: OutlinedButton.styleFrom(
                foregroundColor: Colors.white,
                backgroundColor: Colors.grey,
              ),
              onPressed: () async => await generateDocumentFromAssetTpl(),
              child: const Text('Generate from asset tpl'),
            ),
            const SizedBox(height: 30),
            const Divider(),
            const Text(
              'Generate document from remote .docx template',
            ),
            OutlinedButton(
              style: OutlinedButton.styleFrom(
                foregroundColor: Colors.white,
                backgroundColor: Colors.grey,
              ),
              onPressed: () async => await generateDocumentFromRemoteTpl(),
              child: const Text('Generate from remote tpl'),
            ),
            const SizedBox(height: 30),
            const Divider(),
            OutlinedButton(
              style: OutlinedButton.styleFrom(
                foregroundColor: Colors.white,
                backgroundColor: Colors.grey,
              ),
              onPressed: () async => await _pickTplFile(),
              child: const Text('Generate from local .docx file'),
            ),
            const SizedBox(height: 30),
            const Divider(),
            const Text('Merge fields found'),
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
                            style: const TextStyle(color: Colors.white),
                          ),
                        ),
                      ),
                    ),
                  )
                  .toList(),
            ),
            const Text('Generated word document saved to:'),
            Text(
              savedFile,
              style: const TextStyle(
                color: Colors.blue,
                fontSize: 17,
                fontWeight: FontWeight.w600,
              ),
            ),
            const SizedBox(height: 15),
            OutlinedButton(
              onPressed: () async => await openFile(),
              style: OutlinedButton.styleFrom(
                foregroundColor: Colors.white,
                backgroundColor: Colors.grey,
              ),
              child: const Text('Open generated file'),
            ),
          ],
        ),
      ),
    );
  }
}
