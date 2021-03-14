import 'package:flutter_test/flutter_test.dart';

import 'package:docxtpl/docxtpl.dart';

void main() {
  TestWidgetsFlutterBinding.ensureInitialized();

  final String remoteTpl =
      'https://gitlab.com/DonnChins/dart-mail-merge/-/raw/master/invite.docx';

  var data = {
    'name': 'Dart | Flutter Developer',
    'event': 'DocxTpl Plugin Contribution',
    'venue': 'Github',
    'time': 'Anytime',
    'row': '1',
    'guest': 'Flutter',
    'sender': '@DonnC',
  };

  test('generate .docx from remote template', () async {
    final DocxTpl docxTpl = DocxTpl(
      docxTemplate: remoteTpl,
      isRemoteFile: true,
    );

    var r = await docxTpl.parseDocxTpl();
    print(r.mergeStatus);
    print(r.message);

    var fields = docxTpl.getMergeFields();

    print(fields);

    await docxTpl.writeMergeFields(data: data);

    var saved = await docxTpl.save('generated_tpl_remote.docx');

    print(saved);
  });
}
