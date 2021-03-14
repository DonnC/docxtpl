import 'package:flutter_test/flutter_test.dart';

import 'package:docxtpl/docxtpl.dart';

void main() {
  TestWidgetsFlutterBinding.ensureInitialized();

  var data = {
    'name': 'Dart | Flutter Developer',
    'event': 'DocxTpl Plugin Contribution',
    'venue': 'Github',
    'time': 'Anytime',
    'row': '1',
    'guest': 'Flutter',
    'sender': '@DonnC',
  };

  test('generate .docx from asset', () async {
    final DocxTpl docxTpl = DocxTpl(
      docxTemplate: 'assets/invite.docx',
      isAssetFile: true,
    );

    var r = await docxTpl.parseDocxTpl();
    print(r.mergeStatus);
    print(r.message);

    var fields = docxTpl.getMergeFields();

    print(fields);

    await docxTpl.writeMergeFields(data: data);

    var saved = await docxTpl.save('generated_tpl_asset.docx');

    print(saved);
  });
}
