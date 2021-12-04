import 'dart:developer';

import 'package:docxtpl/docxtpl.dart';

void main() async {
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

  final DocxTpl docxTpl = DocxTpl(
    docxTemplate: remoteTpl,
    isRemoteFile: true,
  );

  var r = await docxTpl.parseDocxTpl();
  log(r.mergeStatus.toString());
  log(r.message.toString());

  var fields = docxTpl.getMergeFields();

  log(fields.toString());

  await docxTpl.writeMergeFields(data: data);

  var saved = await docxTpl.save('generated_tpl_remote.docx');

  log(saved);
}
