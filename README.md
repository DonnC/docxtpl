# docxtpl

A word document template plugin to easily populate and generate word documents from templates

## screenshots
<table>
   <tr>
      <td> Generate From Asset template .docx</td>
      <td> Generate From Remote Template .docx</td>
      <td> Generate From Local Template .docx</td>
   </tr>
   <tr>
      <td><img src="https://raw.githubusercontent.com/DonnC/docxtpl/main/example/demo/docxtpl-asset.gif"></td>
      <td><img src="https://raw.githubusercontent.com/DonnC/docxtpl/main/example/demo/docxtpl-remote.gif"></td>
      <td><img src="https://raw.githubusercontent.com/DonnC/docxtpl/main/example/demo/docxtpl-local.gif"></td>
   </tr>
</table>


## Installation
- add `docxtpl` plugin to your `pubspec.yaml` file
```yaml
dependencies:
  flutter:
    sdk: flutter

  docxtpl: 
```

## Motive
- I tried looking for plugins where i can work with word documents and i wasn't lucky to find what i really needed. From my `python` background i found helpful packages that work around `templating` and i thought maybe i can do something like that in `flutter`
- The idea is to first make an example of the document you want to generate with microsoft word as you want. Then as you are creating your `.docx` word document, you insert `jinja2`-like tags like `{{my-tag-name}}` for example. If you want to put a `name` placeholder to populate later using this plugin, do it like `{{name}}` directly in the document.
- You then save the word document as `.docx` (xml formart) and this is your .docx template file (tpl)
- Now you can use `docxtpl` plugin to generate as many word documents as you want from this tpl file and the fields you will provide

### Template file before and after using `docxtpl` plugin
<table>
   <tr>
      <td> Before: word document template .docx</td>
      <td> After(with docxtpl plugin): template .docx</td>
   </tr>
   <tr>
      <td><img src="https://raw.githubusercontent.com/DonnC/docxtpl/main/example/demo/docxtpl-tpl.png"></td>
      <td><img src="https://raw.githubusercontent.com/DonnC/docxtpl/main/example/demo/docxtpl-filled.png"></td>
   </tr>
</table>

## Usage
First import the `docxtpl` plugin in your dart file
```dart
import 'package:docxtpl/docxtpl.dart';
``` 

Make sure you have created your `.docx` template file and saved it either in your `asset folder` or `remote` or in your device `local storage`.

- `docxtpl` can work with generated templates from `asset folder`, `remote file` and `device storage` file.


### Example: Generate from .docx tpl in asset folder
Make sure you have added your .docx word tpl asset file in `pubspec.yaml` file
```dart
   final DocxTpl docxTpl = DocxTpl(
      docxTemplate: 'assets/invite.docx',  // path where tpl file is
      isAssetFile: true,      // flag to true for tpl file from asset
    );

   // fields corresponding to merge fields found to fill the template with
   var templateData = {
    'name': 'Dart | Flutter Developer',
    'event': 'DocxTpl Plugin Contribution',
    'venue': 'Github',
    'time': 'Anytime',
    'row': '1',
    'guest': 'Flutter',
    'sender': '@DonnC',
  };

   var response = await docxTpl.parseDocxTpl();

   print(response.mergeStatus);
   print(response.message);

    if(response.mergeStatus == MergeResponseStatus.Success) {
      // success, proceed
      // get merge fields extracted by the plugin to know which fields to fill
      var fields = docxTpl.getMergeFields();

      print('Template file fields found: ');
      print(fields);

      await docxTpl.writeMergeFields(data: templateData);

      var savedFile = await docxTpl.save('invitation.docx');
    }
```

## Features & TODO
- [✔]  Simple templating
- [❌] able to pick complex tags
- [❌] add more complex tag formats
- [❌] able to populate a table
- [❌] ability to insert images
- [❌] ability to add custom file formatting (rich-text)
- and more `...`

## Api Changes
Api changes are available on [CHANGELOG](CHANGELOG.md)

### Support
- This plugin offers a very basic word-templating with simple tags
- It was tested with a simple word document
- I really appreciate more support on this, hopefully it can be the ultimate go-to for working with word documents in flutter
- Contributions are welcome with open hands


## references
- A detailed article about .docx word documents [READ HERE](https://www.toptal.com/xml/an-informal-introduction-to-docx)
- `docxtpl` is a inspiration from python libraries that does almost the same i.e word document templating. It is inspired mainly from [python-docx-template](https://github.com/elapouya/python-docx-template) and [docx-mailmerge](https://github.com/Bouke/docx-mailmerge).
- A detailed article on how [docx-mailmerge](https://pbpython.com/python-word-template.html) works (`python`)
- Jinja2 templating [jinja](https://palletsprojects.com/p/jinja/)

