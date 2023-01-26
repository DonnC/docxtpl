/// parse the simple jinja2-like template from element tag text
/// {{...}}
List<String> templateParse(String text) {
  List<String> fields = [];

  final RegExp re = RegExp(
    '{{\\w*}}',
    caseSensitive: true,
    multiLine: true,
  );

  Iterable<Match> matches = re.allMatches(text);

  if (matches.isEmpty) {
    return fields;
  }

// matches found
  else {
    for (var match in matches) {
      int group = match.groupCount;
      String field = match.group(group)!;

      // remove templating braces
      var firstChunk = field.replaceAll('{{', '').trim();
      var secChunk = firstChunk.replaceAll('}}', '').trim();

      fields.add(secChunk);
    }
  }

  return fields;
}
