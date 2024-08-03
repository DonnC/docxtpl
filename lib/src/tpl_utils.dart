/// merge response status to give status updates of [DocxTpl] methods
enum MergeResponseStatus {
  None,
  Success,
  Fail,
  Error,
}

/// Enum to determine the source of the template
enum DocxTemplateSource {
  remote,
  asset,
  local,
}
