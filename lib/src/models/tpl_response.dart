import '../tpl_utils.dart';

/// response model for [DocxTpl] methods
class MergeResponse {
  final MergeResponseStatus? mergeStatus;
  final String? message;
  final dynamic data;

  MergeResponse({
    this.mergeStatus,
    this.message,
    this.data,
  });
}
