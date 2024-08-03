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

extension TplExtensions<E> on List<E> {
  /// Checks if two lists contain the same elements, regardless of order.
  ///
  /// Example:
  /// ```dart
  /// final words = ["Hello", "World"];
  /// final otherWords = ["World", "Hello"];
  /// print(words.same(otherWords)); // Output: true
  /// ```
  ///
  /// Parameters:
  ///   - [items]: The list to compare with.
  ///
  /// Returns a [bool] indicating whether the lists contain the same elements.
  bool same(List<E> items) {
    if (length != items.length) return false;
    bool isTheSame = true;

    for (final E item in items) {
      if (!contains(item)) {
        isTheSame = false;
      }
    }

    return isTheSame;
  }
}
