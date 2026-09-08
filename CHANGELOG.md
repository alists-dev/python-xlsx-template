# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **RichText**: Value-object for styled cell runs, mirrors `docxtpl.RichText`.
  Build styled text in Python and drop into template expressions: `{{ RichText("Bold", bold=True).add(" plain") }}`.
  Supports: `bold`, `italic`, `underline`, `strike`, `color`, `size`, `font`, `subscript`, `superscript`.

- **Variable discovery**: `XlsxTemplate.get_undeclared_template_variables()` scans all sheets without rendering to find variable names referenced in template tags (including `{%r %}`, `{%b %}`, `{% xv %}`, `{% img %}`).

- **Media swapping**: `XlsxTemplate.replace_media(source, replacement)` swaps embedded media (e.g. logos) by CRC32 without re-rendering. Call `reset_replacements()` to clear pending swaps between render cycles.

### Changed

- Refactored OOXML constants and tag patterns into separate modules (`_xml.py`, `_tags.py`) for clarity and reusability. No API changes.

## [0.3.0] - 2025-XX-XX

- Initial public release.
