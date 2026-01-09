# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-01-15

### Added
- Initial release
- Modern CustomTkinter-based GUI
- Dark/Light theme toggle
- Batch CSV file processing
- Folder import functionality
- Multiple encoding support (UTF-8, UTF-8-BOM, Latin-1, CP1251, CP1252, ISO-8859-1, ASCII)
- Custom delimiter selection (Comma, Semicolon, Tab, Pipe)
- Skip rows option
- Custom sheet naming
- Excel styling options:
  - Auto-fit column width
  - Freeze header row
  - Auto-filters
  - Styled header with custom colors
  - Zebra stripes (alternating row colors)
  - Cell borders
- Color presets (Blue, Green, Orange, Purple, Dark)
- Data preview (first 100 rows)
- Progress bar with status updates
- Multi-threaded conversion (non-blocking UI)
- Comprehensive error handling

### Security
- No external network connections
- Local file processing only

## [Unreleased]

### Planned
- Drag and drop support
- Color picker dialog
- Configuration presets (save/load)
- Command-line interface
- Multiple sheets per file
- Data type auto-detection
- Formula support
- Multi-language support
