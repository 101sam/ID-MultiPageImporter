# MultiPageImporter for Adobe InDesign

Script for automating the placing (import) of multi-page PDF, AI, and InDesign files inside Adobe InDesign.

> **Note**: This is an updated fork of the original [mike-edel/ID-MultiPageImporter](https://github.com/mike-edel/ID-MultiPageImporter) with fixes for InDesign 2024-2026+ compatibility and several bug fixes.

## Features

- Import multi-page PDFs, AI files, InDesign documents, INDT templates, and IDML files
- Automatic page detection and sizing
- Multiple crop options (Trim, Bleed, Media, Art, etc.)
- Page mapping to existing document pages
- Reverse page order option
- Fit to page with optional bleed
- Custom scaling (X/Y percentages)
- Position alignment options
- Rotation support
- Place on new layer option
- Auto-fills current page number when placing into existing documents
- Remembers preferences between sessions

## Compatibility

- Adobe InDesign CS3 through InDesign 2026+
- Windows and macOS

## Installation

1. Download `MultiPageImporter.jsx`
2. Place it in your InDesign Scripts folder:
   - **macOS**: `~/Library/Preferences/Adobe InDesign/Version XX.0/en_US/Scripts/Scripts Panel/`
   - **Windows**: `C:\Users\[username]\AppData\Roaming\Adobe\InDesign\Version XX.0\en_US\Scripts\Scripts Panel\`
3. Access via Window > Utilities > Scripts panel in InDesign

## Usage

1. Run the script from the Scripts panel
2. Select a PDF, AI, INDD, INDT, or IDML file to import
3. Configure import options in the dialog
4. Click OK to import

## Changelog

### Version 2.7.0 (February 2026)

**Bug Fixes:**
- Fixed "DocumentTitle.pdf doesn't contain a Crop crop type" error with modern PDFs ([#40](https://github.com/mike-edel/ID-MultiPageImporter/issues/40))
- Fixed "JavaScript Error 89867: The default engine 'main' cannot be deleted" ([#40](https://github.com/mike-edel/ID-MultiPageImporter/issues/40))
- Fixed "This value would cause one or more objects to leave the pasteboard" when baseline grid is relative to margins ([#39](https://github.com/mike-edel/ID-MultiPageImporter/issues/39))
- Fixed indefinite loading/hanging with certain PDFs ([#34](https://github.com/mike-edel/ID-MultiPageImporter/issues/34))

**New Features:**
- Added support for .indt InDesign template files ([#38](https://github.com/mike-edel/ID-MultiPageImporter/issues/38))
- Added support for .idml files ([#13](https://github.com/mike-edel/ID-MultiPageImporter/issues/13))
- Auto-fill "Start Placing on Doc Page" with current active page number ([#35](https://github.com/mike-edel/ID-MultiPageImporter/issues/35))

**Technical Improvements:**
- Added modern PDF parsing that handles cross-reference streams (used by InDesign 2024+ exports)
- Improved error messages with actionable solutions for placement failures
- Fixed variable scoping issues throughout the script
- Better cleanup and error handling
- Maintains full backward compatibility with CS3/CS4/CS5+

### Version 2.6.2
- Added basic support for multi-artboard AI files (PDF-compatible)

### Version 2.6.1
- Added new document scale for easy page scaling
- Tag all placed frames with label

### Version 2.6
- Fixed misleading "objects leave the pasteboard" error message

### Version 2.5JJB
- Added support for InDesign CS5 PDFCrop constants

### Version 2.5
- Fallback to import all pages when PDF page count can't be determined
- Removed dependency on Verdana font

### Version 2.2.1
- Added page rotation option

### Version 2.2
- Added page mapping to existing document pages
- Added reverse page order option

### Version 2.1
- CS4 compatibility fix

## Known Issues / Future Enhancements

The following enhancement requests from the original repo are not yet implemented:

- [#27](https://github.com/mike-edel/ID-MultiPageImporter/issues/27) - Place on even/odd pages only
- [#33](https://github.com/mike-edel/ID-MultiPageImporter/issues/33) - Fit content to margins proportionally
- [#25](https://github.com/mike-edel/ID-MultiPageImporter/issues/25) - Auto-crop bleed overlap for spreads
- [#17](https://github.com/mike-edel/ID-MultiPageImporter/issues/17) - Apply object style option
- [#19](https://github.com/mike-edel/ID-MultiPageImporter/issues/19) - Right binding support

## Original Author

Scott Zanelli (lonelytreesw@gmail.com)

## Fork Maintainer

Maintained by [Kids on the Yard](https://kidsontheyard.com) / [KotyMart](https://kotymart.com)

## License

GNU General Public License v2.0 or later
