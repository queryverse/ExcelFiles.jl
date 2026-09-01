# ExcelFiles.jl v1.0.1
* Loading goes through the rewritten ExcelReaders 1.0: legacy xls files are
  read natively via LibXLS.jl and modern xlsx files via XLSX.jl, without any
  Python dependency
* Restore the FileIO integration: the central FileIO registry now routes
  Excel files to XLSX.jl and no longer covers xls files at all, so loading
  ExcelFiles puts its loader first again and registers a format for xls
  files; users who don't load ExcelFiles are unaffected
* Saving to an xls file gives a clear error (writing the legacy format is
  not supported)
* Minimum supported Julia version is 1.10 (required by XLSX 0.12)

# ExcelFiles.jl v1.0.0
* Drop julia 0.7 support
* Migrate to Project.toml
* Fix column type detection

# ExcelFiles.jl v0.9.1
* Update to latest PyCall syntax

# ExcelFiles.jl v0.9.0
* Add support for "application/vnd.dataresource+json" MIME type

# ExcelFiles.jl v0.8.0
* Export FileIO.File and FileIO.@format_str

# ExcelFiles.jl v0.7.0
* Support writing of xlsx files

# ExcelFiles.jl v0.6.1
* Work around bug in pkg registry conversion script

# ExcelFiles.jl v0.6.0
* Drop julia 0.6 support, add julia 0.7 support

# ExcelFiles.jl v0.5.0
* Add show method

# ExcelFiles.jl v0.4.0
* Export load and save

# ExcelFiles.jl v0.3.1
* Fix bug related to skipstartrows etc.

# ExcelFiles.jl v0.3.0
* Incorporate all table functionality from ExcelReaders.jl.
* Drop dependency on DataTables.jl and DataFrames.jl.

# ExcelFiles.jl v0.2.0
* Move to TableTraits.jl

# ExcelFiles.jl v0.1.0
* Bug fix release

# ExcelFiles.jl v0.0.1
* Initial release
