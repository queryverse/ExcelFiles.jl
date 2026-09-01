# ExcelFiles

[![Project Status: Active - The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![Build Status](https://github.com/queryverse/ExcelFiles.jl/actions/workflows/juliaci.yml/badge.svg?branch=main)](https://github.com/queryverse/ExcelFiles.jl/actions/workflows/juliaci.yml)
[![codecov.io](http://codecov.io/github/queryverse/ExcelFiles.jl/coverage.svg?branch=master)](http://codecov.io/github/queryverse/ExcelFiles.jl?branch=master)

## Overview

This package provides load support for Excel files under the
[FileIO.jl](https://github.com/JuliaIO/FileIO.jl) package. Both modern xlsx
files and legacy xls files can be loaded (the format is detected from the
content of the file); saving is supported for xlsx files.

Note that the central FileIO registry routes Excel files to
[XLSX.jl](https://github.com/JuliaData/XLSX.jl) and does not cover xls files.
Once ExcelFiles is loaded (``using ExcelFiles``), its loader takes priority,
which restores range-based loading (``"Sheet1!A1:C4"``), the lazy iterable
table interface, and support for xls files.

## Installation

Use ``Pkg.add("ExcelFiles")`` in Julia to install ExcelFiles and its dependencies.

## Usage

### Load an Excel file

To read a Excel file into a ``DataFrame``, use the following julia code:

````julia
using ExcelFiles, DataFrames

df = DataFrame(load("data.xlsx", "Sheet1"))
````

The call to ``load`` returns a ``struct`` that is an [IterableTable.jl](https://github.com/queryverse/IterableTables.jl), so it can be passed to any function that can handle iterable tables, i.e. all the sinks in [IterableTable.jl](https://github.com/queryverse/IterableTables.jl). Here are some examples of materializing an Excel file into data structures that are not a ``DataFrame``:

````julia
using ExcelFiles, DataTables, IndexedTables, TimeSeries, Temporal, Gadfly

# Load into a DataTable
dt = DataTable(load("data.xlsx", "Sheet1"))

# Load into an IndexedTable
it = IndexedTable(load("data.xlsx", "Sheet1"))

# Load into a TimeArray
ta = TimeArray(load("data.xlsx", "Sheet1"))

# Load into a TS
ts = TS(load("data.xlsx", "Sheet1"))

# Plot directly with Gadfly
plot(load("data.xlsx", "Sheet1"), x=:a, y=:b, Geom.line)
````

The ``load`` function also takes a number of parameters:

````julia
function load(f::FileIO.File{FileIO.format"Excel"}, range; keywords...)
````
#### Arguments:

* ``range``: either the name of the sheet in the Excel file to read, or a full Excel range specification (i.e. "Sheetname!A1:B2"). If omitted, the first sheet is loaded.
* ``header=true``: whether the first row holds the column names. With ``header=false``, columns are named ``x1``, ``x2``, ... unless ``colnames`` is given.
* ``colnames``: a ``Vector{Symbol}`` of column names to use (e.g. ``colnames=[:a, :b]``), instead of names from the header row.
* ``transpose=false``: with ``transpose=true`` the sheet or range is transposed before the table is constructed, for data organized in rows rather than columns.
* The remaining ``keywords`` arguments are the same as in [ExcelReaders.jl](https://github.com/queryverse/ExcelReaders.jl) (which is used under the hood to read Excel files). When ``range`` is a sheet name, the keyword arguments for the ``readxlsheet`` function from ExcelReaders.jl apply, if ``range`` is a range specification, the keyword arguments for the ``readxl`` function apply.

To read many sheets or ranges from the same file efficiently, open it once
with ``ExcelReaders.openxl`` and use the ExcelReaders functions directly:

````julia
using ExcelReaders

f = openxl("data.xlsx")
data1 = readxlsheet(f, "Sheet1")
data2 = readxlsheet(f, "Sheet2")
````

For advanced xlsx-only reading options (native Excel table names, row
callbacks, string-based missing values), [XLSX.jl](https://github.com/JuliaData/XLSX.jl)'s
``XLSX.readtable`` can always be used directly.

### Save an Excel file

The following code saves any iterable table or Tables.jl source as an excel file:
````julia
using ExcelFiles

save("output.xlsx", it)
````
This will work as long as ``it`` is any of the types supported as sources in IterableTables.jl, or any Tables.jl table.

An existing file is overwritten by default; pass ``overwrite=false`` to get
an error instead. A ``sheetname`` keyword sets the name of the sheet, and any
further keyword arguments are passed on to ``XLSX.writetable``.

Multiple sheets can be written to the same file by passing ``name => table``
pairs:

````julia
save("output.xlsx", "Data" => df1, "Summary" => df2)
````

### Using the pipe syntax

``load`` also support the pipe syntax. For example, to load an Excel file into a ``DataFrame``, one can use the following code:

````julia
using ExcelFiles, DataFrame

df = load("data.xlsx", "Sheet1") |> DataFrame
````

To save an iterable table, one can use the following form:

````julia
using ExcelFiles, DataFrame

df = # Aquire a DataFrame somehow

df |> save("output.xlsx")
````

The pipe syntax is especially useful when combining it with [Query.jl](https://github.com/queryverse/Query.jl) queries, for example one can easily load an Excel file, pipe it into a query, then pipe it to the ``save`` function to store the results in a new file.
