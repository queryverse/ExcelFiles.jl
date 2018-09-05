# ExcelFiles

[![Project Status: Active - The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![Build Status](https://travis-ci.org/queryverse/ExcelFiles.jl.svg?branch=master)](https://travis-ci.org/queryverse/ExcelFiles.jl)
[![Build status](https://ci.appveyor.com/api/projects/status/wfx5avj0s2m0x94w/branch/master?svg=true)](https://ci.appveyor.com/project/queryverse/excelfiles-jl/branch/master)
[![ExcelFiles](http://pkg.julialang.org/badges/ExcelFiles_0.6.svg)](http://pkg.julialang.org/?pkg=ExcelFiles)
[![codecov.io](http://codecov.io/github/queryverse/ExcelFiles.jl/coverage.svg?branch=master)](http://codecov.io/github/queryverse/ExcelFiles.jl?branch=master)

## Overview

This package provides load support for Excel files under the
[FileIO.jl](https://github.com/JuliaIO/FileIO.jl) package.

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

* ``range``: either the name of the sheet in the Excel file to read, or a full Excel range specification (i.e. "Sheetname!A1:B2").
* The ``keywords`` arguments are the same as in [ExcelReaders.jl](https://github.com/queryverse/ExcelReaders.jl) (which is used under the hood to read Excel files). When ``range`` is a sheet name, the keyword arguments for the ``readxlsheet`` function from ExcelReaders.jl apply, if ``range`` is a range specification, the keyword arguments for the ``readxl`` function apply.

### Using the pipe syntax

``load`` also support the pipe syntax. For example, to load an Excel file into a ``DataFrame``, one can use the following code:

````julia
using ExcelFiles, DataFrame

df = load("data.xlsx", "Sheet1") |> DataFrame
````

The pipe syntax is especially useful when combining it with [Query.jl](https://github.com/queryverse/Query.jl) queries, for example one can easily load an Excel file, pipe it into a query, then pipe it to the ``save`` function to store the results in a new file.
