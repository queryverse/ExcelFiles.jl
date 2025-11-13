# Introduction

This package provides support for Excel files under the
[FileIO.jl](https://github.com/JuliaIO/FileIO.jl) package.

It provides functionality to read simple tabular data from 
an Excel (.xlsx) file and to save simple tabular data to an 
Excel file.

For more extensive functionality when reading and writing Excel files,
consider using [XLSX.jl](https://felipenoris.github.io/XLSX.jl/stable/).
Under the hood, `ExcelFiles.jl` uses the `XLSX.jl` functions `readtable` 
and `writetable`. 

# Usage

## Load an Excel file

To read an Excel file into a `DataFrame`, use the following julia code:

```julia
using ExcelFiles, DataFrames

df = DataFrame(load("data.xlsx", "Sheet1"))
```

The call to `load` returns an object that is an [IterableTable.jl](https://github.com/queryverse/IterableTables.jl), so it can be passed to any function that can handle iterable tables, i.e. all the sinks in [IterableTable.jl](https://github.com/queryverse/IterableTables.jl). Here are some examples of materializing an Excel file into data structures that are not a `DataFrame`:

```julia
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
```

The `load` function takes a number of arguments and keywords:

```julia
    FileIO.load(
        source::String,
        [sheet::String,
        [columns::String]];
        [first_row::Int],
        [column_labels::Vector{String}],
        [header::Bool],
        [normalizenames::Bool]
    )
```

### Arguments:

* `source`: The name of the file to be loaded.
* `sheet`: Specifies the sheet name to be loaded. If `sheet` is not given, the first Excel sheet in the file will be used.
* `columns`: Determines which columns to read. For example, "B:D" will select columns B, C and D. If columns is not given, the algorithm will find the first sequence of consecutive non-empty cells. A valid sheet **must** be specified when specifying columns.

### Keywords:

* `first_row`: Indicates the first row of the data table to be read. For example, `first_row=5` will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the sheet.
* `header`: Indicates if the first row is a header. If `header=true` and `column_labels` is not specified, the column labels for the table will be read from the first row of the table. If `header=false` and `column_labels` is not specified, the algorithm will generate column labels. The default value is `header=true`.
* `column_labels`: Specifies column names for the header of the table. If `column_labels` are given and `header=true`, the headers given by `column_labels` will be used, and the first row of the table (containing headers) will be ignored.
* `normalizenames`: Set to `true` to normalize column names to valid Julia identifiers. Default=`false`.

## Save an Excel file

The following code saves any iterable table as an excel file:
```julia
using ExcelFiles

save("output.xlsx", it)
```
This will work as long as it is any of the types supported as sources in IterableTables.jl (such as a `DataFrame`).

The `save` function takes a number of arguments and keywords:

```
    FileIO.save(
        source::String;
        [overwrite::Bool]
    )
```

### Arguments:

* `sheetname`: Specify the sheetname to be used in the created file. Default = `Sheet1`

### Keywords:

* `overwrite`: Set `overwrite=true` to overwite any existing file of the same name. Default = `false`

## Using the pipe syntax

The `load` and `save` functions also support the pipe syntax. For example, to load an Excel file into a `DataFrame`, one can use the following code:

```julia
using ExcelFiles, DataFrame

df = load("data.xlsx", "Sheet1") |> DataFrame
```

To save an iterable table, one can use the following form:

```julia
using ExcelFiles, DataFrame

df = # Aquire a DataFrame somehow

df |> save("output.xlsx")
```

The pipe syntax is especially useful when combining it with [Query.jl](https://github.com/queryverse/Query.jl) queries, for example one can easily load an Excel file, pipe it into a query, then pipe it to the `save` function to store the results in a new file.
