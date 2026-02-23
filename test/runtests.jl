using ExcelFiles
using IteratorInterfaceExtensions
using TableTraits
using TableTraitsUtils
using Dates
using XLSX
using DataValues
using DataFrames
using Test

data_directory = joinpath(dirname(pathof(ExcelFiles)), "..", "data")
@assert isdir(data_directory)

@testset "ExcelFiles" begin

    filename = joinpath(data_directory, "TestData.xlsx")

    efile = load(filename, "Sheet1")

    # XLSX.jl v0.10.4
#    @test sprint((stream, data) -> show(stream, "text/html", data), efile) == "<table><thead><tr><th>Some Float64s</th><th>Some Strings</th><th>Some Bools</th><th>Mixed column</th><th>Mixed with NA</th><th>Float64 with NA</th><th>String with NA</th><th>Bool with NA</th><th>Some dates</th><th>Dates with NA</th><th>Some errors</th><th>Errors with NA</th><th>Column with NULL and then mixed</th></tr></thead><tbody><tr><td>1</td><td>&quot;A&quot;</td><td>true</td><td>2</td><td>9</td><td>3</td><td>&quot;FF&quot;</td><td>#NA</td><td>Date&#40;&quot;2015-03-03&quot;&#41;</td><td>Date&#40;&quot;1965-04-03&quot;&#41;</td><td>#NA</td><td>#NA</td><td>#NA</td></tr><tr><td>1.5</td><td>&quot;BB&quot;</td><td>false</td><td>&quot;EEEEE&quot;</td><td>&quot;III&quot;</td><td>#NA</td><td>#NA</td><td>true</td><td>2015-02-04T10:14:00</td><td>1950-08-09T18:40:00</td><td>#NA</td><td>#NA</td><td>3.4</td></tr><tr><td>2</td><td>&quot;CCC&quot;</td><td>false</td><td>false</td><td>#NA</td><td>3.5</td><td>&quot;GGG&quot;</td><td>#NA</td><td>Date&#40;&quot;1988-04-09&quot;&#41;</td><td>19:00:00</td><td>#NA</td><td>#NA</td><td>&quot;HKEJW&quot;</td></tr><tr><td>2.5</td><td>&quot;DDDD&quot;</td><td>true</td><td>1.5</td><td>true</td><td>4</td><td>&quot;HHHH&quot;</td><td>false</td><td>15:02:00</td><td>#NA</td><td>#NA</td><td>#NA</td><td>#NA</td></tr></tbody></table>"

    # XLSX.jl v0.11.0 (default behaviour in `readtable` switches to `infer_eltypes=true` so the type eg. of Bools is inferred correctly)
    @test sprint((stream, data) -> show(stream, "text/html", data), efile) == "<table><thead><tr><th>Some Float64s</th><th>Some Strings</th><th>Some Bools</th><th>Mixed column</th><th>Mixed with NA</th><th>Float64 with NA</th><th>String with NA</th><th>Bool with NA</th><th>Some dates</th><th>Dates with NA</th><th>Some errors</th><th>Errors with NA</th><th>Column with NULL and then mixed</th></tr></thead><tbody><tr><td>1.0</td><td>&quot;A&quot;</td><td>true</td><td>2</td><td>9</td><td>3.0</td><td>&quot;FF&quot;</td><td>#NA</td><td>2015-03-03</td><td>Date&#40;&quot;1965-04-03&quot;&#41;</td><td>#NA</td><td>#NA</td><td>#NA</td></tr><tr><td>1.5</td><td>&quot;BB&quot;</td><td>false</td><td>&quot;EEEEE&quot;</td><td>&quot;III&quot;</td><td>#NA</td><td>#NA</td><td>true</td><td>2015-02-04T10:14:00</td><td>1950-08-09T18:40:00</td><td>#NA</td><td>#NA</td><td>3.4</td></tr><tr><td>2.0</td><td>&quot;CCC&quot;</td><td>false</td><td>false</td><td>#NA</td><td>3.5</td><td>&quot;GGG&quot;</td><td>#NA</td><td>1988-04-09</td><td>19:00:00</td><td>#NA</td><td>#NA</td><td>&quot;HKEJW&quot;</td></tr><tr><td>2.5</td><td>&quot;DDDD&quot;</td><td>true</td><td>1.5</td><td>true</td><td>4.0</td><td>&quot;HHHH&quot;</td><td>false</td><td>15:02:00</td><td>#NA</td><td>#NA</td><td>#NA</td><td>#NA</td></tr></tbody></table>"
    
    # XLSX.jl v0.10.4
#    @test sprint((stream, data) -> show(stream, "application/vnd.dataresource+json", data), efile) == "{\"schema\":{\"fields\":[{\"name\":\"Some Float64s\",\"type\":\"string\"},{\"name\":\"Some Strings\",\"type\":\"string\"},{\"name\":\"Some Bools\",\"type\":\"string\"},{\"name\":\"Mixed column\",\"type\":\"string\"},{\"name\":\"Mixed with NA\",\"type\":\"string\"},{\"name\":\"Float64 with NA\",\"type\":\"string\"},{\"name\":\"String with NA\",\"type\":\"string\"},{\"name\":\"Bool with NA\",\"type\":\"string\"},{\"name\":\"Some dates\",\"type\":\"string\"},{\"name\":\"Dates with NA\",\"type\":\"string\"},{\"name\":\"Some errors\",\"type\":\"string\"},{\"name\":\"Errors with NA\",\"type\":\"string\"},{\"name\":\"Column with NULL and then mixed\",\"type\":\"string\"}]},\"data\":[{\"Some Float64s\":1,\"Some Strings\":\"A\",\"Some Bools\":true,\"Mixed column\":2,\"Mixed with NA\":9,\"Float64 with NA\":3,\"String with NA\":\"FF\",\"Bool with NA\":null,\"Some dates\":\"2015-03-03\",\"Dates with NA\":\"1965-04-03\",\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":null},{\"Some Float64s\":1.5,\"Some Strings\":\"BB\",\"Some Bools\":false,\"Mixed column\":\"EEEEE\",\"Mixed with NA\":\"III\",\"Float64 with NA\":null,\"String with NA\":null,\"Bool with NA\":true,\"Some dates\":\"2015-02-04T10:14:00\",\"Dates with NA\":\"1950-08-09T18:40:00\",\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":3.4},{\"Some Float64s\":2,\"Some Strings\":\"CCC\",\"Some Bools\":false,\"Mixed column\":false,\"Mixed with NA\":null,\"Float64 with NA\":3.5,\"String with NA\":\"GGG\",\"Bool with NA\":null,\"Some dates\":\"1988-04-09\",\"Dates with NA\":\"19:00:00\",\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":\"HKEJW\"},{\"Some Float64s\":2.5,\"Some Strings\":\"DDDD\",\"Some Bools\":true,\"Mixed column\":1.5,\"Mixed with NA\":true,\"Float64 with NA\":4,\"String with NA\":\"HHHH\",\"Bool with NA\":false,\"Some dates\":\"15:02:00\",\"Dates with NA\":null,\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":null}]}"

    # XLSX.jl v0.11.0 (default behaviour in `readtable` switches to `infer_eltypes=true` so the type eg. of Bools is inferred correctly)
    @test sprint((stream, data) -> show(stream, "application/vnd.dataresource+json", data), efile) == "{\"schema\":{\"fields\":[{\"name\":\"Some Float64s\",\"type\":\"number\"},{\"name\":\"Some Strings\",\"type\":\"string\"},{\"name\":\"Some Bools\",\"type\":\"boolean\"},{\"name\":\"Mixed column\",\"type\":\"string\"},{\"name\":\"Mixed with NA\",\"type\":\"string\"},{\"name\":\"Float64 with NA\",\"type\":\"number\"},{\"name\":\"String with NA\",\"type\":\"string\"},{\"name\":\"Bool with NA\",\"type\":\"boolean\"},{\"name\":\"Some dates\",\"type\":\"string\"},{\"name\":\"Dates with NA\",\"type\":\"string\"},{\"name\":\"Some errors\",\"type\":\"string\"},{\"name\":\"Errors with NA\",\"type\":\"string\"},{\"name\":\"Column with NULL and then mixed\",\"type\":\"string\"}]},\"data\":[{\"Some Float64s\":1.0,\"Some Strings\":\"A\",\"Some Bools\":true,\"Mixed column\":2,\"Mixed with NA\":9,\"Float64 with NA\":3.0,\"String with NA\":\"FF\",\"Bool with NA\":null,\"Some dates\":\"2015-03-03\",\"Dates with NA\":\"1965-04-03\",\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":null},{\"Some Float64s\":1.5,\"Some Strings\":\"BB\",\"Some Bools\":false,\"Mixed column\":\"EEEEE\",\"Mixed with NA\":\"III\",\"Float64 with NA\":null,\"String with NA\":null,\"Bool with NA\":true,\"Some dates\":\"2015-02-04T10:14:00\",\"Dates with NA\":\"1950-08-09T18:40:00\",\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":3.4},{\"Some Float64s\":2.0,\"Some Strings\":\"CCC\",\"Some Bools\":false,\"Mixed column\":false,\"Mixed with NA\":null,\"Float64 with NA\":3.5,\"String with NA\":\"GGG\",\"Bool with NA\":null,\"Some dates\":\"1988-04-09\",\"Dates with NA\":\"19:00:00\",\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":\"HKEJW\"},{\"Some Float64s\":2.5,\"Some Strings\":\"DDDD\",\"Some Bools\":true,\"Mixed column\":1.5,\"Mixed with NA\":true,\"Float64 with NA\":4.0,\"String with NA\":\"HHHH\",\"Bool with NA\":false,\"Some dates\":\"15:02:00\",\"Dates with NA\":null,\"Some errors\":null,\"Errors with NA\":null,\"Column with NULL and then mixed\":null}]}"

#    This test is truncated (... with 8 more columns:) so probably isn't robust - although it passes locally.
#    @test sprint(show, efile) == "4x13 Excel file\nSome Float64s │ Some Strings │ Some Bools │ Mixed column │ Mixed with NA\n──────────────┼──────────────┼────────────┼──────────────┼──────────────\n1.0           │ A            │ true       │ 2            │ 9            \n1.5           │ BB           │ false      │ \"EEEEE\"      │ \"III\"        \n2.0           │ CCC          │ false      │ false        │ #NA          \n2.5           │ DDDD         │ true       │ 1.5          │ true         \n... with 8 more columns: Float64 with NA, String with NA, Bool with NA, Some dates, Dates with NA, Some errors, Errors with NA, Column with NULL and then mixed"

    @test TableTraits.isiterabletable(efile) == true
    @test IteratorInterfaceExtensions.isiterable(efile) == true
    @test showable("text/html", efile) == true
    @test showable("application/vnd.dataresource+json", efile) == true

    @test isiterable(efile) == true

    full_dfs = [create_columns_from_iterabletable(load(filename, "Sheet1", "C:O"; first_row=3)), create_columns_from_iterabletable(load(filename, "Sheet1"))]
    for (df, names) in full_dfs
        @test length(df) == 13
        @test length(df[1]) == 4

        @test df[1] == [1., 1.5, 2., 2.5]
        @test df[2] == ["A", "BB", "CCC", "DDDD"]
        @test df[3] == [true, false, false, true]
        @test df[4] == [2, "EEEEE", false, 1.5]
        @test df[5] == [9., "III", NA, true]
        @test df[6] == [3., NA, 3.5, 4]
        @test df[7] == ["FF", NA, "GGG", "HHHH"]
        @test df[8] == [NA, true, NA, false]
        @test df[9] == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), Date(1988, 4, 9), Dates.Time(15, 2, 0)]
        @test df[10] == [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), NA]
        @test df[11] == [DataValue(), DataValue(), DataValue(), DataValue()]
        @test df[12] == [DataValue(), DataValue(), DataValue(), NA]
        @test df[13] == [NA, 3.4, "HKEJW", NA]
    end

    df, names = create_columns_from_iterabletable(load(filename, "Sheet1", "C:O"; first_row=4, header=false))
    @test names == [:C, :D, :E, :F, :G, :H, :I, :J, :K, :L, :M, :N, :O]
    @test length(df[1]) == 4
    @test length(df) == 13
    @test df[1] == [1., 1.5, 2., 2.5]
    @test df[2] == ["A", "BB", "CCC", "DDDD"]
    @test df[3] == [true, false, false, true]
    @test df[4] == [2, "EEEEE", false, 1.5]
    @test df[5] == [9., "III", NA, true]
    @test df[6] == [3, NA, 3.5, 4]
    @test df[7] == ["FF", NA, "GGG", "HHHH"]
    @test df[8] == [NA, true, NA, false]
    @test df[9] == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
    @test df[10] == [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), NA]
    @test df[11] == [DataValue(), DataValue(), DataValue(), DataValue()]
    @test df[12] == [DataValue(), DataValue(), DataValue(), NA]
    @test df[13] == [NA, 3.4, "HKEJW", NA]
    @test DataValues.isna(df[12][4])

    good_colnames = [:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12, :c13]

    df, names = create_columns_from_iterabletable(load(filename, "Sheet1", "C:O"; first_row=4, header=false, column_labels=good_colnames))
    @test names == good_colnames
    @test length(df[1]) == 4
    @test length(df) == 13
    @test df[1] == [1., 1.5, 2., 2.5]
    @test df[2] == ["A", "BB", "CCC", "DDDD"]
    @test df[3] == [true, false, false, true]
    @test df[4] == [2, "EEEEE", false, 1.5]
    @test df[5] == [9., "III", NA, true]
    @test df[6] == [3, NA, 3.5, 4]
    @test df[7] == ["FF", NA, "GGG", "HHHH"]
    @test df[8] == [NA, true, NA, false]
    @test df[9] == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
    @test df[10] == [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), NA]
    @test df[11] == [DataValue(), DataValue(), DataValue(), DataValue()]
    @test df[12] == [DataValue(), DataValue(), DataValue(), NA]
    @test df[13] == [NA, 3.4, "HKEJW", NA]
    @test DataValues.isna(df[12][4])

# Test for saving DataFrame to XLSX
    input = (Day = ["Nov. 27","Nov. 28","Nov. 29"], Highest = [78,79,75]) |> DataFrame
    file = save("file.xlsx", input)
    output = load("file.xlsx", "Sheet1") |> DataFrame
    @test input == output
    rm("file.xlsx")

# Test for saving DataFrame to XLSX with sheetname keyword
    input = (Day = ["Nov. 27","Nov. 28","Nov. 29"], Highest = [78,79,75]) |> DataFrame
    file = save("file.xlsx", input, sheetname="SheetName")
    output = load("file.xlsx", "SheetName") |> DataFrame
    @test input == output
    rm("file.xlsx")

    df, names = create_columns_from_iterabletable(load(filename, "Sheet1"; column_labels=good_colnames))
    @test names == good_colnames
    @test length(df[1]) == 4
    @test length(df) == 13
    @test df[1] == [1., 1.5, 2., 2.5]
    @test df[2] == ["A", "BB", "CCC", "DDDD"]
    @test df[3] == [true, false, false, true]
    @test df[4] == [2, "EEEEE", false, 1.5]
    @test df[5] == [9., "III", NA, true]
    @test df[6] == [3, NA, 3.5, 4]
    @test df[7] == ["FF", NA, "GGG", "HHHH"]
    @test df[8] == [NA, true, NA, false]
    @test df[9] == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
    @test df[10] == [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), NA]
    @test df[11] == [DataValue(), DataValue(), DataValue(), DataValue()]
    @test df[12] == [DataValue(), DataValue(), DataValue(), NA]
    @test df[13] == [NA, 3.4, "HKEJW", NA]
    @test DataValues.isna(df[12][4])

# Too few column labels
    # XLSX.jl v0.10.4
#    @test_throws AssertionError create_columns_from_iterabletable(load(filename, "Sheet1", "C:O"; header=true, column_labels=[:c1, :c2, :c3, :c4]))

    # XLSX.jl v0.11.0
    @test_throws XLSX.XLSXError create_columns_from_iterabletable(load(filename, "Sheet1", "C:O"; header=true, column_labels=[:c1, :c2, :c3, :c4]))

# Test for constructing DataFrame with empty header cell
    data, names = create_columns_from_iterabletable(load(filename, "Sheet2", "C:E"))
    @test names == [:Col1, Symbol("#Empty"), :Col3]

    # XLSX.jl v0.11.0. The `normalizenames` keyword not available in 0.10.4
#    data, names = create_columns_from_iterabletable(load(filename, "Sheet2", "C:E"; normalizenames=true))
#    @test names == [:Col1, :_Empty, :Col3]


end
