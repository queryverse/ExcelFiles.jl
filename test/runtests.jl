using ExcelFiles
using Tables
using Dates
using XLSX
using DataFrames
using Test

data_directory = joinpath(dirname(pathof(ExcelFiles)), "..", "test","data")
@assert isdir(data_directory)

# Helper: get columns and names from a loaded ExcelFile
function get_cols(source)
    tbl = Tables.columntable(source)
    cols = [collect(tbl[n]) for n in Tables.columnnames(tbl)]
    names = collect(Tables.columnnames(tbl))
    return cols, names
end

@testset "ExcelFiles" verbose=true begin

    filename = joinpath(data_directory, "TestData.xlsx")

    efile = load(filename, "Sheet1")

    @test Tables.istable(efile) == true

    # Test show renders expected number of rows and columns, without depending on exact truncation/wrapping
    @testset "show plain text" begin
        s = sprint(show, efile)
        @test s == "ExcelFile(\"$filename\")"
    end

    @testset "ReadTable" begin
       for source in [load(filename, "Sheet1", "C:O"; first_row=3), load(filename, "Sheet1")]
            df, names = get_cols(source)
            @test length(df) == 13
            @test length(df[1]) == 4

            @test df[1]  == [1., 1.5, 2., 2.5]
            @test df[2]  == ["A", "BB", "CCC", "DDDD"]
            @test df[3]  == [true, false, false, true]
            @test isequal(df[4],  [2, "EEEEE", false, 1.5])
            @test isequal(df[5],  [9., "III", missing, true])
            @test isequal(df[6],  [3., missing, 3.5, 4.])
            @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
            @test isequal(df[8],  [missing, true, missing, false])
            @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), Date(1988, 4, 9), Dates.Time(15, 2, 0)]
            @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
            @test all(ismissing, df[11])
            @test isequal(df[12], [missing, missing, missing, missing])
            @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
        end

        df, names = get_cols(load(filename, "Sheet1", "C:O"; first_row=4, header=false))
        @test names == [:C, :D, :E, :F, :G, :H, :I, :J, :K, :L, :M, :N, :O]
        @test length(df[1]) == 4
        @test length(df) == 13
        @test df[1]  == [1., 1.5, 2., 2.5]
        @test df[2]  == ["A", "BB", "CCC", "DDDD"]
        @test df[3]  == [true, false, false, true]
        @test isequal(df[4],  [2, "EEEEE", false, 1.5])
        @test isequal(df[5],  [9., "III", missing, true])
        @test isequal(df[6],  [3., missing, 3.5, 4.])
        @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
        @test isequal(df[8],  [missing, true, missing, false])
        @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
        @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
        @test all(ismissing, df[11])
        @test all(ismissing, df[12])
        @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
        @test ismissing(df[12][4])

        good_colnames = [:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12, :c13]

        df, names = get_cols(load(filename, "Sheet1", "C:O"; first_row=4, header=false, column_labels=good_colnames))
        @test names == good_colnames
        @test length(df[1]) == 4
        @test length(df) == 13
        @test df[1]  == [1., 1.5, 2., 2.5]
        @test df[2]  == ["A", "BB", "CCC", "DDDD"]
        @test df[3]  == [true, false, false, true]
        @test isequal(df[4],  [2, "EEEEE", false, 1.5])
        @test isequal(df[5],  [9., "III", missing, true])
        @test isequal(df[6],  [3., missing, 3.5, 4.])
        @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
        @test isequal(df[8],  [missing, true, missing, false])
        @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
        @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
        @test all(ismissing, df[11])
        @test all(ismissing, df[12])
        @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
        @test ismissing(df[12][4])

        # Test for saving DataFrame to XLSX
        input = (Day = ["Nov. 27", "Nov. 28", "Nov. 29"], Highest = [78, 79, 75]) |> DataFrame
        save("file.xlsx", input)
        output = load("file.xlsx", "Sheet1") |> DataFrame
        @test input == output
        rm("file.xlsx")

        # Test for saving DataFrame to XLSX with sheetname keyword
        input = (Day = ["Nov. 27", "Nov. 28", "Nov. 29"], Highest = [78, 79, 75]) |> DataFrame
        save("file.xlsx", input, sheetname="SheetName")
        output = load("file.xlsx", "SheetName") |> DataFrame
        @test input == output
        rm("file.xlsx")

        df, names = get_cols(load(filename, "Sheet1"; column_labels=good_colnames))
        @test names == good_colnames
        @test length(df[1]) == 4
        @test length(df) == 13
        @test df[1]  == [1., 1.5, 2., 2.5]
        @test df[2]  == ["A", "BB", "CCC", "DDDD"]
        @test df[3]  == [true, false, false, true]
        @test isequal(df[4],  [2, "EEEEE", false, 1.5])
        @test isequal(df[5],  [9., "III", missing, true])
        @test isequal(df[6],  [3., missing, 3.5, 4.])
        @test isequal(df[7],  ["FF", missing, "GGG", "HHHH"])
        @test isequal(df[8],  [missing, true, missing, false])
        @test df[9]  == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15, 2, 0)]
        @test isequal(df[10], [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19, 0, 0), missing])
        @test all(ismissing, df[11])
        @test all(ismissing, df[12])
        @test isequal(df[13], [missing, 3.4, "HKEJW", missing])
        @test ismissing(df[12][4])

        # Too few column labels
        @test_throws XLSX.XLSXError get_cols(load(filename, "Sheet1", "C:O"; header=true, column_labels=[:c1, :c2, :c3, :c4]))

        # Test for constructing DataFrame with empty header cell
        data, names = get_cols(load(filename, "Sheet2", "C:E"))
        @test names == [:Col1, Symbol("#Empty"), :Col3]

        # normalizenames keyword (XLSX.jl v0.11 only)
        data, names = get_cols(load(filename, "Sheet2", "C:E"; normalizenames=true))
        @test names == [:Col1, :_Empty, :Col3]

    end

    @testset "Transposed tables" begin
        # Note: readtransposedtable cannot handle entirely empty rows/columns,
        # so the Transpose sheet omits those from the original Sheet1 data.
        # Note: eltype of mixed date columns is Dates.TimeType (not Any) when
        # there are no missing values, since a common supertype can be inferred.

        df, names = get_cols(load(filename, "Transpose"; transpose=true, first_column=2))
        @test length(df) == 5
        @test length(df[1]) == 4
        @test names == [Symbol("Some Float64s"), Symbol("Some Strings"), Symbol("Some Bools"), Symbol("Mixed with NA"), Symbol("Some dates")]

        @test df[1] == [1.0, 1.5, 2.0, 2.5]
        @test df[2] == ["A", "BB", "CCC", "DDDD"]
        @test df[3] == Bool[true, false, false, true]
        @test isequal(df[4], Any[9, "III", missing, true])
        @test df[5] == Dates.TimeType[Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), Date(1988, 4, 9), Dates.Time(15, 2, 0)]
    end
end