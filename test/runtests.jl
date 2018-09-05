using ExcelFiles
using ExcelReaders
using IteratorInterfaceExtensions
using TableTraits
using TableTraitsUtils
using Dates
using DataValues
using Test

@testset "ExcelFiles" begin

filename = normpath(dirname(pathof(ExcelReaders)), "..", "test", "TestData.xlsx")

efile = load(filename, "Sheet1")

@test isiterable(efile) == true

full_dfs = [create_columns_from_iterabletable(load(filename, "Sheet1!C3:O7")), create_columns_from_iterabletable(load(filename, "Sheet1"))]
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
    @test df[9] == [Date(2015,3,3), DateTime(2015,2,4,10,14), Date(1988,4,9), Dates.Time(15,2,0)]
    @test df[10] == [Date(1965,4,3), DateTime(1950,8,9,18,40), Dates.Time(19,0,0), NA]
    @test eltype(df[11]) == ExcelReaders.ExcelErrorCell
    @test df[12][1][] isa ExcelReaders.ExcelErrorCell
    @test df[12][2][] isa ExcelReaders.ExcelErrorCell
    @test df[12][3][] isa ExcelReaders.ExcelErrorCell
    @test df[12][4] == NA
    @test df[13] == [NA, 3.4, "HKEJW", NA]
end

df, names = create_columns_from_iterabletable(load(filename, "Sheet1!C4:O7", header=false))
@test names == [:x1,:x2,:x3,:x4,:x5,:x6,:x7,:x8,:x9,:x10,:x11,:x12,:x13]
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
@test df[9] == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15,2,0)]
@test df[10] == [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19,0,0), NA]
@test isa(df[11][1], ExcelReaders.ExcelErrorCell)
@test isa(df[11][2], ExcelReaders.ExcelErrorCell)
@test isa(df[11][3], ExcelReaders.ExcelErrorCell)
@test isa(df[11][4], ExcelReaders.ExcelErrorCell)
@test isa(df[12][1][], ExcelReaders.ExcelErrorCell)
@test isa(df[12][2][], ExcelReaders.ExcelErrorCell)
@test isa(df[12][3][], ExcelReaders.ExcelErrorCell)
@test DataValues.isna(df[12][4])
@test df[13] == [NA, 3.4, "HKEJW", NA]

good_colnames = [:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12, :c13]
df, names = create_columns_from_iterabletable(load(filename, "Sheet1!C4:O7", header=false, colnames=good_colnames))
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
@test df[9] == [Date(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9), Dates.Time(15,2,0)]
@test df[10] == [Date(1965, 4, 3), DateTime(1950, 8, 9, 18, 40), Dates.Time(19,0,0), NA]
@test isa(df[11][1], ExcelReaders.ExcelErrorCell)
@test isa(df[11][2], ExcelReaders.ExcelErrorCell)
@test isa(df[11][3], ExcelReaders.ExcelErrorCell)
@test isa(df[11][4], ExcelReaders.ExcelErrorCell)
@test isa(df[12][1][], ExcelReaders.ExcelErrorCell)
@test isa(df[12][2][], ExcelReaders.ExcelErrorCell)
@test isa(df[12][3][], ExcelReaders.ExcelErrorCell)
@test DataValues.isna(df[12][4])
@test df[13] == [NA, 3.4, "HKEJW", NA]

# Too few colnames
@test_throws ErrorException create_columns_from_iterabletable(load(filename, "Sheet1!C4:O7", header=true, colnames=[:c1, :c2, :c3, :c4]))

# Test for constructing DataFrame with empty header cell
data, names = create_columns_from_iterabletable(load(filename, "Sheet2!C5:E7"))
@test names == [:Col1, :x1, :Col3]

end
