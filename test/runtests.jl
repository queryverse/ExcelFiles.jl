using FileIO
using ExcelFiles
using TableTraits
using TableTraitsUtils
using Base.Test

@testset "ExcelFiles" begin

filename = normpath(Pkg.dir("ExcelReaders"),"test", "TestData.xlsx")

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
    @test df[5] == [9., "III", DataValues.NA, true]
    @test df[6] == [3., DataValues.NA, 3.5, 4]
    @test df[7] == ["FF", DataValues.NA, "GGG", "HHHH"]
    @test df[8] == [DataValues.NA, true, DataValues.NA, false]
    @test df[9] == [Date(2015,3,3), DateTime(2015,2,4,10,14), Date(1988,4,9), Dates.Time(15,2,0)]
    @test df[10] == [Date(1965,4,3), DateTime(1950,8,9,18,40), Dates.Time(19,0,0), DataValues.NA]
    @test eltype(df[11]) == ExcelReaders.ExcelErrorCell
    @test df[12][1] isa ExcelReaders.ExcelErrorCell
    @test df[12][2] isa ExcelReaders.ExcelErrorCell
    @test df[12][3] isa ExcelReaders.ExcelErrorCell
    @test df[12][4] == DataValues.NA
    @test df[13] == [DataValues.NA, 3.4, "HKEJW", DataValues.NA]
end

end
