using FileIO
using IterableTables
using DataFrames
using Base.Test

@testset "ExcelFiles" begin

df = load(joinpath(Pkg.dir("ExcelReaders"), "test", "TestData.xlsx"), "Sheet1") |> DataFrame

@test size(df) == (4,13)

efile = load(joinpath(Pkg.dir("ExcelReaders"), "test", "TestData.xlsx"), "Sheet1")

@test isiterable(efile) == true

end
