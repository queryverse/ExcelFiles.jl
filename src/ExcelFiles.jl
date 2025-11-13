module ExcelFiles


using XLSX, IteratorInterfaceExtensions, TableTraits, DataValues
using TableTraitsUtils, FileIO, TableShowUtils, Dates, Printf
import IterableTables

export load, save, File, @format_str

struct ExcelFile
    filename::String
    sheet::Union{Nothing,String}
    columns::Union{Nothing,String}
    keywords
end

function Base.show(io::IO, source::ExcelFile)
    TableShowUtils.printtable(io, getiterator(source), "Excel file")
end

function Base.show(io::IO, ::MIME"text/html", source::ExcelFile)
    TableShowUtils.printHTMLtable(io, getiterator(source))
end

Base.Multimedia.showable(::MIME"text/html", source::ExcelFile) = true

function Base.show(io::IO, ::MIME"application/vnd.dataresource+json", source::ExcelFile)
    TableShowUtils.printdataresource(io, getiterator(source))
end

Base.Multimedia.showable(::MIME"application/vnd.dataresource+json", source::ExcelFile) = true

function fileio_load(f::FileIO.File{FileIO.format"Excel", String}, sheet, columns; kw...)
    return ExcelFile(f.filename, sheet, columns, kw)
end
function fileio_load(f::FileIO.File{FileIO.format"Excel", String}, sheet; kw...)
    return ExcelFile(f.filename, sheet, nothing, kw)
end
function fileio_load(f::FileIO.File{FileIO.format"Excel", String}; kw...)
    return ExcelFile(f.filename, nothing, nothing, kw)
end

function fileio_save(f::FileIO.File{FileIO.format"Excel"}, data; kw...)
    cols, colnames = TableTraitsUtils.create_columns_from_iterabletable(data, na_representation=:missing)
    return XLSX.writetable(f.filename, cols, colnames; kw...)
end

IteratorInterfaceExtensions.isiterable(x::ExcelFile) = true
TableTraits.isiterabletable(x::ExcelFile) = true

function _readxl(file::ExcelFile)
    if isnothing(file.columns)
        if isnothing(file.sheet)
            table=XLSX.readtable(file.filename, "Sheet1"; file.keywords...)
        else
            table=XLSX.readtable(file.filename, file.sheet; file.keywords...)
        end
    else
        table=XLSX.readtable(file.filename, file.sheet, file.columns; file.keywords...)
    end
    colnames=Vector{Symbol}(undef, length(table.data))
    for (k, v) in table.column_label_index
        colnames[v] = Symbol(k)
    end
    return table.data, colnames
end

function IteratorInterfaceExtensions.getiterator(file::ExcelFile)
    column_data, col_names = _readxl(file)
    return create_tableiterator(column_data, col_names)
end

function Base.collect(file::ExcelFile)
    return collect(getiterator(file))
end

end # module
