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

function dropkey(p::Base.Pairs, key::Symbol)                                                                                                                                 
    nt = NamedTuple(p)                     # convert to NamedTuple                                                                                                           
    NamedTuple{filter(!=(key), keys(nt))}(nt)                                                                                                                                
end

function _readxl(file::ExcelFile)
    kw=NamedTuple(file.keywords)
    if haskey(file.keywords, :transpose)
        if file.keywords[:transpose]==true
            haskey(kw, :first_row) && (kw=NamedTuple{filter(!=(:first_row), keys(kw))}(kw))
            f=XLSX.readtransposedtable
        else
            haskey(kw, :first_column) && (kw=NamedTuple{filter(!=(:first_column), keys(kw))}(kw))
            f=XLSX.readtable
        end
        kw=NamedTuple{filter(!=(:transpose), keys(kw))}(kw)
    else
        haskey(kw, :first_column) && (kw=NamedTuple{filter(!=(:first_column), keys(kw))}(kw))
        f=XLSX.readtable
    end
    if isnothing(file.columns)
        if isnothing(file.sheet)
            table=f(file.filename; kw...)
        else
            table=f(file.filename, file.sheet; kw...)
        end
    else
        table=f(file.filename, file.sheet, file.columns; kw...)
    end
#    else
#        if isnothing(file.columns)
#            if isnothing(file.sheet)
#                table=XLSX.readtable(file.filename; dropkey(file.keywords, :transpose)...)
#            else
#                table=XLSX.readtable(file.filename, file.sheet; dropkey(file.keywords, :transpose)...)
#            end
#        else
#            table=XLSX.readtable(file.filename, file.sheet, file.columns; dropkey(file.keywords, :transpose)...)
#        end
#    end
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
