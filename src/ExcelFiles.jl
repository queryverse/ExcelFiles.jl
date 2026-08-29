module ExcelFiles


using ExcelReaders, XLSX, IteratorInterfaceExtensions, TableTraits, DataValues
using TableTraitsUtils, FileIO, TableShowUtils, Dates, Printf
import IterableTables

export load, save, File, @format_str

function __init__()
    # Since FileIO PR #439 (July 2026) the central FileIO registry routes
    # format"Excel" to XLSX.jl's own FileIO extension, and no longer covers
    # legacy xls files at all. Loading ExcelFiles puts its loader first
    # again, and registers a format for xls files. Users who don't load
    # ExcelFiles are unaffected.
    try
        loaders = get!(Vector{FileIO.ActionSource}, FileIO.sym2loader, :Excel)
        filter!(x -> x !== ExcelFiles, loaders)
        pushfirst!(loaders, ExcelFiles)
        savers = get!(Vector{FileIO.ActionSource}, FileIO.sym2saver, :Excel)
        filter!(x -> x !== ExcelFiles, savers)
        pushfirst!(savers, ExcelFiles)

        if !haskey(FileIO.sym2info, :ExcelLegacy)
            FileIO.add_format(FileIO.format"ExcelLegacy", (), [".xls"],
                [:ExcelFiles => Base.UUID("89b67f3b-d1aa-5f6f-9ca4-282e8d98620d"), FileIO.LOAD])
        end
    catch err
        @warn "ExcelFiles could not register itself with FileIO. Loading Excel files via FileIO.load may not use ExcelFiles." exception = (err, catch_backtrace())
    end
    return nothing
end

struct ExcelFile
    filename::String
    range::String
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

const ExcelFileFormat = Union{FileIO.File{FileIO.format"Excel"},FileIO.File{FileIO.format"ExcelLegacy"}}

function fileio_load(f::ExcelFileFormat, range; keywords...)
    return ExcelFile(f.filename, range, keywords)
end

function fileio_save(f::FileIO.File{FileIO.format"Excel"}, data; sheetname::AbstractString="")
    cols, colnames = TableTraitsUtils.create_columns_from_iterabletable(data, na_representation=:missing)
    return XLSX.writetable(f.filename, cols, colnames; sheetname=sheetname)
end

function fileio_save(f::FileIO.File{FileIO.format"ExcelLegacy"}, data; kwargs...)
    error("Writing legacy xls files is not supported. Save to an xlsx file instead.")
end

IteratorInterfaceExtensions.isiterable(x::ExcelFile) = true
TableTraits.isiterabletable(x::ExcelFile) = true

function gennames(n::Integer)
    res = Vector{Symbol}(undef, n)
    for i in 1:n
        res[i] = Symbol(@sprintf "x%d" i)
    end
    return res
end

function _readxl(file::ExcelReaders.ExcelFile, sheetname::AbstractString, startrow::Integer, startcol::Integer, endrow::Integer, endcol::Integer; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
    data = ExcelReaders.readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)

    nrow, ncol = size(data)

    if length(colnames) == 0
        if header
            headervec = data[1, :]
            NAcol = map(i -> isa(i, DataValues.DataValue) && DataValues.isna(i), headervec)
            headervec[NAcol] = gennames(count(!iszero, NAcol))

            # This somewhat complicated conditional makes sure that column names
            # that are integer numbers end up without an extra ".0" as their name
            colnames = [isa(i, AbstractFloat) ? ( modf(i)[1] == 0.0 ? Symbol(Int(i)) : Symbol(string(i)) ) : Symbol(i) for i in vec(headervec)]
        else
            colnames = gennames(ncol)
        end
    elseif length(colnames) != ncol
        error("Length of colnames must equal number of columns in selected range")
    end

    columns = Array{Any}(undef, ncol)

    for i = 1:ncol
        if header
            vals = data[2:end,i]
        else
            vals = data[:,i]
        end

        # Check whether all non-NA values in this column
        # are of the same type
        type_of_el = length(vals) > 0 ? typeof(vals[1]) : Any
        for val = vals
            type_of_el = promote_type(type_of_el, typeof(val))
        end

        if type_of_el <: DataValue
            columns[i] = convert(DataValueArray{eltype(type_of_el)}, vals)

            # TODO Check wether this hack is correct
            for (j, v) in enumerate(columns[i])
                if v isa DataValue && !DataValues.isna(v) && v[] isa DataValue
                    columns[i][j] = v[]
                end
            end
        else
            columns[i] = convert(Array{type_of_el}, vals)
        end
    end

    return columns, colnames
end

function IteratorInterfaceExtensions.getiterator(file::ExcelFile)
    column_data, col_names = if occursin("!", file.range)
        excelfile = openxl(file.filename)

        sheetname, startrow, startcol, endrow, endcol = ExcelReaders.convert_ref_to_sheet_row_col(file.range)

        _readxl(excelfile, sheetname, startrow, startcol, endrow, endcol; file.keywords...)
    else
        excelfile = openxl(file.filename)
        sheet = ExcelReaders.sheet_handle(excelfile, file.range)

        keywords = filter(i -> !(i[1] in (:header, :colnames)), file.keywords)
        startrow, startcol, endrow, endcol = ExcelReaders.convert_args_to_row_col(sheet; keywords...)

        keywords2 = copy(file.keywords)
        keywords2 = filter(i -> !(i[1] in (:skipstartrows, :skipstartcols, :nrows, :ncols)), file.keywords)

        _readxl(excelfile, file.range, startrow, startcol, endrow, endcol; keywords2...)
    end

    return create_tableiterator(column_data, col_names)
end

function Base.collect(file::ExcelFile)
    return collect(getiterator(file))
end

end # module
