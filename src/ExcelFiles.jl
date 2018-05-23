module ExcelFiles


using ExcelReaders, IteratorInterfaceExtensions, TableTraits, DataValues,
    TableTraitsUtils, FileIO
import IterableTables

export load, save

struct ExcelFile
    filename::String
    range::String
    keywords
end

function fileio_load(f::FileIO.File{FileIO.format"Excel"}, range; keywords...)
    return ExcelFile(f.filename, range, keywords)
end

IteratorInterfaceExtensions.isiterable(x::ExcelFile) = true
TableTraits.isiterabletable(x::ExcelFile) = true

function gennames(n::Integer)
    res = Vector{Symbol}(n)
    for i in 1:n
        res[i] = Symbol(@sprintf "x%d" i)
    end
    return res
end

function _readxl(file::ExcelReaders.ExcelFile, sheetname::AbstractString, startrow::Int, startcol::Int, endrow::Int, endcol::Int; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
    data = ExcelReaders.readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)

    nrow, ncol = size(data)

    if length(colnames)==0
        if header
            headervec = data[1, :]
            NAcol = map(i->isa(i, DataValues.DataValue) && DataValues.isna(i), headervec)
            headervec[NAcol] = gennames(countnz(NAcol))

            # This somewhat complicated conditional makes sure that column names
            # that are integer numbers end up without an extra ".0" as their name
            colnames = [isa(i, AbstractFloat) ? ( modf(i)[1]==0.0 ? Symbol(Int(i)) : Symbol(string(i)) ) : Symbol(i) for i in vec(headervec)]
        else
            colnames = gennames(ncol)
        end
    elseif length(colnames)!=ncol
        error("Length of colnames must equal number of columns in selected range")
    end

    columns = Array{Any}(ncol)

    for i=1:ncol
        if header
            vals = data[2:end,i]
        else
            vals = data[:,i]
        end

        # Check whether all non-NA values in this column
        # are of the same type
        type_of_el = length(vals)>0 ? typeof(vals[1]) : Any
        for val=vals
            type_of_el = promote_type(type_of_el, typeof(val))
        end

        if type_of_el <: DataValue
            columns[i] = convert(DataValueArray, vals)

            # TODO Check wether this hack is correct
            for (j,v) in enumerate(columns[i])
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
    column_data, col_names = if contains(file.range, "!")
        excelfile = openxl(file.filename)

        sheetname, startrow, startcol, endrow, endcol = ExcelReaders.convert_ref_to_sheet_row_col(file.range)
    
        _readxl(excelfile, sheetname, startrow, startcol, endrow, endcol; file.keywords...)
    else
        excelfile = openxl(file.filename)
        sheet = excelfile.workbook[:sheet_by_name](file.range)
        startrow, startcol, endrow, endcol = ExcelReaders.convert_args_to_row_col(sheet; file.keywords...)

        keywords2 = copy(file.keywords)
        keywords2 = filter(i->!(i[1] in (:skipstartrows, :skipstartcols, :nrows, :ncols)), file.keywords)

        _readxl(excelfile, file.range, startrow, startcol, endrow, endcol; keywords2...)    
    end

    return create_tableiterator(column_data, col_names)
end

function Base.collect(file::ExcelFile)
    return collect(getiterator(file))
end

end # module
