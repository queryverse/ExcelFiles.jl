module ExcelFiles

using XLSX, FileIO, Tables, Dates

export load, save, File

struct ExcelFile
    filename::String
    sheet::Union{Nothing,String}
    columns::Union{Nothing,String}
    keywords
end

# --- Display ---

# Radically simplified - now relies on universal adoption of Tables.jl among consumers.
# Retain only basic show method.
function Base.show(io::IO, f::ExcelFile)
    print(io, "ExcelFile(\"$(f.filename)\")")
end

# --- FileIO integration ---

function fileio_load(f::FileIO.File{FileIO.format"Excel"}, sheet, columns; kw...)
    return ExcelFile(f.filename, sheet, columns, kw)
end
function fileio_load(f::FileIO.File{FileIO.format"Excel"}, sheet; kw...)
    return ExcelFile(f.filename, sheet, nothing, kw)
end
function fileio_load(f::FileIO.File{FileIO.format"Excel"}; kw...)
    return ExcelFile(f.filename, nothing, nothing, kw)
end

function fileio_save(f::FileIO.File{FileIO.format"Excel"}, data; kw...)
    XLSX.writetable(f.filename, data; kw...)
end

# --- Tables.jl interface ---

Tables.istable(::ExcelFile) = true
Tables.columnaccess(::ExcelFile) = true

function Tables.schema(file::ExcelFile)
    tbl = _readxl(file)
    return Tables.schema(tbl)
end

function Tables.columns(file::ExcelFile)
    return Tables.columns(_readxl(file))
end

function Tables.rows(file::ExcelFile)
    return Tables.rows(_readxl(file))
end

# --- Internal reader ---

function _readxl(file::ExcelFile)
    kw = NamedTuple(file.keywords)

    if get(kw, :transpose, false)
        f = XLSX.readtransposedtable
        kw = NamedTuple{filter(k -> k ∉ (:transpose, :first_row), keys(kw))}(kw)
    else
        f = XLSX.readtable
        kw = NamedTuple{filter(k -> k ∉ (:transpose, :first_column), keys(kw))}(kw)
    end

    if isnothing(file.columns)
        if isnothing(file.sheet)
            table = f(file.filename; kw...)
        else
            table = f(file.filename, file.sheet; kw...)
        end
    else
        table = f(file.filename, file.sheet, file.columns; kw...)
    end

    return table  # XLSX v0.11 returns a Tables.jl-compatible object directly
end

end # module
