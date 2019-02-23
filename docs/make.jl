using Documenter, ExcelFiles

makedocs(
	modules = [ExcelFiles],
	sitename = "ExcelFiles.jl",
	analytics="UA-132838790-1",
	pages = [
        "Introduction" => "index.md"
    ]
)

deploydocs(
    repo = "github.com/queryverse/ExcelFiles.jl.git"
)
