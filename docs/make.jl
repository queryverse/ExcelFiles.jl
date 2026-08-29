using Documenter, ExcelFiles

makedocs(modules=[ExcelFiles],
	sitename="ExcelFiles.jl",
	format = Documenter.HTML(analytics = "UA-132838790-1"),
	warnonly = [:missing_docs],
	pages=[
        "Introduction" => "index.md"
    ])

deploydocs(repo="github.com/queryverse/ExcelFiles.jl.git")
