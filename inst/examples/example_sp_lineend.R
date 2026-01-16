library(officer)

sp_lineend()
sp_lineend(type = "triangle")
sp_lineend(type = "arrow", width = "lg", length = "lg")
print(sp_lineend(type = "triangle", width = "lg"))
obj <- sp_lineend(type = "triangle", width = "lg")
update(obj, type = "arrow")
