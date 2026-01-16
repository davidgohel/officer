library(officer)

sp_line()
sp_line(color = "red", lwd = 2)
sp_line(lty = "dot", linecmpd = "dbl")
print(sp_line(color = "red", lwd = 2))
obj <- sp_line(color = "red", lwd = 2)
update(obj, linecmpd = "dbl")
