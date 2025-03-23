old_cli_width <- options(cli.width = 200)

withr::defer(
  {
    options(old_cli_width)
    if (file.exists("Rplots.pdf")) file.remove("Rplots.pdf")
  },
  teardown_env()
)
