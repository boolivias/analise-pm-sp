# Carregar pacotes necessários
library(readxl)
library(cluster)
library(openxlsx)
source("silhouette.r")
source("kohonen.r")

cols_to_use <- c("COORPORAÇÃO", "SITUAÇÃO", "DESC_TIPOLOCAL", "COR_PELE", "PROFISSAO", "REGIAO", "FAIXA_ETARIA", "PERIODO_DIA")
file_path <- "./files/MDIP.xlsx"

abas <- excel_sheets(file_path)

for (aba in abas) {
    data <- read_excel(file_path, sheet = aba, col_names = TRUE)
    map <- generate_kohonen_map(data, cols_to_use)

    # Criando um objeto workbook
    wb <- createWorkbook()

    addWorksheet(wb, sheetName = "Dados")
    writeData(wb, sheet = "Dados", x = cbind(Neuronio = map$unit.classif, data), startCol = 1, startRow = 1)

    weights <- as.matrix(map$codes[[1]])
    weights[] <- vapply(weights, as.numeric, numeric(1))

    file_prefix <- "kohonen_"
    path_output <- paste("./files/output/", aba)
    if (!dir.exists(path_output)) {
        dir.create(path_output)
    }

    pretty_palette <- c("red", 'yellow', 'green', 'orange', 'purple', 'cyan', 'blue', 'brown', 'white', 'gray', 'magenta')
    clusters <- calc_and_save_silhouette(
        weights,
        wb,
        paste0(path_output, "/", file_prefix, "_dendrogram.png"),
        pretty_palette
    )
    colors <- pretty_palette[clusters]

    saveWorkbook(wb, file = paste0(path_output, "/", file_prefix, "info.xlsx"), overwrite = TRUE)

    png(paste0(path_output, "/", file_prefix, "count.png"))
    plot(map, type = "count", main = "Mapa de Kohonen - count")
    dev.off()

    map$clustering <- kohonen::classvec2classmat(map$unit.classif)
    png(paste0(path_output, "/", file_prefix, "mapping.png"))
    plot(map, type = "mapping", main = "Kohonen - codes", bgcol = colors)
    add.cluster.boundaries(map, clusters)
    dev.off()
}
