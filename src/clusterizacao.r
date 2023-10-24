# Carregar pacotes necessários
# setwd("c:\\Users\\jean.mendoza\\Desktop\\IC\\projeto\\kohonen-map\\src")
setwd(".")
library(readxl)
library(cluster)
library(openxlsx)
source("./utils/silhouette.r")
source("./utils/kohonen.r")

# cols_to_use <- c("COR_PELE", "REGIAO", "FAIXA_ETARIA", "PERIODO_DIA")
cols_to_use <- c("COR_PELE", "FAIXA_ETARIA")
file_path <- "../data/amostra.xlsx"

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
    path_output <- paste("../out/", aba)
    if (!dir.exists(path_output)) {
        dir.create(path_output)
    }

    pretty_palette <- c('red', 'yellow', 'green', 'orange', 'purple', 'cyan', 'blue', 'brown', 'white', 'gray', 'magenta')
    clusters <- calc_and_save_silhouette(
        map,
        wb,
        paste0(path_output, "/", file_prefix, "_dendrogram.png"),
        pretty_palette
    )
    colors <- pretty_palette[clusters]

    saveWorkbook(wb, file = paste0(path_output, "/", file_prefix, "info.xlsx"), overwrite = TRUE)

    map$clustering <- kohonen::classvec2classmat(map$unit.classif)
    png(paste0(path_output, "/", file_prefix, "codes.png"))
    plot(map, type = "codes", main = "Kohonen - codes", bgcol = colors)
    add.cluster.boundaries(map, clusters)
    dev.off()

    png(paste0(path_output, "/", file_prefix, "count.png"))
    plot(map, type = "count", main = "Mapa de Kohonen - count")
    dev.off()

    # png(paste0(path_output, "/", file_prefix, "codes.png"))
    # plot(map, type = "codes", main = "Mapa de Kohonen - codes")
    # dev.off()
}
