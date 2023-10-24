library(openxlsx)
library(cluster)

calc_and_save_silhouette <- function(som_map, wb, dendrogram_image_path, colors_neurons) {
    w <- som_map$codes[[1]]
    distances <- dist(w)
    hclus <- hclust(
        d = distances,
        members = (1:(som_map$grid$xdim * som_map$grid$ydim))
    )

    hclus["length"] <- length(hclus$order)

    # Inicializar um vetor para armazenar os valores de silhueta
    best_sil <- list(k = 0, sil = -1)

    # Calcular a silhueta para cada passo do dendrograma
    for (i in 2:(hclus$length-1)) {
        clusters <- cutree(hclus, k = i)
        sil <- silhouette(clusters, distances)
        sil_avg <- mean(sil[, "sil_width"])
        if(sil_avg > best_sil$sil) {
            best_sil$k <- i
            best_sil$sil <- sil_avg
        }
    }

    clusters_k <- cutree(hclus, k = best_sil$k)

    to_export <- as.data.frame(w)
    # to_export$neuronio <- som_map$unit.classif
    # to_export$cluster <- clusters_k[som_map$unit.classif]
    to_export["neuronio"] <- sapply(
        rownames(to_export), function(x) { gsub("V", "", x) }
    )
    to_export["cluster"] <- clusters_k
    # to_export <- to_export[, !colnames(to_export) %in% "(Intercept)"]

    for (i in 1:best_sil$k) {
        group <- data.frame(
            lapply(
                to_export[to_export["cluster"] == i, ],
                function(x) { formatC(x, format = "fg") }
            )
        )
        sheet_name <- paste0(
            "Cluster ", colors_neurons[i], " ", i, " (K=", best_sil$k, ")"
        )
        addWorksheet(wb, sheetName = sheet_name)
        writeDataTable(
            wb,
            sheet = sheet_name,
            x = group
        )
    }

    dif <- hclus$length - best_sil$k
    height_k <- (hclus$height[dif] + hclus$height[dif+1]) / 2
    png(dendrogram_image_path)
    plot(as.dendrogram(hclus), type = "rectangle")
    abline(h = height_k, col = "red")
    legend("topright", legend = paste0("K =", best_sil$k), col = "red", lty = 1)
    dev.off()

    return(clusters_k)
}