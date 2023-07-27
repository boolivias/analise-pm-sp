library(openxlsx)
library(cluster)

calc_and_save_silhouette <- function(weights, wb, dendrogram_image_path, colors_neurons) {
    distances <- dist(weights)
    hclus_obj <- hclust(distances)

    hclus_obj["length"] <- length(hclus_obj$order)

    # Inicializar um vetor para armazenar os valores de silhueta
    best_sil <- list(k = 0, sil = -1)

    # Calcular a silhueta para cada passo do dendrograma
    for (i in 2:(hclus_obj$length-1)) {

        clusters <- cutree(hclus_obj, k = i)
        sil <- silhouette(clusters, distances)
        sil_info <- summary(sil)
        if(sil_info$avg.width > best_sil$sil) {
            best_sil$k <- i
            best_sil$sil <- sil_info$avg.width
        }
    }

    clusters_k <- cutree(hclus_obj, k = best_sil$k)
    idxs <- split(seq_along(clusters_k), clusters_k)
    groups <- lapply(idxs, function(idx) cbind(weights[idx, , drop = FALSE], Neuronio = idx))

    for (i in 1:length(groups)) {
        group <- data.frame(
            lapply(
                data.frame(groups[[i]]),
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
            x = group[, !colnames(group) %in% "X.Intercept."]
        )
    }

    dif <- hclus_obj$length - best_sil$k
    height_k <- (hclus_obj$height[dif] + hclus_obj$height[dif+1]) / 2
    png(dendrogram_image_path)
    plot(as.dendrogram(hclus_obj), type = "rectangle")
    abline(h = height_k, col = "red")
    legend("topright", legend = paste0("K =", best_sil$k), col = "red", lty = 1)
    dev.off()

    return(clusters_k)
}