library(kohonen)
library(stringr)
library(fastDummies)
library(dplyr)

map_similar_strings <- function(str) {
    str <- str_to_lower(str)
    str <- str_remove(str, "\\(a\\)")
    # Remove acentos
    str <- iconv(str, from = "UTF-8", to = "ASCII//TRANSLIT")
    # Remove caracteres especiais
    str <- str_replace_all(str, "[^[:alnum:][:space:]]", "")
    
    str <- str_replace_all(str, "[ç]", "c")
    str <- str_replace_all(str, "[áàâãä]", "a")
    str <- str_replace_all(str, "[éèêë]", "e")
    str <- str_replace_all(str, "[íìîï]", "i")
    str <- str_replace_all(str, "[óòôõö]", "o")
    str <- str_replace_all(str, "[úùûü]", "u")
    str <- str_replace_all(str, "\\W", "_")
    
    # Adiciona espaço no final para evitar a união de duas palavras diferentes
    str <- paste0(str, " ")
    return(str)
}

generate_kohonen_map <- function (data, cols_to_map) {
    mapped_data <- data.frame(lapply(data[cols_to_map], map_similar_strings))
    colnames(mapped_data) <- sub("^map_", "", colnames(mapped_data))
    old_columns <- colnames(mapped_data)

    mapped_data <- dummy_cols(mapped_data)
    data.one_hot <- model.matrix(~.+0, data = select(mapped_data, -one_of(old_columns)))
    map <- som(data.one_hot,
                grid = somgrid(4, 4, "hexagonal"),
                keep.data = TRUE,
    )

    return(map)
}