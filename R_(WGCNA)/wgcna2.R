# Install and load required packages
if (!require("BiocManager", quietly = TRUE)) {
  install.packages("BiocManager")
}
# BiocManager::install(c("GO.db", "impute", "preprocessCore", "WGCNA"))
# Note: 'package:stats' is not a Bioconductor package; stats is built into R.

library(WGCNA)
library(tidyverse)
library(gridExtra)
library(openxlsx)

# Enable WGCNA multi-threading for faster computation
enableWGCNAThreads()

# Read the data
data <- read.xlsx("C:/Users/omer/TCGADATA/TCGALUAD/rxnExprluad.xlsx", rowNames = FALSE)

# Data preprocessing
# Ensure data is numeric and handle missing values
data <- as.data.frame(data)
if (!all(sapply(data, is.numeric))) {
  stop("All columns in data must be numeric. Please check your input data.")
}
data[is.na(data)] <- 0 # Replace NA with 0 (or use imputation if preferred)
data <- data[, apply(data, 2, var) > 0] # Remove zero-variance columns

# Transpose data (samples as rows, genes as columns)
datExpr <- data

# Power vector for soft threshold
power <- c(1:10, seq(from = 12, to = 20, by = 2))

# Calculate soft threshold
sft <- pickSoftThreshold(datExpr,
                         powerVector = power,
                         networkType = "unsigned",
                         verbose = 5)

sft.data <- sft$fitIndices

# Plot Scale-Free Topology and Mean Connectivity
a1 <- ggplot(sft.data, aes(Power, SFT.R.sq, label = Power)) +
  geom_point() +
  geom_text(nudge_y = 0.1) +
  geom_hline(yintercept = 0.8, color = "red") +
  labs(x = "Power", y = "Scale free topology model fit, signed R^2") +
  theme_classic()

a2 <- ggplot(sft.data, aes(Power, mean.k., label = Power)) +
  geom_point() +
  geom_text(nudge_y = 0.1) +
  labs(x = "Power", y = "Mean Connectivity") +
  theme_classic()

# Save plots
tiff("ST&MC_data.tif", units = "in", width = 7, height = 7, res = 300)
grid.arrange(a1, a2, nrow = 2)
dev.off()

# Choose soft power based on SFT.R.sq >= 0.8
soft_power <- sft.data$Power[which.max(sft.data$SFT.R.sq >= 0.8)]
cat("Selected soft power:", soft_power, "\n")

# Construct module network
bwnet <- blockwiseModules(datExpr,
                          maxBlockSize = 2000,
                          TOMType = "unsigned",
                          power = soft_power,
                          mergeCutHeight = 0.30,
                          numericLabels = TRUE,
                          verbose = 5)

# Calculate module eigengenes
module_eigengenes <- bwnet$MEs

# Check MEs for validity
if (any(is.na(module_eigengenes)) || any(is.infinite(as.matrix(module_eigengenes)))) {
  stop("Module eigengenes contain NA or infinite values.")
}
module_eigengenes <- module_eigengenes[, apply(module_eigengenes, 2, var) > 1e-5] # Remove low-variance eigengenes

# Module labels and colors
moduleLabels <- bwnet$colors
moduleColors <- labels2colors(bwnet$colors)
MEs <- module_eigengenes

# Plot gene dendrogram and module colors
tiff("Gene_dendrodata.tif", units = "in", width = 10, height = 5, res = 300)
plotDendroAndColors(bwnet$dendrograms[[1]], moduleColors[bwnet$blockGenes[[1]]],
                    "Module colors", main = "Gene dendrogram and module colors in block 1",
                    dendroLabels = FALSE, addGuide = TRUE, hang = 0.03, guideHang = 0.05)
dev.off()

##

# Plot eigengene dendrogram (with error handling)
tiff("Eigengene_dendrogram_data.tif", units = "in", width = 10, height = 7, res = 300)
tryCatch({
  plotEigengeneNetworks(MEs, "Eigengene dendrogram", marDendro = c(0, 4, 2, 0),
                        plotHeatmaps = FALSE)
}, error = function(e) {
  message("Error in plotEigengeneNetworks: ", e$message)
  # Fallback: Plot manual dendrogram
  dissTOM <- 1 - cor(MEs, use = "pairwise.complete.obs")
  hclustObj <- hclust(as.dist(dissTOM), method = "average")
  plot(hclustObj, main = "Eigengene dendrogram (manual)")
})
dev.off()

##


# Save modules to Excel
gene_module_key <- tibble::enframe(bwnet$colors, name = "gene", value = "module") %>%
  mutate(module = paste0("ME", module))
gene_module_key$gene <- rownames(data) # Assign gene names
gene_module_key$moduleColors <- labels2colors(gene_module_key$module)

##

# Save each module to a separate Excel file
start_module <- 0
end_module <- max(as.numeric(gsub("ME", "", unique(gene_module_key$module))))
for (i in start_module:end_module) {
  module_name <- paste0("ME", i)
  filtered_data <- gene_module_key %>%
    dplyr::filter(module == module_name)
  if (nrow(filtered_data) > 0) {
    file_name <- paste0("C:/Users/omer/Desktop/RXNLUAD/", i, ".xlsx")
    write.xlsx(filtered_data, file_name, sheetName = "Sheet1", rowNames = FALSE)
  }
}

# Plot eigengene adjacency heatmap
tiff("Eigengene_adjacency_heatmap_data.tif", units = "in", width = 10, height = 10, res = 300)
par(cex = 1.0)
plotEigengeneNetworks(MEs, "Eigengene adjacency heatmap", marHeatmap = c(3, 4, 2, 2),
                      plotDendrograms = FALSE, xLabelsAngle = 90)
dev.off()

# Combined eigengene dendrogram and heatmap
tiff("Eigengene_dendrogram&Eigengene_adjacency_heatmap_data.tif", units = "in", width = 10, height = 12, res = 300)
par(cex = 0.9)
plotEigengeneNetworks(MEs, "", marDendro = c(0, 4, 1, 2), marHeatmap = c(3, 4, 1, 2), cex.lab = 0.8)
dev.off()