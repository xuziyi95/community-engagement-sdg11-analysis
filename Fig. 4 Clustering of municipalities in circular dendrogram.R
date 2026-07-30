# ============================================================
# Industry-type composition ring by domain-pair cluster
#
# Expected project structure:
# project/
# ├── industry_type_ring.R
# ├── data/
# │   └── Source Data.xlsx
# └── results/
#
# Cluster definition:
# If a column named "Cluster" exists, it is used directly.
# Otherwise, Cluster 1-6 is assigned according to the largest
# value among the six domain-pair ratio columns.
# ============================================================

library(readxl)
library(circlize)

# ============================
# 1. File paths
# ============================
script_args <- commandArgs(trailingOnly = FALSE)
file_arg <- grep("^--file=", script_args, value = TRUE)

if (length(file_arg) > 0) {
  script_path <- normalizePath(
    sub("^--file=", "", file_arg[1]),
    winslash = "/",
    mustWork = FALSE
  )
  project_dir <- dirname(script_path)
} else {
  project_dir <- getwd()
}

file_path <- file.path(
  project_dir,
  "data",
  "Source Data.xlsx"
)

sheet_name <- "Municipal-level data"

output_dir <- file.path(
  project_dir,
  "results"
)

if (!dir.exists(output_dir)) {
  dir.create(
    output_dir,
    recursive = TRUE
  )
}

out_png <- file.path(
  output_dir,
  "Industry_Type_Ring_Grey.png"
)

# ============================
# 2. Read data
# ============================
if (!file.exists(file_path)) {
  stop(
    paste0(
      "Input file not found:\n",
      file_path,
      "\n\nPlace the dataset at 'data/Source Data.xlsx'."
    )
  )
}

df <- read_excel(
  file_path,
  sheet = sheet_name
)

# Clean column names
names(df) <- trimws(
  gsub(
    "\n",
    " ",
    names(df),
    fixed = TRUE
  )
)

# ============================
# 3. Updated column names
# ============================
industry_col <- "Industry_type"

domain_ratio_cols <- c(
  "LIVSRV-LIVSRV_ratio",
  "LIVSRV-MTGADT_ratio",
  "LIVSRV-PLAREG_ratio",
  "MTGADT-MTGADT_ratio",
  "MTGADT-PLAREG_ratio",
  "PLAREG-PLAREG_ratio"
)

required_cols <- c(
  industry_col,
  domain_ratio_cols
)

missing_cols <- setdiff(
  required_cols,
  names(df)
)

if (length(missing_cols) > 0) {
  stop(
    paste0(
      "The following required columns are missing:\n",
      paste(missing_cols, collapse = "\n"),
      "\n\nAvailable columns:\n",
      paste(names(df), collapse = "\n")
    )
  )
}

# ============================
# 4. Numeric conversion
# ============================
df[[industry_col]] <- suppressWarnings(
  as.numeric(df[[industry_col]])
)

for (col in domain_ratio_cols) {
  df[[col]] <- suppressWarnings(
    as.numeric(df[[col]])
  )
}

# Keep valid industry types
df <- df[
  df[[industry_col]] %in% 1:3,
]

# ============================
# 5. Cluster assignment
# ============================
# Use an existing Cluster column if available.
# Otherwise, assign Cluster 1-6 according to the domain-pair
# ratio with the largest value for each municipality.

if ("Cluster" %in% names(df)) {

  clusters <- suppressWarnings(
    as.integer(df[["Cluster"]])
  )

} else {

  ratio_matrix <- as.matrix(
    df[, domain_ratio_cols]
  )

  # Rows with all six values missing cannot be classified
  all_missing <- apply(
    ratio_matrix,
    1,
    function(x) all(is.na(x))
  )

  # Treat partially missing values as unavailable for max selection
  ratio_matrix_for_max <- ratio_matrix
  ratio_matrix_for_max[
    is.na(ratio_matrix_for_max)
  ] <- -Inf

  clusters <- max.col(
    ratio_matrix_for_max,
    ties.method = "first"
  )

  clusters[all_missing] <- NA_integer_
}

df$Cluster <- clusters

# Retain valid clusters and industry types
df <- df[
  !is.na(df$Cluster)
  & df$Cluster %in% 1:6
  & !is.na(df[[industry_col]])
  & df[[industry_col]] %in% 1:3,
]

clusters <- df$Cluster
industry_type <- df[[industry_col]]

if (nrow(df) == 0) {
  stop(
    "No valid municipalities remained after data cleaning."
  )
}

# ============================
# 6. Industry-type share within each cluster
# ============================
cluster_industry_table <- lapply(
  1:6,
  function(cl) {

    values <- industry_type[
      clusters == cl
    ]

    counts <- table(
      factor(
        values,
        levels = 1:3
      )
    )

    total <- sum(counts)

    if (total == 0) {
      return(
        rep(0, 3)
      )
    }

    as.numeric(counts) / total
  }
)

# ============================
# 7. Sector size
# ============================
cluster_sizes <- sapply(
  1:6,
  function(cl) {
    sum(clusters == cl)
  }
)

if (sum(cluster_sizes) == 0) {
  stop(
    "No municipalities were assigned to clusters 1-6."
  )
}

cluster_sizes_norm <- (
  cluster_sizes
  / sum(cluster_sizes)
)

# ============================
# 8. Three-level grayscale palette
# ============================
industry_colors <- c(
  "1" = "#E6E6E6",
  "2" = "#B3B3B3",
  "3" = "#4D4D4D"
)

# ============================
# 9. Sector order
# ============================
custom_order <- c(
  1,
  2,
  4,
  5,
  3,
  6
)

sectors <- paste0(
  "Cluster",
  custom_order
)

cluster_industry_table <- (
  cluster_industry_table[
    custom_order
  ]
)

cluster_sizes_norm <- (
  cluster_sizes_norm[
    custom_order
  ]
)

# Circlize requires positive sector widths.
# Empty clusters receive a very small display width, while
# their industry shares remain zero.
display_sizes <- cluster_sizes_norm

display_sizes[
  display_sizes <= 0
] <- 1e-6

# ============================
# 10. Draw ring
# ============================
png(
  out_png,
  width = 3200,
  height = 3200,
  res = 300
)

par(
  mar = c(
    1,
    1,
    1,
    1
  )
)

circos.clear()

circos.par(
  start.degree = 90,
  gap.degree = 2
)

circos.initialize(
  factors = sectors,
  xlim = cbind(
    rep(
      0,
      length(sectors)
    ),
    display_sizes
  )
)

# ---------- Grayscale ring and outer percentages ----------
circos.trackPlotRegion(
  factors = sectors,
  ylim = c(
    0,
    1.5
  ),
  track.height = 0.22,
  bg.border = NA,
  panel.fun = function(x, y) {

    cluster_name <- CELL_META$sector.index

    cluster_number <- as.numeric(
      gsub(
        "Cluster",
        "",
        cluster_name
      )
    )

    cluster_index <- which(
      custom_order == cluster_number
    )

    proportions <- (
      cluster_industry_table[
        [cluster_index]
      ]
    )

    x0 <- CELL_META$xlim[1]

    for (level in 1:3) {

      proportion_value <- as.numeric(
        proportions[level]
      )

      if (
        is.na(proportion_value)
      ) {
        proportion_value <- 0
      }

      segment_width <- (
        proportion_value
        * CELL_META$xrange
      )

      x1 <- x0 + segment_width
      x_mid <- (x0 + x1) / 2

      # Grayscale segment
      if (segment_width > 0) {
        circos.rect(
          x0,
          0,
          x1,
          1,
          col = industry_colors[
            as.character(level)
          ],
          border = "white"
        )
      }

      # Outer percentage label
      if (proportion_value > 0) {
        circos.text(
          x = x_mid,
          y = 1.55,
          labels = paste0(
            round(
              proportion_value * 100,
              1
            ),
            "%"
          ),
          cex = 0.98,
          facing = "clockwise",
          niceFacing = TRUE,
          adj = c(
            0.5,
            0
          )
        )
      }

      x0 <- x1
    }
  }
)

# ---------- Cluster labels ----------
circos.trackPlotRegion(
  factors = sectors,
  ylim = c(
    0,
    1
  ),
  track.height = 0.08,
  bg.border = NA,
  panel.fun = function(x, y) {

    cluster_name <- CELL_META$sector.index
    cluster_number <- as.numeric(
      gsub(
        "Cluster",
        "",
        cluster_name
      )
    )

    cluster_n <- sum(
      clusters == cluster_number
    )

    circos.text(
      CELL_META$xcenter,
      0.5,
      labels = paste0(
        cluster_name,
        " (n=",
        cluster_n,
        ")"
      ),
      niceFacing = TRUE,
      cex = 1.1
    )
  }
)

title(
  "Industry Type Composition by Cluster",
  cex.main = 1.9
)

dev.off()
circos.clear()

cat(
  "\nIndustry-type ring figure saved to:\n",
  out_png,
  "\n"
)
