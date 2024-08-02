#Select Excel data sheet and store in variable "excel_file"
excel_file <- file.choose()

# Load the readxl package
library(readxl)

#Store all scRNAseq data in variable 'data'
data <- read_excel(excel_file)

# Install dplyr package
install.packages("dplyr")

# Load dplyr library
library(dplyr)

#Prints prompt that asks user to enter a list of genes
cat("Enter a list of genes as either ensemble IDs or common names (separated by spaces or new lines). Enter an empty line when done:\n")

#ENTER GENE LIST AFTER RUNNING THIS
gene_list <- scan(what = character(), quiet = TRUE)

#Prompts user to enter desired tissue
tissue <- readline("Enter tissue: ")

#Prompts user to enter desired diagnosis
diagnosis <- readline("Enter diagnosis (case sensitive to format written in data file): ")

#Prompts user to enter desired cell type
cell_type <- readline("Enter cell type (case sensitive to format written in data file): ")

#Creates data frame where gene information will later be added
processed_genes <- data.frame('Gene' = character(1),
                              'Fold_Change' = 0,
                              'p_adjusted' = 0)

#Create empty data frame for use in later for loops
empty_row <- data.frame('Gene' = character(1),
                        'Fold_Change' = 0,
                        'p_adjusted' = 0)

# Filter all data for desired diagnosis, tissue, and cell type
narrowed_data <- filter(data, data$Diagnosis==diagnosis & data$Tissue == tissue & data$Cell_Type == cell_type)

#Reset indices to 1
gene <- 1
i <- 1

#Loop over each gene & the narrowed data file to fill in the 'processed_genes' data frame
for (gene in gene_list) {
  
  for (i in 1:nrow(narrowed_data)) {
    
    if ( gene == narrowed_data$Gene[i] )
        
     {
      
      # Index of first empty row ("") in the processed_genes data frame
      first_empty_index <- which(processed_genes$Gene == "")[1]
      
      # Fill in empty row with desired values
      processed_genes[first_empty_index, 'Gene'] <- narrowed_data$Gene[i]
      processed_genes[first_empty_index, 'Fold_Change'] <- narrowed_data$log2FoldChange[i]
      processed_genes[first_empty_index, 'p_adjusted'] <- narrowed_data$padj[i]
      
      # Add new empty row to continue
      processed_genes <- rbind(processed_genes, empty_row)
      
    }
      
    }
    
}

# Delete the last row of the processed gene list, which should be empty
processed_genes <- processed_genes[-nrow(processed_genes), ]

# Load the openxlsx library
library(openxlsx)

# Specify the file path where you want to save the Excel file
file_path <- "/Users/sgustin/Desktop/Processed Gene List.xlsx"

# Export the data frame to Excel
write.xlsx(processed_genes, file_path, rowNames = FALSE)

# Confirmation message
cat("Data frame exported to Excel successfully!\n")
