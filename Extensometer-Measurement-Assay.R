# Load necessary libraries
library(dplyr)
library(ggplot2)
library(openxlsx)

# Original filename
original_file <- 'PATH RAW CSV FILE'

# Load the data, skip the first three lines
data <- read.csv(original_file, skip = 2, header = TRUE, stringsAsFactors = FALSE)

# Extract the data starting from the 2nd row for column 2 and 4
data_extracted <- data.frame("Displacement (nm)" = data[[2]][2:nrow(data)], "Force (uN)" = data[[4]][2:nrow(data)])

# Read the value of row 1 column 1 and save it as DisplacementZero
DisplacementZero <- data_extracted[1, 1]

# Read the value from the first row and second column and save it as ForceZero
ForceZero <- data_extracted[1, 2]

# Subtract DisplacementZero from all values in column 1
data_extracted[, 1] <- data_extracted[, 1] - DisplacementZero

# Subtract ForceZero from all values in column 2
data_extracted[, 2] <- data_extracted[, 2] - ForceZero

# Create empty columns "Displacement (m)", and "Force (N)"
data_extracted$'Displacement (m)' <- data_extracted[, 1] / 1000000000
data_extracted$'Force (N)' <- data_extracted[, 2] / 1000000

# Add values for "Cross sectional area (um2)" and "Length (um)"
# NOTE: YOU HAVE TO MEASURE THE CROSS SECIONAL AREA, FOR HYPOCOTYLS OR STEMS WE CAN USE A=πr^2
# NOTE2: MEASURE THE LENGTH OF THE SAMPLE ON BOTH SIDES, FROM END TO END NOT COVERED BY TOUGH TAGS
cross_sectional_value_m2 <-  58192.38 / 10^12 # Conversion from um2 to m2
length_left_side_m <- 5481.7/ 10^6  # Conversion from um to m
length_right_side_m <- 5481.7/ 10^6  # Conversion from um to m

# Calculate the mean of length values
mean_length_m <- mean(c(length_left_side_m, length_right_side_m))

# Create a data frame for "Cross sectional area (m2)" and "Length (m)"
extra_data <- data.frame(
  `Cross sectional area (m2)` = c(cross_sectional_value_m2, rep(NA, nrow(data_extracted) - 1)),
  `Length (m)` = c(length_left_side_m, length_right_side_m, mean_length_m, rep(NA, nrow(data_extracted) - 3))
)

# Bind this data to your main data frame
data_extracted <- cbind(data_extracted, extra_data)

# Rename the columns if the original column name has dots rather than spaces or brackets
if ("Displacement..nm." %in% colnames(data_extracted)) {
  colnames(data_extracted)[which(colnames(data_extracted) == "Displacement..nm.")] <- "Displacement (nm)"
}
if ("Force..uN." %in% colnames(data_extracted)) {
  colnames(data_extracted)[which(colnames(data_extracted) == "Force..uN.")] <- "Force (uN)"
}
if ("Cross.sectional.area..m2." %in% colnames(data_extracted)) {
  colnames(data_extracted)[which(colnames(data_extracted) == "Cross.sectional.area..m2.")] <- "Cross-sectional area (m2)"
}
if ("Length..m." %in% colnames(data_extracted)) {
  colnames(data_extracted)[which(colnames(data_extracted) == "Length..m.")] <- "Length (m)"
}

# Calculate the Strain (%) and add it to the dataframe
# Initialize the column with zeros
data_extracted$`Strain (%)` <- 0
# Calculate Strain for all rows except the first one
data_extracted$`Strain (%)`[-1] <- (data_extracted$`Displacement (m)`[-1] / mean_length_m) * 100

# Calculate the Stress (MPa) and add it to the dataframe
# Initialize the column with zeros
data_extracted$`Stress (MPa)` <- 0
# Calculate Stress for all rows except the first one
data_extracted$`Stress (MPa)`[-1] <- (data_extracted$`Force (N)`[-1] / cross_sectional_value_m2) / 10^6

# Calculate the rate of change in Stress
data_extracted$`Stress Change` <- c(0, diff(data_extracted$`Stress (MPa)`))

# Find the index where the rate of change surpasses the threshold
start_index <- which(data_extracted$`Stress Change` > 0.01)[1] #Modify this value according to your sample type.

# Set the values in 'Stress (MPa)' column to NA for rows before the threshold
data_extracted$`Stress (MPa)`[1:(start_index-1)] <- NA

# Adjust Strain and Stress to start from their respective values at the threshold
data_extracted$`Strain (%)`[start_index:nrow(data_extracted)] <- data_extracted$`Strain (%)`[start_index:nrow(data_extracted)] - data_extracted$`Strain (%)`[start_index]
data_extracted$`Stress (MPa)`[start_index:nrow(data_extracted)] <- data_extracted$`Stress (MPa)`[start_index:nrow(data_extracted)] - data_extracted$`Stress (MPa)`[start_index]

# Move the adjusted values to the top and fill the bottom with NA
rows_to_move = nrow(data_extracted) - start_index + 1
data_extracted$`Strain (%)`[1:rows_to_move] <- data_extracted$`Strain (%)`[start_index:nrow(data_extracted)]
data_extracted$`Strain (%)`[(rows_to_move + 1):nrow(data_extracted)] <- NA
data_extracted$`Stress (MPa)`[1:rows_to_move] <- data_extracted$`Stress (MPa)`[start_index:nrow(data_extracted)]
data_extracted$`Stress (MPa)`[(rows_to_move + 1):nrow(data_extracted)] <- NA

# Remove the temporary 'Stress Change' column
data_extracted$`Stress Change` <- NULL

# Plotting Stress vs. Strain
plot <- ggplot(data_extracted, aes(x=`Strain (%)`, y=`Stress (MPa)`)) +
  geom_line(color = "black") +
  labs(x = "Strain (%)", y = "Stress (MPa)") +
  theme_bw() +
  theme(axis.title.x = element_text(color="black"),
        axis.title.y = element_text(color="black"),
        axis.text.x = element_text(color="black"),
        axis.text.y = element_text(color="black"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())
print(plot)

# Create a new workbook
wb <- createWorkbook()

# Add the data to the first sheet
addWorksheet(wb, "Data")
writeData(wb, sheet = 1, data_extracted)

# Add the plot to the second sheet
addWorksheet(wb, "Plot")
insertPlot(wb, sheet = 2)

# Define the filename
filename <- sub("\\.csv$", "-plotted.xlsx", original_file)

# Save the workbook
saveWorkbook(wb, filename, overwrite = TRUE)