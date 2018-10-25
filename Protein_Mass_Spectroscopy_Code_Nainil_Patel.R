#author: Nainil Patel

#install.packages('XLConnect')
#install.packages('ggplot2')
#install.packages('corrplot')
#install.packages("dplyr")
#install.packages('skimr')
#install.packages('gridExtra')
#install.packages("rpart")
#install.packages("caret")
#install.packages("e1071")
#install.packages("randomForest")
library(XLConnect)
library(ggplot2)
library(corrplot)
library(dplyr)
library(skimr)
library(gridExtra)
library(rpart)
library(caret)
library(e1071)
library(randomForest)

# Set working directory to the folder containing data files
setwd("/Users/vikrant/Desktop/Data Mining/data")

# Define function to merge multiple worksheets
merge_sheets <- function(x, y) {
  if (0 == length(x)) {
    return(y)
  } else {
    x.diff <- setdiff(colnames(x), colnames(y))
    y.diff <- setdiff(colnames(y), colnames(x))
    x[, c(as.character(y.diff))] <- NA
    y[, c(as.character(x.diff))] <- NA
    return(rbind(x, y))
  }
}

# Create an empty global dataframe
global_df <- data.frame()

# Iterate over all .xls files present in data folder
fileNames <- list.files(pattern = "\\.xls$")

for (f in fileNames) {
  # Create an empty excel file specific global dataframe
  file_df <- data.frame()
  
  # Load the Excel file as workbook
  wb = loadWorkbook(f)
  
  # Get list of all Excel sheets present in particular workbook
  sheets <- getSheets(wb)
  
  # Iterate over all worksheets from workbook
  for (s in sheets) {
    
    # Read the worksheet with headers as dataframe
    sheet_df <- readWorksheet(wb, sheet = s, header = TRUE)
    
    # Add extra column 'SheetName' as identifier
    sheet_df['SheetName'] = s
    
    # Merge/Append this dataframe to file specific dataframe
    file_df <- merge_sheets(file_df, sheet_df)
    rm(sheet_df)
  }
  
  # Add extra column 'FileName' as identifier
  file_df['FileName'] = f
  
  # Merge/Append this dataframe to file specific dataframe
  global_df <- merge_sheets(global_df, file_df)
  
  # Remove unwanted vectors to free memory
  rm(file_df)
  rm(s)
  rm(sheets)
  rm(wb)
}
rm(f)
rm(fileNames)
rm(merge_sheets)

# Let's ignore the Columns with less than 50% rows
# Create empty vector to store conditional/boolean filters
flag <- vector()
# Get total number of rows in global dataframe
totalRows <- nrow(global_df)
# Iterate over each column of dataframe
for (c in colnames(global_df)) {
  # Check for the Columns with less than 50% rows
  if (50 <= (100 * length(which(!is.na(global_df[[c]]))) / totalRows)) {
    flag <- c(flag, TRUE)
  } else {
    flag <- c(flag, FALSE)
  }
}
rm(c)
rm(totalRows)

# Use conditional flag vector to filter by columns
df1 <- global_df[c(flag)]

# Get Column wise count for dataframe
# sapply(df2, function(x) sum(complete.cases(x)))

# Also ignore the empty rows where majority of values are NA
df1 <- df1[!is.na(df1$Charge),]
df2 <- global_df[!is.na(global_df$Protein.name),]
rm(flag)
rm(global_df)

# Function to generate Boxplots
get_simple_boxplot <- function(df, column, ylab) {
  colName <- deparse(substitute(column))
  return(qplot(data = df, x = colName,
               y = column, geom = 'boxplot',
               xlab = '',
               ylab = ylab))
}

############################################################################################
# Data Frame 1:
############################################################################################
# Load dataframe 1 to memory
attach(df1)

# Get Structure of final dataframe
str(df1)

# Get Summary of final dataframe
summary(df1)

# Get Skim Summary of final dataframe
skim(df1)

# Visualize the histograms for each variable in dataframe
grid.arrange(qplot(Observed.mass..Da.),
             qplot(Observed.RT..min.),
             qplot(Response),
             qplot(Observed.m.z),
             qplot(Charge),
             ncol = 2)

# Visualize the boxplots for each variable in dataframe
grid.arrange(get_simple_boxplot(df1, Observed.mass..Da., 'Observed.mass..Da.'),
             get_simple_boxplot(df1, Observed.RT..min., 'Observed.RT..min.'),
             get_simple_boxplot(df1, Response, 'Response'),
             get_simple_boxplot(df1, Observed.m.z, 'Observed.m.z'),
             get_simple_boxplot(df1, Charge, 'Charge'),
             ncol = 2)

# Explore the attributes of variable Charge
length(Charge)
unique(Charge)
class(Charge)

# Unload dataframe 1 from memory
detach(df1)

# Visualize the corr plot of the data
cor_data = df1[sapply(df1, function(x) !is.character(x))]
corr_matrix <- cor(cor_data)
corrplot(corr_matrix, method = 'number', type = "lower")
rm(corr_matrix)
rm(df1)

# Factorise the intensity column from caterogical to factor class
cor_data$Charge = factor(cor_data$Charge)

# Visualize the histogram of the intensity
ggplot(cor_data, aes(Charge)) + 
  geom_bar() +
  ggtitle("Histogram for Charge of the Proteins") +
  xlab("Charge") +
  ylab("Count") +
  theme(axis.title.x = element_text(color = "DarkGreen", face = "bold.italic", size = 10),
        axis.title.y = element_text(color = "DarkGreen", face = "bold.italic", size = 10),
        axis.text.x = element_text(size = 10),
        axis.text.y = element_text(size = 10),
        plot.title = element_text(color = "Black", face = "bold.italic", hjust = 0.5, size = 15),
        panel.border = element_rect(colour = "black", fill=NA, size=1))

# Split the dataframe into two Training & Validation sets
train <- sample(nrow(cor_data), 0.7 * nrow(cor_data), replace = FALSE)
TrainSet1 <- cor_data[train,]
ValidSet1 <- cor_data[-train,]
rm(train)
rm(cor_data)

############################################################################################
# Decision Tree

#install.packages("rpart")
#install.packages("caret")
#install.packages("e1071")
library(rpart)
library(caret)
library(e1071)

# We will compare model 1 of Random Forest with Decision Tree model
dt_model = train(Charge ~ ., data = TrainSet1, method = "rpart")
dt_model

predTrain = predict(dt_model, data = TrainSet1)
mean(predTrain == TrainSet1$Charge)
table(predTrain, TrainSet1$Charge)

# Running on Validation Set
predValid = predict(dt_model, newdata = ValidSet1)
mean(predValid == ValidSet1$Charge)
table(predValid, ValidSet1$Charge)

rm(dt_model)
rm(predValid)
rm(predTrain)

############################################################################################
# Random Forest

#install.packages("randomForest")
library(randomForest)

# Create a Random Forest model with fine tuned parameters
rf_model = randomForest(formula = Charge ~ ., data = TrainSet1, importance = TRUE)
rf_model

# Predicting on train set
predTrain <- predict(rf_model, TrainSet1)
mean(predTrain == TrainSet1$Charge)
table(predTrain, TrainSet1$Charge)

# Predicting on Validation set
predValid <- predict(rf_model, ValidSet1, type = "class")
# Checking classification accuracy
mean(predValid == ValidSet1$Charge)
table(predValid, ValidSet1$Charge)

importance(rf_model)
varImpPlot(rf_model)

rm(rf_model)
rm(predValid)
rm(predTrain)

rm(TrainSet1)
rm(ValidSet1)

############################################################################################
# Data Frame 2:
############################################################################################
# Load dataframe 2 to memory
attach(df2)

# Get Structure of final dataframe
str(df2)

# Get Summary of final dataframe
summary(df2)

# Get Skim Summary of final dataframe
skim(df2)

# Visualize the histograms for each variable in dataframe
grid.arrange(qplot(Sequence.start),
             qplot(Sequence.end),
             qplot(Observed.mass..Da.),
             qplot(Observed.m.z),
             qplot(Mass.error..ppm.),
             qplot(Observed.RT..min.),
             qplot(Response),
             qplot(Observed.m.z),
             qplot(Charge),
             qplot(Matched.1st.Gen.Primary.Ions),
             qplot(X..Matched.1st.Gen.Primary.Ions....),
             qplot(Assigned.intensity....),
             ncol = 3)

# Visualize the boxplots for each variable in dataframe
grid.arrange(get_simple_boxplot(df2, Sequence.start, 'Sequence.start'),
             get_simple_boxplot(df2, Sequence.end, 'Sequence.end'),
             get_simple_boxplot(df2, Observed.mass..Da., 'Observed.mass..Da.'),
             get_simple_boxplot(df2, Observed.m.z, 'Observed.m.z'),
             get_simple_boxplot(df2, Mass.error..ppm., 'Mass.error..ppm.'),
             get_simple_boxplot(df2, Observed.RT..min., 'Observed.RT..min.'),
             get_simple_boxplot(df2, Response, 'Response'),
             get_simple_boxplot(df2, Observed.m.z, 'Observed.m.z'),
             get_simple_boxplot(df2, Charge, 'Charge'),
             get_simple_boxplot(df2, Matched.1st.Gen.Primary.Ions, 'Matched.1st.Gen.Primary.Ions'),
             get_simple_boxplot(df2, X..Matched.1st.Gen.Primary.Ions...., 'X..Matched.1st.Gen.Primary.Ions....'),
             get_simple_boxplot(df2, Assigned.intensity...., 'Assigned.intensity....'),
             ncol = 3)

# Unload dataframe 2 from memory
detach(df2)

# Create new varibale called as 'Sequence.length' which is diff of End and Start of the Sequence
df2['Sequence.length'] <-  df2$Sequence.end - df2$Sequence.start


# Visualize the corr plot of the data
cor_data = df2[sapply(df2, function(x) !is.character(x))]
corr_matrix <- cor(cor_data)
corrplot(corr_matrix, method = 'number', type = "lower")
rm(corr_matrix)
rm(df2)

# Create new caterogical varibale called as 'intensity' from Assigned.intensity....
cor_data <- mutate(cor_data, intensity = ifelse(ceiling(Assigned.intensity....) %in% 0:20, "Very Low",
                                      ifelse(ceiling(Assigned.intensity....) %in% 21:40, "Low",
                                             ifelse(ceiling(Assigned.intensity....) %in% 41:60, "Moderate",
                                                    ifelse(ceiling(Assigned.intensity....) %in% 61:80, "High",
                                                           ifelse(ceiling(Assigned.intensity....) %in% 81:100, "Very High", "NA"))))))

# Factorise the intensity column from caterogical to factor class
cor_data$intensity = factor(cor_data$intensity) 

# Remove redundant columns Sequence.start, Sequence.end, Assigned.intensity....
cor_data$Assigned.intensity.... <- NULL
cor_data$Sequence.end <- NULL
cor_data$Sequence.start <- NULL

# Visualize the histogram of the intensity
ggplot(cor_data, aes(intensity)) + 
  geom_bar() +
  ggtitle("Histogram for Intensity Type of the Proteins") +
  xlab("Intensity") +
  ylab("Count") +
  theme(axis.title.x = element_text(color = "DarkGreen", face = "bold.italic", size = 10),
        axis.title.y = element_text(color = "DarkGreen", face = "bold.italic", size = 10),
        axis.text.x = element_text(size = 10),
        axis.text.y = element_text(size = 10),
        plot.title = element_text(color = "Black", face = "bold.italic", hjust = 0.5, size = 15),
        panel.border = element_rect(colour = "black", fill=NA, size=1))

# Split the dataframe into two Training & Validation sets
train <- sample(nrow(cor_data), 0.7 * nrow(cor_data), replace = FALSE)
TrainSet2 <- cor_data[train,]
ValidSet2 <- cor_data[-train,]
rm(train)
rm(cor_data)

############################################################################################
# Decision Tree

#install.packages("rpart")
#install.packages("caret")
#install.packages("e1071")
library(rpart)
library(caret)
library(e1071)

# We will compare model 1 of Random Forest with Decision Tree model
dt_model = train(intensity ~ ., data = TrainSet2, method = "rpart")
dt_model

predTrain = predict(dt_model, data = TrainSet2)
mean(predTrain == TrainSet2$intensity)
table(predTrain, TrainSet2$intensity)

# Running on Validation Set
predValid = predict(dt_model, newdata = ValidSet2)
mean(predValid == ValidSet2$intensity)
table(predValid, ValidSet2$intensity)

rm(dt_model)
rm(predValid)
rm(predTrain)

############################################################################################
# Random Forest

#install.packages("randomForest")
library(randomForest)

# Create a Random Forest model with fine tuned parameters
rf_model = randomForest(formula = intensity ~ ., data = TrainSet2, importance = TRUE, mtry = 3)
rf_model

# Predicting on train set
predTrain <- predict(rf_model, TrainSet2)
mean(predTrain == TrainSet2$intensity)
table(predTrain, TrainSet2$intensity)

# Predicting on Validation set
predValid <- predict(rf_model, ValidSet2, type = "class")
# Checking classification accuracy
mean(predValid == ValidSet2$intensity)
table(predValid, ValidSet2$intensity)

importance(rf_model)
varImpPlot(rf_model)

rm(rf_model)
rm(predValid)
rm(predTrain)

rm(TrainSet2)
rm(ValidSet2)
rm(get_simple_boxplot)
############################################################################################