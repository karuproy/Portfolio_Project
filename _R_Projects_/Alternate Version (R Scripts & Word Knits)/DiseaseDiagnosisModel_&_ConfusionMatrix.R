#  Confusion Matrix Creation & Accuracy Calculation for Disease Diagnosis using KNN Model

  
install.packages("class")
install.packages("gmodels")
  
  
# Loading dataframe

prc <- read.csv("C:/Users/Arup/Documents/DS_DiseaseDiagnose.csv",stringsAsFactors = FALSE)  
str(prc)    


# Removing the ID variable (not required) from the data set.

prc <- prc[-1]


# The data set contains patients who have been diagnosed with either Malignant (M) or Benign (B) cancer.

prc$diagnosis <- factor(prc$diagnosis_result, levels = c("B", "M"), labels = c("Benign", "Malignant"))
round(prop.table(table(prc$diagnosis)) * 100, digits = 1)  # Percentage form rounded to 1 decimal place


# Normalizing numeric data

normalize <- function(x) {
  return ((x - min(x)) / (max(x) - min(x))) }

prc_n <- as.data.frame(lapply(prc[2:9], normalize)) # Normalize everything except the result

summary(prc_n) # Checking if normalized


# Trainnig and Testing the Dataset

prc_train <- prc_n[01:070,]
prc_test  <- prc_n[71:100,]

prc_train_labels <- prc[01:070, 1] # Target Variable is in Column-1 as label
prc_test_labels  <- prc[71:100, 1]


# KNN Modelling and Confusion Matrix 

library(class)
prc_test_pred <- knn(train = prc_train, test = prc_test,cl = prc_train_labels, k=10)

library(gmodels)
CrossTable(x=prc_test_labels, y=prc_test_pred, prop.chisq=FALSE)


#Measuring Accuracy

percentAccuracy <- 100*(7+15)/30; #TN+TP, whereas FN=0 & FP=8
percentAccuracy
