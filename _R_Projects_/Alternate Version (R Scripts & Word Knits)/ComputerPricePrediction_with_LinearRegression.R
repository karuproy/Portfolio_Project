# Confusion Matrix for Disease Diagnosis using KNN Model
  
  
# Loading dataframe

c_prices <- read.csv("C:/Users/Arup/Documents/DS_ComputerConfigure.csv",stringsAsFactors = FALSE)  
str(c_prices)    


# Converting Yes to 1 & No to 0

c_prices<-c_prices[-1]

c_prices[c(which(c_prices$cd == "no")),]$cd<-0;            c_prices[c(which(c_prices$cd == "yes")),]$cd<-1
c_prices[c(which(c_prices$multi == "no")),]$multi<-0;      c_prices[c(which(c_prices$multi == "yes")),]$multi<-1
c_prices[c(which(c_prices$premium == "no")),]$premium<-0;  c_prices[c(which(c_prices$premium == "yes")),]$premium<-1

c_prices$cd
#c_price <- sapply( c_prices[,2:10], as.numeric ) 
str(c_price)


# Training and Testing Dataset

rn_train <- sample(nrow(c_prices), floor(nrow(c_prices)*0.7))

train <- c_prices[rn_train,c("price","ram","trend","speed")]
test <- c_prices[-rn_train,c("price","ram","trend","speed")]

model_ulm <- lm(price~ram, data=train) 
prediction <- predict(model_ulm, interval="prediction", newdata =test)



# Computing the Root Mean Square Error and Determining the Percentage of cases with less than 25% error.

errors <- prediction[,"fit"] - test$price
hist(errors)


rmse <- sqrt(sum((prediction[,"fit"] - test$price)^2)/nrow(test))
rel_change <- 1 - ((test$price - abs(errors)) / test$price)

pred25 <- table(rel_change<0.25)["TRUE"] / nrow(test)
paste("RMSE:", rmse)
paste("PRED(25):", pred25)



# Feature Selection (Foward Algorithm)

library(MASS) # stepwise regression

full   <- lm(price~ram+hd+speed+screen+ads+trend,data=c_prices)
null   <- lm(price~1,data=c_prices)

stepF  <- stepAIC(null, scope=list(lower=null, upper=full), direction= "forward", trace=TRUE)
summary(stepF)





library(leaps) # all subsets regression

subsets <- regsubsets(price~ram+hd+speed+screen+ads+trend,data=c_prices, nbest=1,)
sub.sum <- summary(subsets)
as.data.frame(sub.sum$outmat)



library(FNN)

dataset <- rbind(c_prices, c(7000,0,32,90,8,15,'no','no','yes',200,2))  
datasets<-dataset

datasets[c(which(datasets$cd == "no")),]$cd<-0;            datasets[c(which(datasets$cd == "yes")),]$cd<-1
datasets[c(which(datasets$multi == "no")),]$multi<-0;      datasets[c(which(datasets$multi == "yes")),]$multi<-1
datasets[c(which(datasets$premium == "no")),]$premium<-0;  datasets[c(which(datasets$premium == "yes")),]$premium<-1

dataset.numeric <- sapply( datasets[,2:11], as.numeric )   #Should convert data to numeric to use knn.reg
dataset.numeric <- as.data.frame(dataset.numeric)

prediction <- knn.reg(train = dataset.numeric[1:nrow(c_prices),-1], 
                      test = dataset.numeric[nrow(c_prices)+1,-1],
                      y = dataset.numeric[1:nrow(c_prices),]$price, k = 7 , algorithm="kd_tree")  

prediction$pred




