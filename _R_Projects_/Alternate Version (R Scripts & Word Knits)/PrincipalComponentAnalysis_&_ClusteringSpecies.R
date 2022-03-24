# Clustering Flower Species and Principle Component Analysis

  
# Load and Explore Data

data(iris)
head(iris,30)


# Relation between Sepal Length and Sepal Width

plot(scale(iris$Sepal.Length), scale(iris$Sepal.Width))
plot((iris$Sepal.Length - mean(iris$Sepal.Length)) / sd(iris$Sepal.Length))


# Applying Pricipal Component Analysis (PCA) on Independent Variables

iris_x <-iris[,1:4]

iris.pca.rawdata <- prcomp(iris_x, scale = FALSE, center= FALSE)
iris.pca.rawdata

summary(iris.pca.rawdata)

str(iris.pca.rawdata)


# Normalize the Data

iris.pca.normaldata <- prcomp(iris_x, scale = TRUE, center= TRUE)
iris.pca.normaldata

summary(iris.pca.normaldata)

str(iris.pca.normaldata)


# Variances of Components (Both Raw vs Normalized)

plot(iris.pca.rawdata, type = "l", main='with data normalization')
plot(iris.pca.normaldata, type = "l", main='with data normalization')


# Comparing Boxplots of Raw vs Normal Datasets

boxplot(iris.pca.rawdata$x, main='Raw Data Transformation')
boxplot(iris.pca.normaldata$x, main='Normal Data Transformation')


# 2-D Graph of PCA Components

biplot(iris.pca.rawdata, choices = 1:2, main='Raw')

biplot(iris.pca.normaldata, choices = 1:2, main='Normal')


# Clustering the Species in Groups

iris.pca.normaldata$x
iris2 <- cbind(iris, iris.pca.normaldata$x)

library(ggplot2)

ggplot(iris2, aes(PC1, PC2, col = Species, fill = Species)) +
  stat_ellipse(geom = "polygon", col = "black", alpha = 0.5) +
  geom_point(shape = 21, col = "black")


# PCA Coefficients

cor(iris[, 1:4], iris2[, 6:9])

