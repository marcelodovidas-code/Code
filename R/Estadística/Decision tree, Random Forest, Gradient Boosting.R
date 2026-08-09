#Modelos no lineales de machine learning

#Árbol de decisión
install.packages("rpart")
library(rpart)
tree_model <- rpart(mpg ~ ., data = mtcars)

install.packages("rpart.plot")
library(rpart.plot)
rpart.plot(tree_model)

pred_tree <- predict(tree_model, mtcars)
mean((mtcars$mpg - pred_tree)^2)   # MSE

#Random forest
install.packages("randomForest")
library(randomForest)
rf_model <- randomForest(mpg ~ ., data = mtcars, ntree = 500)
pred_rf <- predict(rf_model, mtcars)
mean((mtcars$mpg - pred_rf)^2)    # Impresionante 1.47!

#Cómo visualizar el bosque
#Muestra qué variables usa más el bosque, cuánto contribuyen a reducir el error, relevancia para la predicción
#Se ve la importancia de disp, wt, hp, cyl. Apareció disp. Random forest no penaliza multicolinealidad!
#Logra capturar interacciones y relaciones no lineales
varImpPlot(rf_model)
#Importancia numérica (cuánto mejora la pureza de los nodos)
importance(rf_model)
#Visualizar un árbol individual del bosque, por ej. el árbol 1
getTree(rf_model, k = 1)
#Partial Dependence Plots. Muestra cómo cambia la predicción cuando una variable cambia, ceteris paribus
install.packages("randomForest")
library(randomForest)
par(mfrow=c(2,2))
partialPlot(rf_model, mtcars, "disp")
partialPlot(rf_model, mtcars, "wt")
partialPlot(rf_model, mtcars, "hp")
partialPlot(rf_model, mtcars, "cyl")

#Gradient boosting
install.packages("gbm")
library(gbm)
gbm_model <- gbm(
  mpg ~ ., 
  data = mtcars,
  distribution = "gaussian",
  n.trees = 1000,
  interaction.depth = 3,
  shrinkage = 0.01,
  bag.fraction = 1,        # usa todos los datos
  n.minobsinnode = 3       # permite nodos pequeños
)

pred_gbm <- predict(gbm_model, mtcars, n.trees = 1000)
mean((mtcars$mpg - pred_gbm)^2)

par(mfrow=c(1,1))
summary(gbm_model)
par(mfrow=c(2,2))
plot(gbm_model, i = "wt")
plot(gbm_model, i = "disp")
plot(gbm_model, i = "hp")
plot(gbm_model, i = "cyl")
pretty.gbm.tree(gbm_model, i.tree = 1)

