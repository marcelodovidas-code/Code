#Ejemplo de regresión lineal con regularización
install.packages("glmnet")
library(glmnet)

data(mtcars)
df <- mtcars
dim(df)         #son 32 obs multivariadas con 11 variables

#Vamos a predecir el consumo gpm con todas las variables predictoras
X <- as.matrix(df[, -1])   # todas menos mpg
y <- df$mpg

#Ajuste del modelo con Ridge (alpha=0). Ajusta muchos modelos con varios lambdas.
ridge_model <- glmnet(X, y, alpha = 0)

# El lambda óptimo se selecciona por verificación cruzada (cross validation).
# Por omisión hace un k-fold cross validation con k = 10.
# En este caso, con 32 observaciones, selecciona al azar folds de:
# 3–3–3–3–3–3–3–3–3–4–4 elementos.
# Genera una secuencia de lambdas.
# Toma un lambda.
# Toma los folds 1–9, ajusta (entrena) el modelo,
# valida el modelo con el fold 10 y calcula el MSE.
# Luego ajusta el modelo con los folds 2–10,
# valida el modelo con el fold 1 y calcula el MSE.
# Así sucesivamente con el resto de los folds.
# Calcula el promedio (cvm) y la desviación estándar (cvsd)
# de los 10 MSE.
# Toma otro lambda y repite todo el procedimiento.
# lambda.min corresponde al cvm mínimo.
# lambda.1se es el lambda más grande cuyo cvm <= cvm(lambda.min) + cvsd(lambda.min),
# es decir, el modelo más simple dentro del rango aceptable.
# Recordar que mayor lambda implica coeficientes más penalizados
# y por lo tanto más empujados hacia cero.

cv_ridge <- cv.glmnet(X, y, alpha = 0)
plot(cv_ridge)           # en este gráfico se ve perfectamente el concepto  
cv_ridge$lambda.min      # λ que minimiza el error
cv_ridge$lambda.1se      # λ más simple dentro de 1 SE

ridge_final_min <- glmnet(X, y, alpha = 0, lambda = cv_ridge$lambda.min)
coef(ridge_final_min)

ridge_final_1se <- glmnet(X, y, alpha = 0, lambda = cv_ridge$lambda.1se)
coef(ridge_final_1se)   # se evidencia perfectamente el shrinkage. Conviene este en datasets pequeños como este (más robusto)

cbind(coef(ridge_final_min),coef(ridge_final_1se))  #para comparar


#Ahora vamos a aplicar Lasso
cv_lasso <- cv.glmnet(X, y, alpha = 1)
plot(cv_lasso)
coef(cv_lasso, s = "lambda.min")
coef(cv_lasso, s = "lambda.1se")
cbind(coef(cv_lasso, s = "lambda.min"),coef(cv_lasso, s = "lambda.1se"))
#Como en ambos casos se conservan las mismas variables, vemos que es robusto
#Hay multicolinealidad fuerte
#Arroja resultados más interpetables

cv_ridge$cvm[which.min(cv_ridge$cvm)]
cv_ridge$cvm[cv_ridge$lambda == cv_ridge$lambda.1se]

cv_lasso$cvm[which.min(cv_lasso$cvm)]
cv_lasso$cvm[cv_lasso$lambda == cv_lasso$lambda.1se]

#Prefiero Lasso con lambda.mínimo por ser más parsimonioso y tener un cvm comprendido en el rango de Ridge.

#Probemos Elastic Net

cv_en_05 <- cv.glmnet(X, y, alpha = 0.5)    #equilibrado
coef(cv_en_05, s = "lambda.min")
coef(cv_en_05, s = "lambda.1se")

cv_en_025 <- cv.glmnet(X, y, alpha = 0.25)  #parecido a Ridge
coef(cv_en_025, s = "lambda.min")
coef(cv_en_025, s = "lambda.1se")

cv_en_075 <- cv.glmnet(X, y, alpha = 0.75)  #parecido a Lasso
coef(cv_en_075, s = "lambda.min")
coef(cv_en_075, s = "lambda.1se")

cv_en_05$cvm[which.min(cv_en_05$cvm)]
cv_en_05$cvm[cv_en_05$lambda == cv_en_05$lambda.1se]

cv_en_025$cvm[which.min(cv_en_025$cvm)]
cv_en_025$cvm[cv_en_025$lambda == cv_en_025$lambda.1se]

cv_en_075$cvm[which.min(cv_en_075$cvm)]
cv_en_075$cvm[cv_en_075$lambda == cv_en_075$lambda.1se]

#Si mi objetivo es predecir, entonces elijo Ridge mín
#Si es parsimonia extrema, Lasso mín
#Si equilibrio entre predicción y parsimonia, Elastic Net alpha 0.75 lmín

#Esto ha sido un típico problema de machine learning supervisado
#El pipeline: datos con una variable objetivo, modelo que aprende, entrenamiento + validación, selección de hiperparámetros

