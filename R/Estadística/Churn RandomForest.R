#install.packages(c("cli", "rlang", "ranger", "vip", "tidymodels"), repos = "https://cloud.r-project.org")
library(tidymodels)


# Cargamos el dataset usando la función base de R (sin paquetes extra)

url <- "https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/master/data/Telco-Customer-Churn.csv"
df_churn <- read.csv(url, stringsAsFactors = TRUE)

# Eliminamos la columna de ID (no aporta poder predictivo)
df_churn$customerID <- NULL

# 1. Dimensiones del dataset
dim(df_churn)
#Son 7043 observaciones multivariadas de 20 variables

# 2. Resumen rápido de la variable Churn
table(df_churn$Churn)                   #tabla de contigencia 1x2
prop.table(table(df_churn$Churn)) * 100 #ídem anterior pero proporción
#Se observa que casi el 27% de los clientes abandonan, Churn=Yes

#Tipo de datos de cada variable
str(df_churn)


#-------------------------------------------------------------------
#LIMPIEZA DE DATOS Y VALORES FALTANTES NA
# 1. Verificamos si hay NAs por columna
colSums(is.na(df_churn))                #sumamos las columnas con 0 y 1(NA)

# 2. Eliminamos las filas con NAs (son apenas 11 casos sobre 7043)
df_churn <- na.omit(df_churn)           #elimina todo renglón con al menos 1 NA

# 3. Verificamos la dimensión final
dim(df_churn)

#-------------------------------------------------------------------
#DIVISIÓN DE LOS DATOS: 75% TRAIN / 25%TEST SPLIT
set.seed(123)

#Realiza partición del dataset por muestreo aleatorio ESTRATIFICADO según Churn
#Divide en estratos Churn=Yes y Churn=No, saca 75% de los Yes y 75% de los No y los junta
data_split <- initial_split(df_churn, prop = 0.75, strata = Churn)
#data_split arroja training/testing/total 5273/1759/7032, pero es un objeto que guarda los índices

train_data <- training(data_split)        #la tabla de datos para el training
test_data  <- testing(data_split)         #la tabla de datos para el test

# Comprobamos las proporciones en el conjunto de entrenamiento
prop.table(table(train_data$Churn)) * 100 #efectivamente aprox. 27% Churn=Yes

#-------------------------------------------------------------------
#RECETA DE PREPROCESAMIENTO
#Se especifica qué transformaciones necesita el df antes de entrar al algoritmo
#%>% es el "pipe" pasa el resultado de la izquierda como argumento de la función de la derecha
#nos evita anidar funciones y/o definir funciones auxiliares
#recipe(Churn~.)  =Churn es la variable respuesta, el resto de las variables (.) son predictoras
receta_churn <- recipe(Churn ~ ., data = train_data) %>%
  step_dummy(all_nominal_predictors()) %>%   # Convertimos factores a variables 0/1
  step_zv(all_predictors())                  # Eliminamos variables con varianza cero si las hubiera
#Si una variable tiene el mismo valor en todos los registros (var=0), no es informativa y se elimina.

#-------------------------------------------------------------------
# Paso 5: Especificar el modelo Random Forest
modelo_rf <- rand_forest(trees = 500) %>%
  set_engine("ranger") %>%            
  set_mode("classification")                  #Yes/No, no es regresión numérica  

# Paso 6: Crear el Workflow y Entrenar
workflow_churn <- workflow() %>%          
  add_recipe(receta_churn) %>%                #toma la receta de preprocesamiento
  add_model(modelo_rf)                        #toma la especificación del algoritmo

# Entrenamos el modelo con los datos de train
fit_churn <- fit(workflow_churn, data = train_data)

#---------------------------------------------------------------------
#Ya entrenado el modelo, evaluamos su capacidad de generalización sobre el 25% de los datos reservados
#Generación de predicciones de clase y probabilidades 
predicciones <- predict(fit_churn, test_data) %>%
  bind_cols(predict(fit_churn, test_data, type = "prob")) %>%
  bind_cols(test_data %>% select(Churn))

# B. Matriz de Confusión
predicciones %>%
  conf_mat(truth = Churn, estimate = .pred_class)

# C. Métricas Globales (Accuracy y ROC AUC)
metricas <- metric_set(accuracy, roc_auc)
predicciones %>%
  metricas(truth = Churn, estimate = .pred_class, .pred_Yes)
#Accuracy es la proporción de clasificación correctas (81%)
#ROC AUC mide la capacidad del modelo para discriminar la probabilidad entre clientes que abandonan y los que no


#Si se acepta el modelo, el modelo se pone en producción.
#1. Guardar el modelo entrenado en un archivo liviano.

# Guardar el modelo entrenado a un archivo
saveRDS(fit_churn, "modelo_churn_v1.rds")

# Para usarlo mañana o en otro servidor, solo lo cargás:
# modelo_cargado <- readRDS("modelo_churn_v1.rds")
# Se guarda en el work directory que se consulta con
#getwd()



#---Script de scoring de clientes (generador de alertas)
#Un código que se ejecuta periódicamente (diario o mensual), toma la base de clientes activos, 
#calcula su riesgo y exporta la lista de acción para el equipo comercial o el CRM.
library(dplyr)
library(readr)

# 1. Cargar modelo guardado
modelo <- readRDS("modelo_churn_v1.rds")

# Supongamos que recibimos la base diaria/mensual actualizada
nuevos_clientes <- read_csv("clientes_activos_hoy.csv")

# Calcula las probabilidades
resultados <- nuevos_clientes %>%
  bind_cols(predict(modelo, nuevos_clientes, type = "prob"))

#Filtra, Ordena, exporta la lista operativa
lista_priorizada <- resultados %>%
  rename(probabilidad_churn = .pred_Yes) %>%
  filter(probabilidad_churn >= 0.50) %>%   # Regla de corte: riesgo medio/alto
  arrange(desc(probabilidad_churn)) %>%    # Los más críticos primero
  select(customerID, MonthlyCharges, Contract, probabilidad_churn)

write_csv(lista_priorizada, "gestion_retencion_hoy.csv")

#3 Exposición como API en tiempo real (Servicio web)






# 2. Predecir probabilidad sobre la base activa
clientes_riesgo <- df_churn %>%
  bind_cols(predict(modelo, df_churn, type = "prob")) %>%
  select(MonthlyCharges, tenure, Contract, .pred_Yes) %>%
  rename(probabilidad_churn = .pred_Yes) %>%
  filter(probabilidad_churn >= 0.50) %>%   # Filtrar clientes en riesgo medio/alto
  arrange(desc(probabilidad_churn))        # Ordenar de mayor a menor riesgo

# 3. Exportar la lista operativa
write_csv(clientes_riesgo, "clientes_en_riesgo_hoy.csv")
