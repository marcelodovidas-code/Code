#Limpieza estándar de datos de una tabla

#tabla ejemplo
url <- "https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/master/data/Telco-Customer-Churn.csv"
df <- read_csv(url, show_col_types = FALSE)

#Visualización rápida de la tabla
glimpse(df)
dim(df)

#NA
#cuántos NAs tenemos en cada columna
df %>% is.na() %>% colSums
#elimina las filas que contengan al menos un NA
df<- df %>% na.omit()

#REGISTROS DUPLICADOS
#duplicated() recorre el df de arriba a abajo. Si el registro apareció por primera vez es FALSE, caso contrario TRUE
df %>% duplicated() %>% sum()
#elimina las filas duplicadas
df<-df %>% unique(.)          #a veces se requiere el punto . en el argumento

#CELDAS EN BLANCO
#elimina espacios en blanco innecesarios antes y después de strings
df %>% as.matrix() %>% trimws()%>% {. == ""} %>% colSums()
#elimina las filas que contengan al menos un blank
filas_ok <- rowSums(trimws(as.matrix(df)) == "") == 0
df <- df[filas_ok, ]



