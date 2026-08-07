#Leer una tabla de una URL
#Por ej. https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/master/data/Telco-Customer-Churn.csv

url <- "https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/master/data/Telco-Customer-Churn.csv"
raw_df <- read_csv(url, show_col_types = FALSE)
raw_df

#Por omisión show_col_types es TRUE, y devuelve un mensaje informando sobre el tipo de dato en cada columna