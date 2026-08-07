#Cómo hacer un split train/test con estratificación

#tabla ejemplo
url <- "https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/master/data/Telco-Customer-Churn.csv"
df <- read_csv(url, show_col_types = FALSE)

glimpse(df)
table(df$Churn)

#Estratificamos por Churn.
ind_train_Y<-which(df$Churn=="Yes")
ind_train_N<-which(df$Churn=="No")

set.seed(42)
p<-0.80
ind_train<- c(
sample(ind_train_Y,floor(p*length(ind_train_Y))),
sample(ind_train_N,floor(p*length(ind_train_N)))
)

train<-df[ind_train,]
test<-df[-ind_train,]


#Verificación que el split es representativo
table(train$Churn)/nrow(train)*100
table(test$Churn)/nrow(test)*100
