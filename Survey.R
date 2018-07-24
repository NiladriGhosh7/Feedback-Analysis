install.packages('xlsx')
library(xlsx)
library(sentimentr)
library(dplyr)
library(tidytext)
library(tm)
library(plyr)
library(stringr)
library(RColorBrewer)
library(wordcloud)
library(syuzhet)
library(ggplot2)

survey<- read.xlsx('45.xls',sheetName = '45',header = T)

irrelevant_columns <- c("StartDate","EndDate","RecordedDate","Progress","Duration (in seconds)",
                        "Q1","Q12","Q12_4_TEXT","Q13","Q17","Q32","Q_DL","UploadDate",
                        "ExpireDate")

Most_Relevant_Info<-survey[,-which(names(survey) %in% irrelevant_columns)]

Employee_Feedback <- subset(survey,select = c('Q33','Age','Gender','RecruitmentSource'))
colnames(Employee_Feedback)<-c('Feedback','Age','Gender','RecruitmentSource')
Employee_Feedback<-Employee_Feedback[,c('Age','Gender','RecruitmentSource','Feedback')]
Employee_Feedback<-na.omit(Employee_Feedback)
Employee_Feedback$RecruitmentSource<-factor(Employee_Feedback$RecruitmentSource,
                                            levels = 'null',labels = 'Not Specified')

Employee_Feedback<-Employee_Feedback[2:500,]
Employee_Feedback[!is.na(Employee_Feedback$RecruitmentSource),]
# Sentiment Analysis Of Employee Feedback


Employee_Feedback_Sentiments <- sentiment(as.character(Employee_Feedback$Feedback))

Employee_Feedback_with_sentiment<-Employee_Feedback

Employee_Feedback_Sentimentst<-Employee_Feedback_Sentiments[1:148,]
Employee_Feedback_with_sentiment$Sentiment<- Employee_Feedback_Sentimentst$sentiment

boxplot(Employee_Feedback_with_sentiment$Sentiment~Employee_Feedback$RecruitmentSource)


Number_Of_Responder<-length(which(Most_Relevant_Info$Q2_NPS_GROUP == 'Detractor'))+length(which(Most_Relevant_Info$Q2_NPS_GROUP == 'Promoter'))+length(which(Most_Relevant_Info$Q2_NPS_GROUP== 'Passive'))
Total_Promoter <-length(which(Most_Relevant_Info$Q2_NPS_GROUP == 'Promoter'))

Total_Detractor <-length(which(Most_Relevant_Info$Q2_NPS_GROUP == 'Detractor'))

Percentage_Promoter <- (Total_Promoter/Number_Of_Responder)*100
Percentage_Detractor <- (Total_Detractor/Number_Of_Responder)*100


#Clustering As per Satisfaction On Training

X <- survey['Q20']

wcss<-vector()
for (i in 1:10) wcss[i] <- sum(kmeans(X,i)$withinss)
plot(1:10, wcss,type = 'b',main = "Cluster Of Trainees",xlab = "Number Of Clusters",
     ylab= "WCSS")

#Analysis of Q4 I enjoy doing my job against Q2 Based on your experience, how likely are you to recommend HGS to a friend or colleague?
Most_Relevant_Info$Q2_NPS_GROUP=factor(Most_Relevant_Info$Q2_NPS_GROUP,levels=c('Detractor','Promoter','Passive'))
Most_Relevant_Info$Q4_NPS_GROUP=factor(Most_Relevant_Info$Q4_NPS_GROUP,levels=c('Detractor','Promoter','Passive'))
install.packages('randomForest')
library(randomForest)

subdata<-read.csv('SubData.csv')
subdata$Q2_NPS_GROUP<-factor(subdata$Q2_NPS_GROUP,levels = c('Promoter','Passive','Detractor'),labels=c(1:3))
subdata$Country<-factor(subdata$Country,levels = c('USA','Canada'),labels=c(1,2))
subdata$Gender<-factor(subdata$Gender,levels = c('Male','Female'),labels=c(1,2))
subdata$EmployeeType<-factor(subdata$EmployeeType,levels = c('Permanent','FULLTIME REGULAR','PARTTIME REGULAR'),labels=c(1,2,3))


subdata<-na.omit(subdata)

subdata$Age<-scale(subdata$Age)

classifier = randomForest(x=subdata[2:6],
                          y=subdata$Q2_NPS_GROUP,
                          ntree = 10)

#Visualising Random Forest :


#Logistic Regression

classifier = glm(formula = Q2_NPS_GROUP~.,family=binomial,data = subdata)
summary(classifier)

classifier1 = glm(formula = Q2_NPS_GROUP~Q27+Country,family=binomial,data = subdata)
summary(classifier1)



#Word Cloud :

library(SnowballC)

speech_corpus<-Corpus(VectorSource(survey$Q33))

speech_corpus_clean<-tm_map(speech_corpus,tolower)

speech_corpus_clean<-tm_map(speech_corpus_clean,removeNumbers)

speech_corpus_clean<-tm_map(speech_corpus_clean,removePunctuation)
#removing stop words like the ,and, I , we ,  etc..

speech_corpus_clean<-tm_map(speech_corpus_clean,removeWords,stopwords())

#to remove specific words like 'will' , 'work'

#speech_corpus_clean<-tm_map(speech_corpus_clean,removeWords,c('will','work'))

speech_corpus_clean<-tm_map(speech_corpus_clean,stripWhitespace)

speech_corpus_clean<-tm_map(speech_corpus_clean,stemDocument)

library(wordcloud)

wordcloud(speech_corpus_clean,min.freq = 2,random.order = "F")

#Creating Term Document Matrix 

Survey_tdm= TermDocumentMatrix(speech_corpus_clean)

Survey_tdm=as.matrix(Survey_tdm)
#Finding Frequent Terms

Freq_words_said = findFreqTerms(Survey_tdm,lowfreq = 10)

print(Freq_words_said)



#Associative Term

findAssocs(Survey_tdm,'work',0.3)

findAssocs(Survey_tdm,'work',0.5)


extract_sentiment_terms(as.character(survey$Q33))[1:20]


plot(sentiment(as.character(survey$Q33)))

# Unhappy Employees

Unhappy_Employees<- subset(Most_Relevant_Info,Q2_NPS_GROUP=='Detractor' &Q4_NPS_GROUP == 'Detractor'& Q8_NPS_GROUP == 'Detractor'& Q14_NPS_GROUP == 'Detractor' & Q20_NPS_GROUP == 'Detractor'&
                             Q25_NPS_GROUP == 'Detractor'& Q27_NPS_GROUP == 'Detractor', select = c('ExternalReference' , 'Q33' , 'Country' , 'State' , 'City' , 'BuildingName' , 'CompanyName' , 'Vertical' , 'Department' , 'EmployeeType' , 'Age' , 'Gender'))

Employee_Ratings <- subset(Most_Relevant_Info,select = c('ExternalReference','Q2','Q4','Q8','Q14','Q20','Q25','Q27','Q33','Country','State','City','BuildingName','CompanyName','Vertical','Department','Age','Gender'))

Employee_Ratings$Avg_Rating = rowMeans(Employee_Ratings[(2:500),c('Q2' , 'Q4' , 'Q8' , 'Q14' , 'Q20' , 'Q25' , 'Q27' )])


Emp_Rating = read.xlsx('Employee_Rating.xls',sheetName = 'Sheet1',header = T)

Emp_Rating_Less <- subset(Emp_Rating,Avg_Rating<mean(Emp_Rating$Avg_Rating))

#Word Cloud For Unhappy Employees

Emp_corpus<-Corpus(VectorSource(Emp_Rating_Less$Q33))

Emp_corpus_clean<-tm_map(Emp_corpus,tolower)

Emp_corpus_clean<-tm_map(Emp_corpus_clean,removeNumbers)

Emp_corpus_clean<-tm_map(Emp_corpus_clean,removePunctuation)
#removing stop words like the ,and, I , we ,  etc..

Emp_corpus_clean<-tm_map(Emp_corpus_clean,removeWords,stopwords())

cl_list=as.character(c("in","on","to","of","for","from","at","am","is","are","was","were","'s","we","'re","the","our","there","me","my","you",
                       "https://t","http://t","https","a","and","or","with","how","what","where","why","you","your","it","that","those","there","their",
                       "'ve","i","be","by","who","but","t","an","as"))

#to remove specific words like 'will' , 'work'

Emp_corpus_clean<-tm_map(Emp_corpus_clean,removeWords,cl_list)

Emp_corpus_clean<-tm_map(Emp_corpus_clean,stripWhitespace)

Emp_corpus_clean<-tm_map(Emp_corpus_clean,stemDocument)

library(wordcloud)

wordcloud(Emp_corpus_clean,min.freq = 2,random.order = "F")

survey$Q2=ifelse(is.na(survey$Q2),ave(survey$Q2,FUN = function(X)mean(X,NA.rm = TRUE)),survey$Q2)
X<-read.xlsx('X.xls',sheetName = 'Sheet1',header = T)
print(X)
na.omit(X)
wcss<-vector()
for (i in 1:10) wcss[i] <- sum(kmeans(clust,i)$withinss)
plot(1:10, wcss,type = 'b',main = "Cluster Of Clients",xlab = "Number Of Clusters",
     ylab= "WCSS")

#Applying kmeans in dataset
set.seed(29)

clust<-read.xlsx('clust.xls',sheetName = 'Sheet1',header = T)
kmeans <- kmeans(clust,5,iter.max = 300,nstart = 10)
cluster_profile<-t(kmeans$centers)
row_mean<-apply(cluster_profile, 1, mean)
col_mean<-apply(cluster_profile, 2, mean)

sorted_profile<-cluster_profile[order(row_mean)]


matplot(cluster_profile, type = c("b"), lwd=3)
hc=hclust(dist(X,method = 'euclidean'),method = 'ward.D')
y_hc=cutree(hc,4)

#Visualising the clusters

library(cluster)
clusplot(X,kmeans$cluster,shade = FALSE,color = TRUE,labels = 2,
         plotchar = FALSE,span = TRUE,main = "Cluster Of Employee Feedback",
         xlab = "Employees",ylab = "Employee Referral Probablity")

plot(X,col=kmeans$cluster,pch=20, cex=2)

dendrogram = hclust(d = dist(X, method = 'euclidean'), method = 'ward.D')
plot(dendrogram,
     main = paste('Dendrogram'),
     xlab = 'Customers',
     ylab = 'Euclidean distances')

# Fitting Hierarchical Clustering to the dataset
hc = hclust(d = dist(X, method = 'euclidean'), method = 'ward.D')
y_hc = cutree(hc, 2)

str(X)
X<-as.character(X)
# Visualising the clusters
library(cluster)
clusplot(X,
         y_hc,
         lines = 0,
         shade = TRUE,
         color = TRUE,
         labels= 2,
         plotchar = FALSE,
         span = TRUE,
         main = paste('Clusters of customers'),
         xlab = 'Annual Income',
         ylab = 'Spending Score')