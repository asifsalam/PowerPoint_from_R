# E-mail: asif.salam@hotmail.com

library(RDCOMClient)
library(plyr)
library(dplyr)
library(stringr)
library(rvest)

# IMDB's Clint Eastwood page
clint_url <- "http://www.imdb.com/name/nm0000142/"

# Save the page locally, so you don't load the web-site
local_path <- "C://Users//tksassa//Documents//20-GitProjects//2014//210-AutomatingPowerpoint//AutomatingPPT"

#Set your local path here
local_path <- "C://"
local_file <- paste0(local_path,"//Clint Eastwood - IMDb.html")
test_page <- html(local_file)
clint_page <- html(clint_url)

#film_selector <- ".filmo-row"
film_selector <- "#filmo-head-actor+ .filmo-category-section .filmo-row"
filmography <- clint_page %>% html_nodes(film_selector)

# Remove "TV Series" & "TV Movie" from the data
filmography <- filmography[-grep("TV Movie|TV Series",html_text(filmography))]

# Create the films data frame
films <- NULL
films$year <- filmography %>% html_nodes("span") %>% html_text() %>% str_trim()
films$title <- filmography %>% html_nodes("b a") %>% html_text()
films$url <- paste0("http://www.imdb.com",filmography %>% html_nodes("b a") %>% html_attr("href"))
films <- as.data.frame(films,stringsAsFactors=FALSE)

# Create the img sub-directory
if (!file.exists("img")) dir.create("img")
#Create an index so the dataframe can be sorted back.
films$index <- sprintf("%02d",seq_along(1:length(films$year)))

films$img_file <- paste0("img/img",films$index,".jpg")

# Extract the character name and add to dataframe
get_character <- function(film,filmography) {
    i <- as.integer(film$index)
    character_name <- filmography[i] %>% html_nodes(xpath=".//a[2]") %>% html_text()
    if (length(character_name)==0) {
        character_name <- filmography[[i]] %>% 
                            html_nodes(xpath="text()[preceding-sibling::br]") %>%
                            html_text() %>%
                            str_trim() %>%
                            str_replace("\n"," ")
    }
    return(character_name)
}

films$character_name <- daply(films,.(index),get_character,filmography)

# Loop through the films and download the poster image into the "img" subdirectory.
# If the poster is not found, flag the file name with 0.
for (i in 1:nrow(films)) {
    img_node <- html(films$url[i]) %>% html_nodes(xpath='//td[@id="img_primary"]//img')
    if (length(img_node)==0) {
        films$img_file[i] <- "img/img00.jpg"
        cat(i," : img file NOT FOUND: ",films$img_file[i],"\n")
    } 
    else {
        img_link <- html_attr(img_node,"src")
        cat(i," :",films$img_file[i]," : ",img_link,"\n")
        download.file(img_link,films$img_file[i],method="internal",mode="wb")
    }
}

# Check which of the files were not found and download them manually
films$title[which(films$img_file=="img/img00.jpg")]

# These images don't exist.  Download manually
films[55,"img_file"] <- "img/img55.jpg"
films[54,"img_file"] <- "img/img54.jpg"
films[52,"img_file"] <- "img/img52.jpg"

# Correct this title (appears with strange characters because of my locale)
films[40,"title"] <- "Kelly's Heroes"

# Save the data frame
write.table(films,file="eastwood_films.csv",append=FALSE,quote=TRUE,sep="\t",row.names=FALSE)

#------------------------------------- Films dataframe done -------------------------------------#


# =====================================Create a dataframe for box office earnings data ========#
##  Get box office earnings data for the films
clint_box_office_url <- "http://www.boxofficemojo.com/people/chart/?id=clinteastwood.htm"
bo_page <- html(clint_box_office_url)
# Extract tables. The fourth table is the one we want, with adjusted box office returns
bo <- box_office_page %>% html_table(header=TRUE,fill=TRUE) %>% (function(x) {x[[4]]})

# Clean up dataframe and correct formats
names(bo) <- c("bo_rank","title_name","studio","adjusted_gross","unadjusted_gross","release_date")
bo$adjusted_gross <- as.numeric(gsub("[\\$\\,]","",bo$adjusted_gross))
bo$unadjusted_gross <- as.numeric(gsub("[\\$\\,]","",bo$unadjusted_gross))
bo$release_date <- strptime(bo$release_date,"%m/%d/%y")
bo$release_date[32] <- strptime("1975-06-15",format("%Y-%m-%d"))
bo$release_date <- correct_date(box_office$release_date)

# Create a key for joining dataframes using the film title
bo$key <- tolower(gsub("[^[:alnum:]]", "", bo$title))
films$key <- tolower(gsub("[^[:alnum:]]", "", films$title))
bo$key[34] <- "therookie"

# Create a dataframe for the box office gross for the movies
box_office <- left_join(select(bo,bo_rank,studio,adjusted_gross,key),select(films,year,title,index,key),by="key")

# Save the box_office data frame.
write.table(box_office,file="eastwood_box_office.csv",append=FALSE,quote=TRUE,sep="\t",row.names=FALSE)
