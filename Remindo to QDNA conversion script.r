# Remindo QDNA conversion script

# First create Remindo_files folder in the working directory
# First create QDNA_files folder in the working directory
# Also add your candidates info to the working directory to replace remindo ID with "code / kenmerk"

# In Remindo export the raw results through an admin account in the "beheer omgeving"

# Select the opleiding of interest.
# Select the exam of interest.
# Click tab "Exporteren".
# Click "Resultaten".
# Select period exam startdate and time, end period current time.
# Click "Download bestand".

# Clear memory
rm(list=ls())

# Set working directory
setwd("Fill in your working directory")

# Install the readxl package if it is not installed already
if(!'readxl' %in% installed.packages()) { install.packages('readxl') }

# Load readxl package
library('readxl')

# List all xlsx files in "Remindo_files" and assign to files vector
files <- list.files("Remindo_files",pattern = "\\.xlsx$", full.names = TRUE)

# Choose file to convert from files
print(files)
file_name = files[3]

# Read data
results <- read_excel(file_name, col_names = T)

##Cleanup of read in data.

# Change names of columns such that spaces are replaced with dots 
names(results) <- gsub("\\s|-", "\\.", names(results))

## Get rid of aggregated rows.

# Do not include rows where 'Interactietype' contains '-'. 
# '-' stands for aggregated rows.
results <- subset(results, Interactietype != '-')

## Show item names
items <- unique(results$Vraag.ID)

## unique users
users <- unique(results$Gebruiker.ID)

# Use regular expression to find answers that consist of a single capital letter.
# Then count the unique results to determine amount of possible mc answers. 
number_of_mc_answers <- length(unique(results[grep("^[A-Z]$", results$Antwoord), "Antwoord"]))

# Replace answer letters with numbers for easier computations later on  
for (i in 1:number_of_mc_answers){
  results[grep(paste("^[",toupper(letters)[i],"]$",sep = ""), results$Antwoord), "Antwoord"] = i
}


# Start reshape of data
## Combine 'Vraag.ID' with 'Interactienummer' to get a unique id for each interaction.

results$Vraag.ID <- paste(results$Vraag.ID, results$Interactienummer, sep=".")

# Make answer key
## Create variable 
unique_answers <- results[!duplicated(results$Vraag.ID),]
#Determine columns that are returned if answer is correct
answer_correct_columns <- unique_answers[grep(names(results),pattern = "Mapping\\.[0-9]{1,2}\\.{3}Correct")]

## transform answer columns into one vector
answer_correct_columns <- data.matrix(answer_correct_columns) * matrix(rep(1:number_of_mc_answers,dim(answer_correct_columns)[1]),dim(answer_correct_columns)[1],
                                                                       dim(answer_correct_columns)[2],byrow = TRUE)
answer_key <- apply(answer_correct_columns,1,sum)
answer_key <- data.frame(Vraag.ID= unique_answers$Vraag.ID,answer_key = answer_key)

# Extract these columns and reshape long to wide.
## Select all colums that we are not interested in. 
# drop.columns <- grep("[^Antwoord|Gebruiker\\.ID|Vraag\\.ID]", names(results), value = TRUE)
# drop.columns[length(drop.columns) + 1] <- "Vragenbank.ID" 

# Reshape from long to wide. Use only columns of interest

scores <- reshape(results[c("Gebruiker.ID","Vraag.ID", "Antwoord")],
                  # There are multiple lines per user based on the number of items
                  # and we only need one line per user.
                  timevar  =   "Vraag.ID",
                  idvar    = "Gebruiker.ID",
                  # By reshaping the dataframe form long to wide we end up with 
                  # the columns we want.
                  direction =  "wide" )

# Determine mc questions
# Also manual check
mc_questions_grep <- paste("^[", paste(1:number_of_mc_answers,collapse = ""),"]$",sep = "")
mc_questions <- grep(mc_questions_grep, scores[1,])
mc_questions_names <- names(grep(mc_questions_grep, scores[1,], value = TRUE))

# Disregard open questions
scores <- scores[, c(1,mc_questions)]

# Replace Remindo ID's with student numbers
# Read data
kandidaten <- read_excel("FILL IN THE FILE WITH USERS", col_names = T)
kandidaten <- kandidaten[, c("ID", "Code/Kenmerk")]
scores <- merge(scores, kandidaten, by.x = "Gebruiker.ID", by.y = "ID", all.x = TRUE)

View(scores)

# Get rid of prefix "Antwoord." to merge with correct answers order
answer_order <- data.frame(Vraag.ID = sub(x = names(scores),pattern = "Antwoord\\.",replacement = ""))

# Merge correct order with answer key, dropping all non MC questions from answer_key
answer_key_df <- merge(answer_order, answer_key, by.x="Vraag.ID", all.x = FALSE)

# sort the answer key to match order it appears in scores.
answer_key_vector <- na.omit(answer_key_df[answer_order$Vraag.ID, "answer_key"])

# Create first aDNA answer key row
answer_key_vector <- c("0001", "1", answer_key_vector)

# Create qDNA answers matrix with student number on first column and version in second column
answers <- cbind(studentnr = as.numeric(scores[, 'Code/Kenmerk']), 1, scores[grep(names(scores), pattern ="Antwoord")])

# Combine answer key with results
qDNAdata <- rbind(answer_key_vector, answer_key_vector, answers)

#Replace NA with 9 for teleform

qDNAdata[is.na(qDNAdata)] <- 9 
# Write results to csv
write.table(qDNAdata, paste("QDNA_files/", sub("xlsx", "csv", gsub("Remindo_files/",replacement = "",x = file_name)), sep=""), 
            row.names = FALSE, 
            col.names = FALSE,
            sep       = ",")

# Write Remindo item ID's to file in corresponding order
write.table(sub("\\.1$", "", grep(x = answer_order$Vraag.ID,"\\.1$",value = TRUE)), paste("QDNA_files/", sub("xlsx", "_item_names.csv", gsub("Remindo_files/",replacement = "",x = file_name)), sep=""), 
            row.names = TRUE, 
            col.names = FALSE,
            sep       = ",")

View(qDNAdata)


