# Remindo QDNA conversion script

# First create Remindo_files folder in the working directory
# First create QDNA_files folder in the working directory
# Also add your candidates info to the working directory to replzce remindo ID with "code / kenmerk"

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
setwd("replace_with_your_working_directory")

# Install the readxl package if it is not installed already
if(!'readxl' %in% installed.packages()) { install.packages('readxl') }
# Load readxl package
library('readxl')

# List all files in subdirectory and assign to files vector
files <- list.files("Remindo_files")

# Subset to only suffix xlsx
files <- grep("\\.xlsx$", files, value = TRUE)

# Set file name. In the above command a list of multiple files could be ready for processing.
# You can now use for example files[1] to run the conversion for the first file name.
file_name = files["replace_with_number_of_file_to_convert"]

# Read data
results <- read_excel(paste("Remindo_files/", file_name, sep=''), col_names = T)

# Start cleaning up.

# Replace all white spaces with . 
names(results) <- gsub("\\s|-", "\\.", names(results))

names(results)

## Get rid of aggregated rows.

# Do not include all rows where 'Interactietype' contains '-'. This minus sign means that it is an aggregated row.
results <- subset(results, Interactietype != '-')

## Show item names
items <- unique(results$Vraag.ID)

## unique users
users <- unique(results$Gebruiker.ID)


# Replace Letters by numbers

## Manualy check number of answer options

# Use regular expression to find start ^ and end $ with capital letters and give unique answer number
number_of_mc_answers <- length(unique(results[grep("^[A-Z]$", results$Antwoord), "Antwoord"]))

# Replace eacht letter with a number. 
for (i in 1:number_of_mc_answers){
  results[grep(paste("^[",toupper(letters)[i],"]$",sep = ""), results$Antwoord), "Antwoord"] = i
}


# Start reshape of data
## Combine 'Vraag.ID' with 'Interactienummer' to get a unique id for each interaction.

results$Vraag.ID <- paste(results$Vraag.ID, results$Interactienummer, sep=".")

# Make answer key
## Grep the columns with correct answers and merge into answer key
unique_answers <- results[!duplicated(results$Vraag.ID),]

answer_correct_columns <- unique_answers[grep(names(results),pattern = "Mapping\\.[0-9]{1,2}\\.{3}Correct")]

## transform answer columns into one vector
answer_correct_columns <- data.matrix(answer_correct_columns) * matrix(rep(1:3,dim(answer_correct_columns)[1]),dim(answer_correct_columns)[1],
                                                                       dim(answer_correct_columns)[2],byrow = TRUE)
answer_key <- apply(answer_correct_columns,1,sum)
answer_key <- data.frame(Vraag.ID= unique_answers$Vraag.ID,answer_key = answer_key)

# Extract these columns and reshape long to wide.
## Select all colums that we are not interested in. 
drop.columns <- grep("[^Antwoord|Gebruiker\\.ID|Vraag\\.ID]", names(results), value = TRUE)
drop.columns[69] <- "Vragenbank.ID" 

# Reshape from long to wide. Use only columns of interest

scores <- reshape(results,
                  # There are multiple lines per user based on the number of items
                  # and we only need one line per user.
                  timevar  =   "Vraag.ID",
                  # We don't need all these columns so lets drop some
                  drop     = drop.columns,
                  # The two columns that we want are:
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
kandidaten <- read_excel("C:/Users/asasiad1/surfdrive/rprojects/remindo2qdnatool/kandidaten.xlsx", col_names = T)
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

# Write results to csv
write.table(qDNAdata, paste("QDNA_files/", sub("xlsx", "csv", "test.csv"), sep=""), 
            row.names = FALSE, 
            col.names = FALSE,
            sep       = ",")

# Write Remindo item ID's to file in corresponding order
write.table(sub("\\.1$", "", answer_key_df$Vraag.ID), paste("QDNA_files/", sub("xlsx", "_item_names.csv", "testmap"), sep=""), 
            row.names = TRUE, 
            col.names = FALSE,
            sep       = ",")

View(qDNAdata)


