# Remindo QDNA conversion script

# In Remindo export the raw results through an admin account in the "beheer omgeving"

# Select the opleiding of interest.
# Select the exam of interest.
# Click tab "Exporteren".
# Click "Resultaten".
# Select period exam startdate and time, end period current time.
# Click "Download bestand".

# Set working directory
setwd("replace_with_your_working_directory")

# Install some packages
if(!'readxl' %in% installed.packages()) { install.packages('readxl') }
library('readxl')

# Set file name
file_name = "replace_with_file_name.xlsx"

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

# Use regular expression to find start ^ and end $ with A
results[grep("^A$", results$Antwoord), "Antwoord"] = 1
results[grep("^B$", results$Antwoord), "Antwoord"] = 2
results[grep("^C$", results$Antwoord), "Antwoord"] = 3

# Start reshape of data
## Combine 'Vraag.ID' with 'Interactienummer' to get a unique id for each interaction.

results$Vraag.ID <- paste(results$Vraag.ID, results$Interactienummer, sep=".")

# Extract these columns and reshape long to wide.
# Select all colums that we are not interested in. 
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
mc_questions <- grep("^1$|^2$|^3$", scores[1,])
mc_questions_names <- names(grep("^1$|^2$|^3$", scores[1,], value = TRUE))

# Disregard open questions
scores <- scores[, c(1, mc_questions)]

# Replace Remindo ID's with student numbers
# Read data
kandidaten <- read_excel("add_kandidaten_to_working_dir.xlsx", col_names = T)
kandidaten <- kandidaten[, c("ID", "Code/Kenmerk")]
scores <- merge(scores, kandidaten, by.x = "Gebruiker.ID", by.y = "ID", all.x = T)

View(scores)

# Get rid of prefix "Antwoord." to merge with correct answers
mc_question_id <- sub("Antwoord\\.", "", mc_questions_names)

# Subset to only mc questions
correct_results <- results[results$Vraag.ID %in% mc_question_id, c("Vraag.ID","Correct.beantwoord", "Antwoord" )]
# subset to only correct answers
correct_results <- subset(correct_results, Correct.beantwoord == 1)
# Get unique item answer combinations
answer_key <- unique(correct_results)

# Get rid of prefix "Antwoord." to merge with correct answers order
answer_order <- data.frame(Vraag.ID = sub("Antwoord\\.","", names(scores[,mc_questions])) )

# Merge correct order with correct answers
answer_key_df <- merge(answer_order, answer_key, by="Vraag.ID", all.x = T)

# Correct the order of the questions
answer_key_vector <- answer_key_df[answer_order$Vraag.ID, "Antwoord"]

# Create first aDNA answer key row
answer_key_vector <- c("0001", "1", answer_key_vector)

# Create qDNA answers matrix with student number on first row and version in second row.
answers <- cbind(studentnr = as.numeric(scores[, 'Code/Kenmerk']), 1, scores[,mc_questions])

# Combine answer key with results
qDNAdata <- rbind(answer_key_vector, answer_key_vector, answers)

# Write results to csv
write.table(qDNAdata, paste("QDNA_files/", sub("xlsx", "csv", file_name), sep=""), 
            row.names = FALSE, 
            col.names = FALSE,
            sep       = ",")

# Write Remindo item ID's to file in corresponding order
write.table(sub("\\.1$", "", answer_order$Vraag.ID), paste("QDNA_files/", sub("xlsx", "_item_names.csv", file_name), sep=""), 
            row.names = TRUE, 
            col.names = FALSE,
            sep       = ",")

View(qDNAdata)


