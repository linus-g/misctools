# Written by: Linus Gustafsson
# First written: 26 February 2015
# Last modified by: Linus Gustafsson
# Date modified: 12 March 2015

# This script creates a list of all XLS files in a folder and attempts to download the most recent version of each file from
# the website of the Australian Bureau of Statistics, using the functionality described on the following website:
# http://www.abs.gov.au/ausstats/abs@.nsf/97adb482c0aba769ca2570460017d0e7/15d1d9943893067dca2572ec00157f40!OpenDocument#Bookmarking%20ABS%20Time%20Series%20Sprea

# Open directory selection dialogue and set current working directory to returned value
dataDirectory <- choose.dir()
setwd(dataDirectory)

# List of catalogue numbers in each of the ABS databases
meiList <- c("5206", "5302", "5368", "5609", "5625", "5671", "5676", "6202", "6302", "6345", "6354", "6401", "6416", "6427", "6457", "8501", "8731", "9314")
archiveList <- c("3101", "5204", "6291", "1364") 

# Names for temporary files [should update this to use random characters]
temp_file <- "temp_1212121212"
log_file <- "log_1212121212.log"

# Set up empty file connection for storing output from download.file
con <- NULL

# Create list of all files in the selected directory
files_to_download <- list.files(pattern = "xls")

# Check if user wants to proceed with replacing all ABS files in selected directory
checkProceed <- winDialog(type = c("yesno"),
        message=paste("This script will replace all Excel files with a name starting with a valid ABS catalogue number, in the following folder:\n\n",
        dataDirectory, "\n\nAre you sure you want to proceed?", sep=" "))

if(checkProceed=="YES"){
        
        for(i in 1:length(files_to_download)){
                
                if(substr(files_to_download[i], 1, 4) %in% c(meiList, archiveList)){
                        
                        if(substr(files_to_download[i], 1, 4) %in% meiList){
                                file <- paste("http://www.ausstats.abs.gov.au/ausstats/Meisubs.nsf/LatestTimeSeries/", gsub(".xls", "", gsub(".xlsx", "", files_to_download[i])),
                                              "/$FILE/", files_to_download[i], sep="")
                        }
                        else if(substr(files_to_download[i], 1, 4) %in% archiveList) {
                                file <- paste("http://www.ausstats.abs.gov.au/ausstats/ABS@Archive.nsf/LatestTimeSeries/", 
                                              gsub(".xls", "", gsub(".xlsx", "", files_to_download[i])), "/$FILE/", files_to_download[i], sep="")
                        }
                        
                        con <- file(log_file); sink(con, append=TRUE); sink(con, append=TRUE, type="message")
                        
                        download.file(file, temp_file, method="internal", mode="wb")
                        
                        sink(); sink(type="message")
                        
                        # Check if downloaded file is Excel file
                        a <- readLines(log_file)
                        
                        if(grepl("application/vnd.ms-excel", a[2])) {
                                
                                file.rename(temp_file, files_to_download[i])
                        }
                        
                }
                else{
                        # In future this should be changed to outputting a file with a list of updated and non-updated files.
                        winDialog(type = c("ok"),message=paste("Catalogue number not in list, please edit list of allowed catalogue numbers if file is from valid ABS catelogue\n\n", files_to_download[i]))
                }             
        }
                
        file.remove(log_file)
}
        
     