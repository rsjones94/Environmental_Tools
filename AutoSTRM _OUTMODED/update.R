rm(list=ls())


# if you don't have a local copy of the package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)

# This series of scripts was last checked to work with R v3.3.0, 9/27/16




########### USER INPUT

mywd = "X:\\2017\\161702811 TSMP Monitoring\\161702811B_Chilogatee_MY3\\Field\\Pebbles" # filepath to folder where your data is

workXS = F # set these to true to update XS, profile and particle collections respectively
workPro = F
workPart = T

xssource ='NFMC_Stitch_Clean_xs' # where is the XS data coming from? (.xlsx files only)
xsout = 'NFMC_Stitch_Clean_xs_database' # where is the data going? If this file exists, update.R will attempt to update the database. If it does not, a new database will be created.

prosource ='NFMC_Stitch_Clean_pro' # where is the profile data coming from? (.xlsx files only)
proout = 'NFMC_Stitch_Clean_pro_database' # where is the data going? If this file exists, update.R will attempt to update the database. If it does not, a new database will be created.

partsource ='Pebble_Count' # where is the particle data coming from? (.xlsx files only)
partout = 'chil_pebble_database' # where is the data going? If this file exists, update.R will attempt to update the database. If it does not, a new database will be created.

########### END USER INPUT






























### start

mywd = normalizePath(mywd)
setwd(mywd)

### let's do the xs data first

if (workXS){

  xsdatasource = paste(xssource,'.xlsx',sep='')
  xswritename = paste(xsout,'.xlsx',sep='')
  
  sourceWb = loadWorkbook(xsdatasource)
  sourceSheets = getSheets(sourceWb)
  #sourceNames = sourceSheets
  sourceNames = names(sourceSheets)
  
  
  outExists = file.exists(xswritename)
  
  if (outExists){
    
    print('Old XS data found. Appending new data...', quote = FALSE)
    
    outWb = loadWorkbook(xswritename)
    outSheets = getSheets(outWb)
    outNames = names(outSheets)
    
    notUpdated = names(outSheets)
    
    for (i in 1:length(sourceNames)){ 
      
      thisSheet = sourceNames[i]
      
      if (!(thisSheet %in% outNames)){ 
        warnString = paste('No match for worksheet', thisSheet, 'found in target file. Skipping to next sheet...')
        print(warnString, quote = FALSE)
        next
      }
      
      data = read.xlsx(xsdatasource, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      myr = data[1,8]
      
      elevations = data[,4]
      stations = data[,7]
      descriptions = data[,5]
      collectDate = data[1,14]
      
      addvecLength = length(elevations)-1
      myrVec = c(myr,rep(NA,addvecLength))
      myrVec[2] = collectDate
      
      dataDF = cbind.data.frame(myrVec,elevations,descriptions,stations)
      namerVec = c('Monitoring Year', 'Elevation', 'Description', 'Station')
      names(dataDF) = namerVec
      
      targetData = read.xlsx(xswritename, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      rowdiff = (nrow(targetData)-nrow(dataDF))
      if (rowdiff > 0) {
        
        naVec = rep(NA,ncol(dataDF))
        naMat = as.data.frame(matrix(rep(naVec),rowdiff,ncol=ncol(dataDF)))
        names(naMat) = namerVec
        dataDF = rbind(dataDF,naMat)
        
      } else if (rowdiff < 0){
        
        naVec = rep(NA,ncol(targetData))
        naMat = as.data.frame(matrix(rep(naVec),-rowdiff,ncol=ncol(targetData)))
        names(naMat) = names(targetData)
        targetData = rbind(targetData,naMat)
        
      }
      
      writeDF = cbind.data.frame(targetData,dataDF)
      names(writeDF)[1] = ''
      
      
      writeSheet = outNames[which(outNames==thisSheet)]
      notUpdated = notUpdated[-which(writeSheet==notUpdated)]
      
      removeSheet(outWb, sheetName = writeSheet)
      mySheet = createSheet(outWb, sheetName= writeSheet)
      addDataFrame(writeDF, sheet = mySheet, col.names = TRUE, row.names = FALSE, showNA = FALSE)
      
      saveWorkbook(outWb, xswritename)
      
      #write.xlsx(x = writeDF, file = xswritename, sheetName = thisSheet, row.names = FALSE, showNA = FALSE, append = TRUE)
      
    }
    
    if (length(notUpdated) != 0) {
      
      warningstring = paste('The following XS target sheets were not updated:')
      print(warningstring, quote = FALSE)
      print(notUpdated, quote = FALSE)
      
    } else {
      
      print('All XS target sheets updated.', quote = FALSE)
      
    }
    
  } else {
    
    print('No previous XS data found. Writing new file...', quote = FALSE)
    
    for (i in 1:length(sourceNames)) {
      
      thisSheet = sourceNames[i]
      data = read.xlsx(xsdatasource, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      reachType = data[1,11]
      reachMembership = data[1,12]
      lowBankHeight = data[1,13]
      myr = data[1,8]
      bankfull = data[1,9]
      xsName = data[1,10]
      collectDate = data[1,14]
      
      elevations = data[,4]
      stations = data[,7]
      descriptions = data[,5]
      
      addvecLength = length(elevations)-1
      myrVec = c(myr,rep(NA,addvecLength))
      
      myrVec[2] = collectDate
      dataDF = cbind.data.frame(myrVec,elevations,descriptions,stations)
      
      infoString = c('Bankfull', 'XS Name', 'Reach Type (p/r)', 'Reach Membership', 'Basin', 'Watershed', 'Drainage Area', 'Image Code', 'Image Folder', 'Prep Date (as string)', 'Crew', 'Low Bank Height', 'Chart Title')
      infoDat = rep(NA,length(infoString))
      infoDat[1] = bankfull
      infoDat[2] = xsName
      infoDat[3] = reachType
      infoDat[4] = reachMembership
      infoDat[12] = lowBankHeight
      
      diff = nrow(dataDF) - length(infoString)
      diffVec = rep(NA,abs(diff))
      
      ##### THIS LOGIC GATE NEEDS TO BE ADDED FOR ALL CODE SUBSECTIONS - HANDLES CASE WHERE THERE ARE FEWER SHOTS THAN USER INPUT FIELDS
      if (diff > 0) {
        infoString = c(infoString,diffVec)
        infoDat = c(infoDat,diffVec)
      } else if (diff < 0) {
        blankmat = matrix(NA,length(diffVec),ncol(dataDF))
        colnames(blankmat) = names(dataDF) 
        dataDF = rbind(dataDF,blankmat)
      }
      
      outDF = cbind.data.frame(infoString,infoDat,dataDF)
      names(outDF) = c('','XS Data', 'Monitoring Year', 'Elevation', 'Description', 'Station')
      
      write.xlsx(x = outDF, file = xswritename, sheetName = sourceNames[i], row.names = FALSE, showNA = FALSE, append = TRUE)
      
    }
    
    print('New XS collection created.', quote = FALSE)
    
  }
}







### now for profile data

if (workPro){
  
  prodatasource = paste(prosource,'.xlsx',sep='')
  prowritename = paste(proout,'.xlsx',sep='')
  
  sourceWb = loadWorkbook(prodatasource)
  sourceSheets = getSheets(sourceWb)
  #sourceNames = sourceSheets
  sourceNames = names(sourceSheets)
  
  
  outExists = file.exists(prowritename)
  
  if (outExists){
    
    print('Old profile data found. Appending new data...', quote = FALSE)
    
    outWb = loadWorkbook(prowritename)
    outSheets = getSheets(outWb)
    outNames = names(outSheets)
    
    notUpdated = names(outSheets)
    
    for (i in 1:length(sourceNames)){ 
      
      thisSheet = sourceNames[i]
      
      if (!(thisSheet %in% outNames)){ 
        warnString = paste('No match for worksheet', thisSheet, 'found in target file. Skipping to next sheet...')
        print(warnString, quote = FALSE)
        next
      }
      
      data = read.xlsx(prodatasource, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      myr = data[,11]
      collectDate = data[1,15]
      myr[2] = collectDate
      
      elevations = data[,6:9]
      stations = data[,10]
    
      
      addMat = matrix(NA,ncol=3,nrow=length(stations))
      
      dataDF = cbind.data.frame(myr,elevations,stations,addMat)
      namerVec = c('Monitoring Year', 'thw', 'ws', 'bkf', 'tob', 'Station', 'Cross-Sections', 'Structures', 'Adjustment')
      names(dataDF) = namerVec
      
      targetData = read.xlsx(prowritename, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      rowdiff = (nrow(targetData)-nrow(dataDF))
      if (rowdiff > 0) {
        
        naVec = rep(NA,ncol(dataDF))
        naMat = as.data.frame(matrix(rep(naVec),rowdiff,ncol=ncol(dataDF)))
        names(naMat) = namerVec
        dataDF = rbind(dataDF,naMat)
        
      } else if (rowdiff < 0){
        
        naVec = rep(NA,ncol(targetData))
        naMat = as.data.frame(matrix(rep(naVec),-rowdiff,ncol=ncol(targetData)))
        names(naMat) = names(targetData)
        targetData = rbind(targetData,naMat)
        
      }
      
      writeDF = cbind.data.frame(targetData,dataDF)
      names(writeDF)[1] = ''
      writeDF[1,ncol(writeDF)] = 0
      
      writeSheet = outNames[which(outNames==thisSheet)]
      notUpdated = notUpdated[-which(writeSheet==notUpdated)]
      
      removeSheet(outWb, sheetName = writeSheet)
      mySheet = createSheet(outWb, sheetName= writeSheet)
      addDataFrame(writeDF, sheet = mySheet, col.names = TRUE, row.names = FALSE, showNA = FALSE)
      
      saveWorkbook(outWb, prowritename)
      
      
    }
    
    if (length(notUpdated) != 0) {
      
      warningstring = paste('The following profile target sheets were not updated:')
      print(warningstring, quote = FALSE)
      print(notUpdated, quote = FALSE)
      
    } else {
      
      print('All profile target sheets updated.', quote = FALSE)
      
    }
    
  } else {
    
    print('No previous profile data found. Writing new file...', quote = FALSE)
    
    for (i in 1:length(sourceNames)) {
      
      thisSheet = sourceNames[i]
      data = read.xlsx(prodatasource, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)

      myr = data[,11]
      collectDate = data[1,15]
      myr[2] = collectDate
      
      elevations = data[,6:9]
      stations = data[,10]
      
      
      addMat = matrix(NA,ncol=3,nrow=length(stations))
      
      dataDF = cbind.data.frame(myr,elevations,stations,addMat)
      namerVec = c('Monitoring Year', 'thw', 'ws', 'bkf', 'tob', 'Station', 'Cross-Sections', 'Structures', 'Adjustment')
      names(dataDF) = namerVec
      
      infoString = c('Site Name', 'Reach')
      infoDat = rep(NA,length(infoString))
      
      diff = nrow(dataDF) - length(infoString)
      diffVec = rep(NA,diff)
      
      if (diff > 0) {
        infoString = c(infoString,diffVec)
        infoDat = c(infoDat,diffVec)
      } else if (diff < 0) {
        blankmat = matrix(NA,length(diffVec),ncol(dataDF))
        colnames(blankmat) = names(dataDF) 
        dataDF = rbind(dataDF,blankmat)
      }
      
      outDF = cbind.data.frame(infoString,infoDat,dataDF)
      names(outDF) = c('','Reach Data', 'Monitoring Year', 'thw', 'ws', 'bkf', 'tob', 'Station', 'Cross-Sections', 'Structures', 'Adjustment')
      outDF[1,ncol(outDF)] = 0
      
      write.xlsx(x = outDF, file = prowritename, sheetName = sourceNames[i], row.names = FALSE, showNA = FALSE, append = TRUE)
      
      
    }
    
    print('New profile collection created.', quote = FALSE)
    
  }
}




### finally particle data

if (workPart){
  
  partdatasource = paste(partsource,'.xlsx',sep='')
  partwritename = paste(partout,'.xlsx',sep='')
  
  sourceWb = loadWorkbook(partdatasource)
  sourceSheets = getSheets(sourceWb)
  sourceNames = sourceSheets
  sourceNames = names(sourceSheets)
  
  
  outExists = file.exists(partwritename)
  
  if (outExists){
    
    print('Old particle data found. Appending new data...', quote = FALSE)
    
    outWb = loadWorkbook(partwritename)
    outSheets = getSheets(outWb)
    outNames = names(outSheets)
    
    notUpdated = names(outSheets)
    
    for (i in 1:length(sourceNames)){ 
      
      thisSheet = sourceNames[i]
      
      if (!(thisSheet %in% outNames)){ 
        warnString = paste('No match for worksheet', thisSheet, 'found in target file. Skipping to next sheet...')
        print(warnString, quote = FALSE)
        next
      }
      
      data = read.xlsx(partdatasource, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      myr = data[,7]
      
      all = data[,1:4]
      
      dataDF = cbind.data.frame(myr,all)
      namerVec = c('Monitoring Year', 'Particle Type', 'Size', '', 'Count')
      names(dataDF) = namerVec
      dataDF = dataDF[1:24,]
      
      targetData = read.xlsx(partwritename, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      targetData = targetData[1:24,]
      
      rowdiff = (nrow(targetData)-nrow(dataDF))
      # if (rowdiff > 0) {
      #   
      #   naVec = rep(NA,ncol(dataDF))
      #   naMat = as.data.frame(matrix(rep(naVec),rowdiff,ncol=ncol(dataDF)))
      #   names(naMat) = namerVec
      #   dataDF = rbind(dataDF,naMat)
      #   
      # } else if (rowdiff < 0){
      #   
      #   naVec = rep(NA,ncol(targetData))
      #   naMat = as.data.frame(matrix(rep(naVec),-rowdiff,ncol=ncol(targetData)))
      #   names(naMat) = names(targetData)
      #   targetData = rbind(targetData,naMat)
      #   
      # }
      
      writeDF = cbind.data.frame(targetData,dataDF)
      names(writeDF)[1] = ''
      names(writeDF)[length(names(writeDF))-1] = ''
      
      writeSheet = outNames[which(outNames==thisSheet)]
      notUpdated = notUpdated[-which(writeSheet==notUpdated)]
      
      removeSheet(outWb, sheetName = writeSheet)
      mySheet = createSheet(outWb, sheetName= writeSheet)
      addDataFrame(writeDF, sheet = mySheet, col.names = TRUE, row.names = FALSE, showNA = FALSE)
      
      saveWorkbook(outWb, partwritename)
      
      
    }
    
    if (length(notUpdated) != 0) {
      
      warningstring = paste('The following particle target sheets were not updated:')
      print(warningstring, quote = FALSE)
      print(notUpdated, quote = FALSE)
      
    } else {
      
      print('All particle target sheets updated.', quote = FALSE)
      
    }
    
  } else {
    
    print('No previous particle data found. Writing new file...', quote = FALSE)
    
    for (i in 1:length(sourceNames)) {
      
      thisSheet = sourceNames[i]
      data = read.xlsx(partdatasource, sheetName = thisSheet, header = TRUE, stringsAsFactors=FALSE)
      
      myr = data[,7]
      
      all = data[,1:4]
      
      membership = data[1,6]
      xsname = data[1,5]
      
      
      dataDF = cbind.data.frame(myr,all)
      namerVec = c('Monitoring Year', 'Particle Type', 'Size', '', 'Count')
      names(dataDF) = namerVec
      dataDF = dataDF[1:24,]
      
      infoString = c('Site Name', 'XS Type (p/r)', 'Membership', 'XS Name')
      infoDat = rep(NA,length(infoString))
      infoDat[3] = membership
      infoDat[4] = xsname
      
      diff = nrow(dataDF) - length(infoString)
      diffVec = rep(NA,diff)
      
      if (diff > 0) {
        infoString = c(infoString,diffVec)
        infoDat = c(infoDat,diffVec)
      } else if (diff < 0) {
        blankmat = matrix(NA,length(diffVec),ncol(dataDF))
        colnames(blankmat) = names(dataDF) 
        dataDF = rbind(dataDF,blankmat)
      }
      
      outDF = cbind.data.frame(infoString,infoDat,dataDF)
      names(outDF) = c('','XS Data', 'Monitoring Year', 'Particle Type', 'Size (mm)', '', 'Count')
      
      write.xlsx(x = outDF, file = partwritename, sheetName = sourceNames[i], row.names = FALSE, showNA = FALSE, append = TRUE)
      
      
    }
    
    print('New particle collection created.', quote = FALSE)
    
  }
}