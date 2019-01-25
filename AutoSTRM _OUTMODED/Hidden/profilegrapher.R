#rm(list=ls())



# if you don't have a local copy of package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)
library(gridExtra)
library(RGraphics)
library(gridBase)
library(jpeg)


########### USER INPUT

# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes

#mywd = 'C:/Users/Randall.Jones/Desktop/pro_mock' # where your file is located (no slash at end)
#prosource = 'ladderpro'

########### END USER INPUT
































### check out the input data

plot = TRUE # true to save graphs, false to not

#currentsheet = 1
mywd = normalizePath(mywd)
setwd(mywd)
prosource = paste(prosource,'.xlsx',sep='')
sheets = getSheets(loadWorkbook(prosource))
numsheets = length(sheets)


for (currentsheet in 1:numsheets) {
  
  data = read.xlsx(prosource, sheetIndex = currentsheet, header = TRUE, stringsAsFactors=FALSE)
  
  sitename = data[1,2]
  reach = data[2,2]
  
  if (plot){
    savename = paste(sitename,', Reach ',reach,'.pdf',sep='')
    pdf(file = savename, width = 17, height = 11) # 12.7 x 7.7
  }
  
  cutData = data[,-(1:2)]
  numYears = ncol(cutData)/9
  colSeq = (0:(numYears-1))*9
  
  monYears = as.matrix(cutData[1,colSeq+1])
  thws = as.matrix(cutData[,colSeq+2])
  wss = as.matrix(cutData[,colSeq+3])
  bkfs = as.matrix(cutData[,colSeq+4])
  tobs = as.matrix(cutData[,colSeq+5])
  stations = as.matrix(cutData[,colSeq+6])
  crossSections = as.matrix(cutData[,colSeq+7])
  structures = as.matrix(cutData[,colSeq+8])
  adjustments = as.matrix(cutData[1,colSeq+9])
  
  for (i in 1:numYears) {
    stations[,i] = stations[,i] + as.numeric(adjustments[i])
  }
  
  
  lowest = min(thws, na.rm = TRUE)
  highest = max(bkfs, na.rm = TRUE)
  leftest = min(stations, na.rm = TRUE)
  rightest = max(stations, na.rm = TRUE)
  
  if(highest == -Inf) {highest = max(wss, na.rm=TRUE)}
  
  plotTitle = paste(sitename,' ','Stream Restoration Site', '\nLongitudinal Profile', '\n', reach, ' ', monYears[length(monYears)], sep = '')
  plot(NULL, ylim = c(lowest,highest), xlim = c(leftest,rightest), main = plotTitle, xlab = 'Station (ft)', ylab = 'Elevation (ft)')
  
  grid(nx = NA, ny = NULL, col = "lightgray", lty = 1, lwd = 1, equilogs = TRUE)
  
  saveCols = NULL
  saveWidths = NULL
  for (i in 1:numYears) {
    
    thiscol = i + 1
    saveCols[i] = thiscol
    
    if (i == numYears) {
      thick = 3
    } else {
      thick = 0.75
    }
    
    saveWidths[i] = thick
    
    lines(stations[,i], thws[, i], col = thiscol, lwd = thick)
    
  }
  
  thisstation = stations[,i]
  
  thisbkf = bkfs[,i]
  thiswater = wss[,i]
  thisxs= crossSections[,i]
  thisStructure = structures[,i]
  
  bkfmat = cbind(thisstation,thisbkf)
  bkfmat = na.omit(bkfmat)
  watermat = cbind(thisstation,thiswater)
  watermat = na.omit(watermat)
  xsmat = cbind(thisstation,thisxs)
  xsmat = na.omit(xsmat)
  structmat = cbind(thisstation,thisStructure)
  structmat = na.omit(structmat)
  
  
  bkfColor= 'brown'
  bkfShape = 22
  bkfThick = 2
  bkfcex = 1
  points(bkfmat, col = 'black', pch = bkfShape, lwd = bkfThick, bg = bkfColor, cex = bkfcex)
  
  waterColor = 'skyblue2'
  waterShape = 2
  waterThick = 2
  lines(watermat, col = waterColor, lty = waterShape, lwd = waterThick)
  
  crossColor= 'yellow'
  crossShape = 23
  crossThick = 2
  crosscex = 1.75
  points(xsmat, col = 'black', pch = crossShape, lwd = crossThick, bg = crossColor, cex = crosscex)
  
  structColor= 'orangered'
  structShape = 25
  structThick = 2
  structcex = 1.75
  points(structmat, col = 'black', pch = structShape, lwd = structThick, bg = structColor, cex = structcex)
  
  
  ## legend
  legText = c(monYears, 'Cross Sections', 'Structures', 'Bankfull', 'Water Surface')
  legLines = c(seq(from = 1, to = 1, length = numYears), NA, NA, NA, waterShape)
  legPoints = c(rep(NA, numYears), crossShape, structShape, bkfShape, NA)
  legWidths = c(saveWidths, crossThick, structThick, bkfThick, waterThick)
  legCols = c(saveCols,'black','black','black',waterColor)
  legBGs = c(rep(NA, numYears), crossColor, structColor, bkfColor, NA)
  legcex = c(rep(NA, numYears), crosscex, structcex, bkfcex, NA)
  
  
  #check to see if there are actually any bkfs or structures to display in legend
  if ((all(is.na(thisbkf)) | all(thisbkf == '')) & (all(is.na(thisStructure)) | all(thisStructure == ''))) {
    legText = legText[-(length(legText)-1)]
    legLines = legLines[-(length(legLines)-1)]
    legPoints = legPoints[-(length(legPoints)-1)]
    legWidths = legWidths[-(length(legWidths)-1)]
    legCols = legCols[-(length(legCols)-1)]
    legBGs = legBGs[-(length(legBGs)-1)]
    legcex = legcex[-(length(legcex)-1)]
    
    legText = legText[-(length(legText)-1)]
    legLines = legLines[-(length(legLines)-1)]
    legPoints = legPoints[-(length(legPoints)-1)]
    legWidths = legWidths[-(length(legWidths)-1)]
    legCols = legCols[-(length(legCols)-1)]
    legBGs = legBGs[-(length(legBGs)-1)]
    legcex = legcex[-(length(legcex)-1)]
  } else if (all(is.na(thisStructure)) | all(thisStructure == '')) {
    legText = legText[-(length(legText)-2)]
    legLines = legLines[-(length(legLines)-2)]
    legPoints = legPoints[-(length(legPoints)-2)]
    legWidths = legWidths[-(length(legWidths)-2)]
    legCols = legCols[-(length(legCols)-2)]
    legBGs = legBGs[-(length(legBGs)-2)]
    legcex = legcex[-(length(legcex)-2)]
  } else if (all(is.na(thisbkf)) | all(thisbkf == '')) {
    legText = legText[-(length(legText)-1)]
    legLines = legLines[-(length(legLines)-1)]
    legPoints = legPoints[-(length(legPoints)-1)]
    legWidths = legWidths[-(length(legWidths)-1)]
    legCols = legCols[-(length(legCols)-1)]
    legBGs = legBGs[-(length(legBGs)-1)]
    legcex = legcex[-(length(legcex)-1)]
  }
  
  legendX = leftest
  legendY = lowest+(highest-lowest)*0.5
  legend(legendX, legendY, legText, lty = legLines, lwd = legWidths, pt.cex = legcex, col = legCols, pt.bg = legBGs, pch = legPoints, bg = 'white')
  

  
  # calculate line of best fit for bankfulls
  
  if (!all(is.na(bkfmat[,2])) | !all(bkfmat[,2])) {
    bkfReg = lm(bkfmat[,2]~ bkfmat[,1])
    bint = summary(bkfReg)$coefficients[1,1]
    bslope = summary(bkfReg)$coefficients[2,1]
    
    bkfexes = c(min(thisstation, na.rm = TRUE)-500,max(thisstation, na.rm = TRUE)+500)
    bkfwhys = bslope*bkfexes + bint
    
    lines(bkfexes,bkfwhys)
    
    roundbslope = toString(round(bslope,4))
    roundbint = toString(round(bint,2))
  
    
    while(nchar(roundbslope) < 7) {
      roundbslope = paste(roundbslope,'0',sep='')
    }
    while(nchar(roundbint) < 6) {
      roundbint = paste(roundbint,'0',sep='')
    }
    
  
    
    whichBkf = round(nrow(bkfmat)*.7)
    bkfEq = bquote(S[bkf] == .(roundbslope) * x + .(roundbint))
    text(bkfmat[whichBkf,1],bkfmat[whichBkf,2]+(highest-lowest)*0.2, bkfEq)
  
  }
  
  # calculate line of best fit for water
  
  waterReg = lm(watermat[,2]~ watermat[,1])
  wint = summary(waterReg)$coefficients[1,1]
  wslope = summary(waterReg)$coefficients[2,1]
  

  
  roundwslope = toString(round(wslope,4))
  roundwint = toString(round(wint,2))
  
  
  while(nchar(roundwslope) < 7) {
    roundwslope = paste(roundwslope,'0',sep='')
  }
  while(nchar(roundwint) < 6) {
    roundwint = paste(roundwint,'0',sep='')
  }
  
  addWaterBestFit = FALSE
  if (addWaterBestFit) {
    
    abline(a = wint, b = wslope, col = waterColor)
    
  }
  
  # display average slope of water surface or eqn of line of best fit? avg or bf
  toDisplay = 'avg'
  

  if (toDisplay == 'bf') {
    
  waterEq = bquote(S[ws] == .(roundwslope) * x + .(roundwint))
  
  } else if (toDisplay == 'avg') {
    
  deltaX = watermat[nrow(watermat),1]-watermat[1,1]
  deltaY = watermat[nrow(watermat),2]-watermat[1,2]
  avgSlope = deltaY/deltaX
  roundAvg = round(avgSlope, 4)
  waterEq = bquote(S[ws] == .(roundAvg))
    
  }
  whichWater = round(nrow(watermat)*.4)
  text(watermat[whichWater,1],watermat[whichWater,2]+(highest-lowest)*(-0.3), waterEq)
  
  
  if (plot) {
    dev.off()
  }
  
}


















