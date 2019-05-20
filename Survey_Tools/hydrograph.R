rm(list=ls())

library(readxl)
library(tibble)
#library(ggplot2)



########### USER INPUT
# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes

mywd = "X:\\2017\\161706201_WTSP\\Field\\Stream Gauge" # filepath to folder where your data is

datasource = 'masterSG' # filename of your input data (MUST be .xlsx)

########### END USER INPUT

mywd = normalizePath(mywd)
setwd(mywd)


datasource = paste(datasource,'.xlsx',sep='')

print('Reading XLSX files...')
gaugeData = read_xlsx(datasource, sheet = 1)
rainfall = read_xlsx(datasource, sheet = 2)
fieldStats = read_xlsx(datasource, sheet = 3)
axisDates = read_xlsx(datasource, sheet = 4)
print('Data successfully read.')

fsColNames = colnames(as.data.frame(fieldStats))
streamNames = fsColNames[2:length(fsColNames)]

axisDates = axisDates[[1]]

rainfall = as.data.frame(rainfall)

nGauges = ncol(gaugeData)/2
if (ncol(gaugeData) %% 2 != 0) {
  print('Error in gauge data. Make sure each gauge has a date-time and depth column. Execution halted.')
}

collate.hydro.data = function(position) {
  
  tdpos = position * 2
  timeAndDepth = gaugeData[,(tdpos-1):tdpos]
  
  sspos = position + 1
  streamStats = fieldStats[sspos]
  
  return(list(timeAndDepth,streamStats))
  
} 

create.stage = function(hyd,sensorHeight) {
  
  depth = as.data.frame(hyd)[,2]
  depth[depth<0] = 0
  
  stage = depth + sensorHeight
  
  return(stage)
  
}

create.discharge = function(stage,stats) {
  
  stats = stats[[1]][[1]]
  c3 = stats[5]
  c2 = stats[6]
  c1 = stats[7]
  
  get.flow = function(currentStage,c3,c2,c1) {
    flow = c3*currentStage^3 + c2*currentStage^2 + c1*currentStage
    return(flow)
  }
  
  Q = rep(0,length(stage))
  
  for (i in 1:length(Q)) {
    
    Q[i] = get.flow(stage[i],c3,c2,c1)
    
  }
  
  return(Q)
  
}

get.sensor.and.bkf = function(stats) {
  
  stats = stats[[1]][[1]]
  
  gaugeCapEl = stats[1]
  strLen = stats[2]
  thalwegEl = stats[3]
  bkfEl = stats[4]
  
  sensorHeight = gaugeCapEl - strLen - thalwegEl
  bkfHeight = bkfEl - thalwegEl
  
  return(c(sensorHeight,bkfHeight))
  
}


for (i in 1:nGauges) {
  
  print(paste('Processing gauge', i, 'of', nGauges))

  currentStream = collate.hydro.data(i)
  currentHyd = currentStream[1]
  currentStats = currentStream[2] # gauge cap el, str length, thalweg el, bkf el, c3, c2, c1
  
  sensorHeight = get.sensor.and.bkf(currentStats)[1]
  bkfHeight = get.sensor.and.bkf(currentStats)[2]
  bkfQ = create.discharge(bkfHeight,currentStats)
  
  stage = create.stage(currentHyd,sensorHeight)
  discharge = create.discharge(stage,currentStats)

  hydTable = as.data.frame(currentHyd)
  hydTable$stage = stage
  hydTable$discharge = discharge
  
  stageTab = hydTable[,c(1,3)]
  QTab = hydTable[,c(1,4)]
  
  xmin = min(stageTab[,1],na.rm = T)
  xmax = max(stageTab[,1],na.rm = T)
  dateLim = c(xmin,xmax)
  
  bkfCol = 'red'
  rainCol ='dodgerblue2'
  sensorCol = 'darkorange2'
  

  
  wd = 10.5
  ht = 8.5
  
  rainWidth = 2
  
  #stage plot
  filename = paste(streamNames[i],'_stage.pdf',sep='')
  pdf(filename,width=wd,height=ht,paper='special')
  par(mar = c(5, 4, 4, 4) + 0.3)
  
  plot(stageTab,type='l', lwd = 0, xlab = 'Date', ylab = 'Stage (ft)', main = paste(streamNames[i], ', ', 'Stage',sep=''), ylim = c(0,max(hydTable$stage,na.rm=T)), xlim = dateLim)
  abline(h=sensorHeight, col = sensorCol)
  abline(h=bkfHeight, col = bkfCol)
  abline(v=axisDates, col = 'grey')
  
  par(new=T)
  plot(rainfall, type = "h", axes = FALSE, bty = "n", xlab = "", ylab = "", col = rainCol, lwd = rainWidth, xlim = dateLim)
  axis(side=4, at = pretty((rainfall[,2])))
  mtext("Daily Rainfall (in)", side=4, line=3)
  
  par(new=T)
  plot(stageTab,type='l', lwd = 1, xlab = '', ylab = '', ylim = c(0,max(hydTable$stage,na.rm=T)), xlim = dateLim, axes = FALSE, bty = "n")
  
  legText = c('Stage','Rainfall','Bankfull Stage', 'Sensor Elevation')
  legCols = c('Black',rainCol,bkfCol,sensorCol)
  widths = c(1,rainWidth,1,1)
  legend('topleft', legend = legText, col = legCols, lwd = widths, ncol = 2, cex = 1, x.intersp = 0.5, bg = 'white')
  
  dev.off()
  
  # discharge plot
  filename = paste(streamNames[i],'_discharge.pdf',sep='')
  pdf(filename,width=wd,height=ht,paper='special')
  par(mar = c(5, 4, 4, 4) + 0.3)
  
  plot(QTab,type='l', lwd = 1, xlab = 'Date', ylab = 'Discharge (cu ft/s)', main = paste(streamNames[i], ', ', 'Discharge',sep=''), ylim = c(0,max(hydTable$discharge,na.rm=T)), xlim = dateLim)
  abline(h=bkfQ, col =bkfCol)
  abline(v=axisDates, col = 'grey')
  
  par(new=T)
  plot(rainfall, type = "h", axes = FALSE, bty = "n", xlab = "", ylab = "", col = rainCol, lwd = rainWidth, xlim = dateLim)
  axis(side=4, at = pretty((rainfall[,2])))
  mtext("Daily Rainfall (in)", side=4, line=3)
  
  par(new=T)
  plot(QTab,type='l', lwd = 0, xlab = '', ylab = '', ylim = c(0,max(hydTable$discharge,na.rm=T)), xlim = dateLim, axes = FALSE, bty = "n")
  
  legText = c('Discharge','Rainfall','Bankfull Discharge')
  legCols = c('Black',rainCol,bkfCol)
  widths = c(1,rainWidth,1)
  legend('topleft', legend = legText, col = legCols, lwd = widths, ncol =2, cex = 1, x.intersp = 0.5, bg = 'white')

  dev.off()
}

print('Finished processing.')

































