#rm(list=ls())



# if you don't have a local copy of package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)
library(gridExtra)
library(RGraphics)
library(gridBase)
library(jpeg)
library(Matrix)

########### USER INPUT

# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes

#mywd = 'C:/Users/Randall.Jones/Desktop/pebble_mock' # where your file is located (no slash at end)
#partsource = 'ladderpebble'

########### END USER INPUT






























plot = TRUE # true to save graphs, false to not
partsource = paste(partsource,'.xlsx',sep='')


dX = function(dist,X) { # take a distribution dist and the percentile X
  
  granu = 100000 # number of interpolations between bins
  
  
  greater = which(dist >= X)
  lesser = which(dist <= X)
  
  upbound = greater[1]
  lowbound = lesser[length(lesser)]
  
  
  upphi = phibins[upbound]
  lowphi = phibins[lowbound]
  
  if (length(upphi) == 0) {
    return(2048)
    break
  }
  if (length(lowphi) == 0) {
    return(0.062)
    break
  }
  
  if(upbound == lowbound) {
    
    exactphi = phibins[upbound]
    grainvalue = 2^(-exactphi)
    return(grainvalue)
    break
    
  }
  
  phirange = c(upphi, lowphi)
  percrange = c(dist[upbound], dist[lowbound])
  
  interp = approx(phirange, percrange, n = granu)
  
  whys = interp$y
  exes = interp$x
  
  closest = which(abs(whys-X)==min(abs(whys-X)))
  
  phivalue = exes[closest][1]
  
  grainvalue = 2^(-phivalue)
  
  return(grainvalue)
  
}


# ranges for different pebble classes 
SCr = c(0.01)
SDr = c(0.062)
GRr = c(2)
CBr = c(64)
BDr = c(256)
RKr = c(2048)

sizes = c(SCr,SDr,GRr,CBr,BDr,RKr)

# size bins
s1 = 0.01
s2 = .062
s3 = .125
s4 = .25
s5 = .5
s6 = 1
s7 = 2
s8 = 4
s9 = 5.7
s10 = 8
s11 = 11.3
s12 = 16
s13 = 22.6
s14 = 32
s15 = 45
s16 = 64
s17 = 90
s18 = 128
s19 = 180
s20 = 256
s21 = 362
s22 = 512
s23 = 1024
s24 = 2048

bins = c(s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,s15,s16,s17,s18,s19,s20,s21,s22,s23,s24)









### check out the input data

#currentsheet = 1
mywd = normalizePath(mywd)
setwd(mywd)
sheets = getSheets(loadWorkbook(partsource))
numsheets = length(sheets)


for (currentsheet in 1:numsheets) {
  
  data = read.xlsx(partsource, sheetIndex = currentsheet, header = TRUE, stringsAsFactors=FALSE)
  
  sitename = data[1,2]
  type = data[2,2]
  membership = data[3,2]
  xsname = data[4,2]

  reachtype = ''
  if (is.na(type)) {
    
    reachtype = ''
    
  } else if (type == 'r') {
    
    reachtype = ', Riffle'
    
  } else if (type == 'p') {
    
    reachtype = ', Pool'
    
  }
  
  if (plot){
    savename = paste(sitename,' Pebbles, ', xsname,'.pdf',sep='')
    pdf(file = savename, width = 11, height = 8.5) # 12.7 x 7.7
  }
  
  
  
  cutData = data[,-(1:2)]
  numYears = ncol(cutData)/5
  colSeq = (0:(numYears-1))*5
  
  monYears = as.matrix(cutData[1,colSeq+1])
  counts = as.matrix(cutData[1:24,colSeq+5])
  
  particles = cutData[1:24,2]
  mills = cutData[1:24,3]
  words = cutData[1:24,4]
  
  
  #layWidths = c(3,2.5,2.5,2.5)
  #layHeights = c(7,3)
  #layout(matrix(c(5,1,1,1,5,3,4,2), nrow = 2, ncol = 4, byrow = TRUE), widths = layWidths, heights = layHeights)
  
  layWidths = c(0.25,3,2.5-(.5/3),2.5-(.5/3),2.5-(.5/3),0.25)
  layHeights = c(0.4,7-0.4,3-0.4,0.4)
  layout(matrix(c(0,0,0,0,0,0,0,5,1,1,1,0,0,5,3,4,2,0,0,0,0,0,0,0), nrow = 4, ncol = 6, byrow = TRUE), widths = layWidths, heights = layHeights)
  
   
  plotTitle = paste('Particle Size Distribution','\n',sitename,'\n', xsname, reachtype, sep = '')
  plot(NULL, ylim = c(0,100), xlim = c(0.01,10000), xaxt='n', log='x', main = plotTitle, xlab = 'Particle Size (mm)', ylab = '')
  grid(nx = NA, ny = NULL, col = "lightgray", lty = 1, lwd = 1, equilogs = TRUE)
  axis(1, at=c(0.01,0.1,1,10,100,1000,10000), labels=c(0.01,0.1,1,10,100,1000,10000))
  title(ylab="% Finer Than (Cumulative)", mgp=c(2.25,1,0))
  
  
  counts[which(is.na(counts))] = 0
  
  # columnwise cumsums
  colCumSum = apply(counts, 2, cumsum)
  
  # columnwise sums
  colSum = apply(counts, 2, sum)
  
  # get column percents
  invDiag = Diagonal(x = (1/colSum))
  colPercents = (counts %*% invDiag)*100
  
  # get cumulative percents
  colCumPercents = (colCumSum %*% invDiag)*100 # essentially dividing each row of colCumSum by the respective element in colSum
  
  saveCols = NULL
  saveWidths = NULL
  for (i in 1:numYears) {
    
    thiscol = i
    # note - make color vector
    saveCols[i] = thiscol
    
    if (i == numYears) {
      thick = 3
    } else {
      thick = 0.75
    }
    
    saveWidths[i] = thick
    
    lines(c(bins,10000), c(colCumPercents[,i],100), col = thiscol, lwd = thick)
    
  }
  
  percents = colPercents[,i]
  cumpercs = colCumPercents[,i]
  count = counts[,i]
  
  
  # legend
  legText = c(monYears)
  legSyms = c(rep(1,numYears))
  legWidths = c(saveWidths)
  legCols = c(saveCols)
  
  legendX = 400
  legendY = 50
  legend(legendX, legendY, legText, lty = legSyms, lwd = legWidths, col = legCols, bg = 'white')
  
  
  
  # calculate the Silt-Clay/Sand/Gravel/Cobble/Boulder/Bedrock breakdown
  clayp = sum(percents[bins >= sizes[1] & bins < sizes[2]])
  sandp = sum(percents[bins >= sizes[2] & bins < sizes[3]])
  gravp = sum(percents[bins >= sizes[3] & bins < sizes[4]])
  cobbp = sum(percents[bins >= sizes[4] & bins < sizes[5]])
  boulp = sum(percents[bins >= sizes[5] & bins < sizes[6]])
  bedrp = sum(percents[bins >= sizes[6]])
  hardp = 0
  wooddetp = 0
  artificialp = 0
  
  classbreakdown = c(clayp,sandp,gravp,cobbp,boulp,bedrp,hardp,wooddetp,artificialp)
  
  typestring = c('Silt/Clay','Sand','Gravel','Cobble','Boulder','Bedrock','Hardpan','Wood/Det.','Artificial')
  
  rclassbreakdown = round(classbreakdown)
  
  if(sum(rclassbreakdown) != 100) {
    
    diff = 100-sum(rclassbreakdown)
    atmax = which(rclassbreakdown==max(rclassbreakdown))
    rclassbreakdown[atmax] = rclassbreakdown[atmax] + diff
    
  }
  
  breakdownstring = paste(rclassbreakdown,'%', sep = '')
  typeDF = cbind(typestring,breakdownstring)
  
  
  # convert particle diameter to phi so we can get dX values
  phibins = -log2(bins)
  
  dees = c(16,35,50,65,84,95)
  
  deevals = NULL
  for (i in 1:length(dees)) {
    
    deevals[i] = dX(cumpercs,dees[i])
    
  }
  
  

  deestring = signif(deevals,2)
  
  for (j in 1:length(deestring)) {
    
    alph = toString(deestring[j])
    
    if (as.numeric(deestring[j]) < 0.1) {
      while(nchar(alph) < 5) {
        alph = paste(alph,0,sep='')
      }
    } else if (as.numeric(deestring[j]) < 1){
      while(nchar(alph) < 4) {
        alph = paste(alph,0,sep='')
      }
    } else if (as.numeric(deestring[j]) < 10){
      if (nchar(alph) == 1) {
        alph = paste(alph,'.0',sep='')
      }
    }
    
    deestring[j] = alph
    
  }
  
  sizerstring = paste('D',dees,sep='')
  
  deeDF = cbind(sizerstring,deestring)
  
  
  
  ### get descriptive statistics
  
  d16 = deevals[1]
  d35 = deevals[2]
  d50 = deevals[3]
  d65 = deevals[4]
  d84 = deevals[5]
  d95 = deevals[6]
  
  mean = (d16*d84)^0.5
  dispersion = (d84/d16)^0.5
  skewness = log(mean/d50)/log(dispersion)
  
  statvec = c(mean,dispersion,skewness)
  
  statstring = round(statvec,1)
  statwords = c('Mean (mm)', 'Dispersion', 'Skewness')
  
  statDF = cbind(statwords,statstring)
  
  # big table
  countsum = sum(count)
  count[which(count==0)] = NA
  bigTable = cbind(particles,mills,words)
  bigTable = cbind(bigTable,count)
  bottomRow = c('','','Total',countsum)
  bigTable = rbind(bigTable,bottomRow)
  
  bigTable[which(is.na(bigTable))] = ''
  
  
  #### plot the tables
  
  
  bigtheme <- gridExtra::ttheme_default(
    core = list(fg_params=list(cex = .65)),
    colhead = list(fg_params=list(cex = .65, fontface = 4L)),
    rowhead = list(fg_params=list(cex = .65, fontface= 1L)))
  
  littletheme <- gridExtra::ttheme_default(
    core = list(fg_params=list(cex = .5)),
    colhead = list(fg_params=list(cex = .5)),
    rowhead = list(fg_params=list(cex = .5)))
  
  specialtheme <- gridExtra::ttheme_default(
    core = list(fg_params=list(cex = .6, fontface = 'bold')),
    colhead = list(fg_params=list(cex = .6)),
    rowhead = list(fg_params=list(cex = .6, fontface=1L)))
  
  
  typeGrob = tableGrob(breakdownstring, theme = bigtheme, cols = 'Type', rows = typestring)
  sizeGrob = tableGrob(deestring, theme = bigtheme, cols = 'Size (mm)', rows = sizerstring)
  distGrob = tableGrob(statstring, theme = bigtheme, cols = 'Size Distribution', rows = statwords)
  bigGrob = tableGrob(bigTable, theme = bigtheme, cols = c('Particle','Millimeter','','Count'), rows = NULL)
  
  typeGrob$widths = unit(rep(.5, ncol(typeGrob)), 'npc') # for XL, .9. For 8.5x11, .8
  typeGrob$heights = unit(rep(0.18, nrow(typeGrob)), 'npc') #rep(2.4/nrow(typeGrob)
  
  sizeGrob$widths = unit(rep(.6, ncol(sizeGrob)), 'npc') # for XL, .9. For 8.5x11, .8
  sizeGrob$heights = unit(rep(0.24, nrow(sizeGrob)), 'npc')
  
  distGrob$widths = unit(rep(.6, ncol(distGrob)), 'npc') # for XL, .9. For 8.5x11, .8
  distGrob$heights = unit(rep(0.24, nrow(distGrob)), 'npc')  
  
  bigGrob$widths = unit(rep(.305, ncol(bigGrob)), 'npc') # for XL, .9. For 8.5x11, .8
  bigGrob$heights = unit(rep(.039, nrow(bigGrob)), 'npc')
  
  
  
  
  
  # second base plot 
  frame()
  # Grid regions of current base plot (ie from frame)
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)
  
  grid.draw(typeGrob)
  popViewport(3)
  
  # third base plot 
  frame()
  # Grid regions of current base plot (ie from frame)
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)
  
  grid.draw(sizeGrob)
  popViewport(3)
  
  # fourth base plot 
  frame()
  # Grid regions of current base plot (ie from frame)
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)
  
  grid.draw(distGrob)
  popViewport(3)
  
  # fifth base plot 
  frame()
  # Grid regions of current base plot (ie from frame)
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)
  
  grid.draw(bigGrob)
  popViewport(3)
  
  dev.off()
  
}
  
  
  
  
  
  
  
  
  # need to add trailing zeros for size and stat/dist tables
  
  
  
  