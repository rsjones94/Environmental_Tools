#rm(list=ls())


# if you don't have a local copy of package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)


########### USER INPUT

# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes
#mywd = "M:/2011/20110669-Jacobs Ladder/Monitoring/MY-03/Stream/" # where your file is located
#partsource = 'ladder_pebbles' # name of the file with the data
#partout = ''


########### END USER INPUT























### check out the input data
mywd = normalizePath(mywd)
setwd(mywd)
partsource = paste(partsource,'.xlsx',sep='')
sheets = getSheets(loadWorkbook(partsource))
numsheets = length(sheets)

outname = paste(partout,'.xlsx',sep='')


### set up for calculations

# make a function that finds dX via log-normal graphical interpolation

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

name = NULL
reach = NULL
outputlist = NULL
for (currentsheet in 1:numsheets) {
  
  data = read.xlsx(partsource, sheetIndex = currentsheet, header = TRUE)
  
  counts = data[,4]
  name[currentsheet] = toString(data[1,5])
  reach[currentsheet] = data[1,6]
  
  for (i in 1:length(counts)) {
    
    if (is.na(counts[i])) {
      
      counts[i] = 0
      
    }
    
  }
  
  total = sum(counts)
  
  percents = (counts/total)*100
  
  cumcounts = cumsum(counts)
  cumpercs = cumsum(percents)
  
  # convert particle diameter to phi
  phibins = -log2(bins)
  
  # calculate the Silt-Clay/Sand/Gravel/Cobble/Boulder/Bedrock breakdown
  
  clayp = sum(percents[bins >= sizes[1] & bins < sizes[2]])
  sandp = sum(percents[bins >= sizes[2] & bins < sizes[3]])
  gravp = sum(percents[bins >= sizes[3] & bins < sizes[4]])
  cobbp = sum(percents[bins >= sizes[4] & bins < sizes[5]])
  boulp = sum(percents[bins >= sizes[5] & bins < sizes[6]])
  bedrp = sum(percents[bins >= sizes[6]])
  
  classbreakdown = c(clayp,sandp,gravp,cobbp,boulp,bedrp)
  
  # calculate d16/d35/d50/d65/d84/d95
  dees = c(16,35,50,65,84,95)
  
  deevals = NULL
  for (i in 1:length(dees)) {
    
    deevals[i] = dX(cumpercs,dees[i])
    
  }
  
  outputmat = as.data.frame(matrix(NA,6,7)) 
  outputmat[,2] = deevals
  outputmat[,5] = classbreakdown
  
  deedown = round(deevals,2)
  deestring =  paste(deedown[1], deedown[2], deedown[3], deedown[4], deedown[5], deedown[6], sep = ' / ')
  
  rounddown = round(classbreakdown,1)
  breakdownstring =  paste(rounddown[1], rounddown[2], rounddown[3], rounddown[4], rounddown[5], rounddown[6], sep = ' / ')
    
  outputmat[1,7] = 'd16/d35/d50/d65/d84/d95 (mm)'  
  outputmat[2,7] = deestring
  
  outputmat[4,7] = 'SC/S/G/C/B/Bdrk (%)'
  outputmat[5,7] = breakdownstring
  
  dexes = paste('d',dees,sep='')
  seds = c('Silt/Clay', 'Sand', 'Gravel', 'Cobbles', 'Boulders', 'Bedrock')
  
  outputmat[,1] = dexes
  outputmat[,4] = seds
  
  columnnames = c('', 'dX (mm)', '', '', 'Sediment Breakdown (%)', '', '')
  names(outputmat) = c(columnnames)
  
  
  outputlist[[currentsheet]] = outputmat
  
  
}

names(outputlist) = name

for (i in 1:length(outputlist)) {
  
  write.xlsx(x = outputlist[[i]], file = outname, sheetName = names(outputlist)[i], row.names = FALSE, showNA = FALSE, append = TRUE)
  
}





