#rm(list=ls())



# if you don't have a local copy of package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)


########### USER INPUT

# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes
#mywd = 'M:/2016/161600959-Stony Fork FDP/Field/Stream/raw data' # where your file is located (no slash at end)
#xssource = 'cleandataxs-c' # name of the file with the data (MUST end in .xlsx)
#xsout = ''

########### END USER INPUT


















### check out the input data
mywd = normalizePath(mywd)
setwd(mywd)
xssource = paste(xssource,'.xlsx',sep='')
sheets = getSheets(loadWorkbook(xssource))
numsheets = length(sheets)

outname = paste(xsout,'.xlsx',sep='')

### work with each set of data in sequence

statnames = c('Bkf Width', 'FPA Width', 'Bkf XS Area', 'Bkf Mean Depth', 'Bkf Max Depth', 'W/D Ratio', 'Entrenchment Ratio', 'Bank Height Ratio', 'Reach Membership')

istats = as.data.frame(matrix(0,length(statnames),numsheets)) # stats for individual cross sections
row.names(istats) = statnames

reachdata = NULL # will store what reach a XS belongs to
typevec = NULL
for (currentsheet in 1:numsheets) {

  data = read.xlsx(xssource, sheetIndex = currentsheet, header = TRUE, stringsAsFactors=FALSE)
  
  alts = data[,4] # elevation
  stations = data[,7] # station
  desc = data[,5]
  bank = data[1,9] # bankfull elevation
  xsname = toString(data[1,10]) # assign name of the cross section
  type = toString(data[1,11]) # pool or riffle?
  typevec[currentsheet] = type
  reachdata[currentsheet] = data[1,12] # reach this XS is a part of
  lowBankHeight = data[1,13]
  
  wherelb = grep('ltob', desc)
  whererb = grep('rtob', desc)
  
  
  #### now let's calculate the XS statistics
  bin = 0 # smoothing reach, in feet
  # 0 will turn smoothing off
  int = 0.001 # interpolation granularity, in feet
  tol = int*10 # maximum deviation for the program to identify a station as being the bankfull station
  # technically should be int/2 as this is the smallest tolerance that is 'guaranteed' to find a result, but does not seem to be the case so use int*10
  
  
  interp = approx(stations, alts, xout=seq(from = min(stations), to = max(stations),by=int))
  model = ksmooth(interp$x, interp$y, bandwidth=bin, x.points=interp$x)
  
  whys = model$y
  exes = model$x
  
  mindex = which.min(whys) # index at which the channel is deepest
  schanmin = exes[mindex] # station of channel minimum
  echanmin = whys[mindex] # elevation of channel minimum
  
  
  #identify the right bankfull station
  rbkind = NULL
  rbk = NULL
  useTOBasbkf = FALSE
  
  if (length(whererb) != 0 & useTOBasbkf) {
    
    rbk = stations[whererb]
    rbkind = which.min(abs(exes - rbk))
    
  } else {
    for (j in mindex:length(whys)) {
      
      if (abs(bank-whys[j]) < tol) {
        rbkind = j
        rbk = exes[j]
        break
      }
      
      if (j == length(whys)) {
        rbkind = j
        rbk = exes[j]
      }
      
    }
  }
  
  #identify the left bankfull station
  lbkind = NULL
  lbk = NULL
  
  if (length(wherelb) != 0 & useTOBasbkf) {
    
    lbk = stations[wherelb]
    lbkind = which.min(abs(exes - lbk))
    
  } else {
    for (j in mindex:1) {
      
      if (abs(bank-whys[j]) < tol) {
        lbkind = j
        lbk = exes[j]
        break
      }
      
      if (j == 1) {
        lbkind = j
        lbk = exes[j]
      }
      
    }
  }
  
  
  # calculate bankful width (in feet)
  width = rbk-lbk
  
  # calculate river cross sectional area (in square feet) at bankfull (simple method of leftbound rectangles)
  subwhys = whys[lbkind:rbkind]
  diffs = bank - subwhys[which(subwhys <= bank)]
  area = sum(diffs*int)
  
  # calculate the maximum channel depth
  dmax = bank-echanmin
  
  # calculate the average channel depth
  channeldepths = bank - whys[lbkind:rbkind]
  dmean = mean(channeldepths)
  
  # calculate W/D ratio
  wdratio = width/dmean

  ##### calculate the width of the floodprone area (IF RIFFLE)
  
  isRiff = FALSE
  if(type == 'r') {
    isRiff = TRUE
  }
  
  fpe = bank + dmax # flood prone elevation, equivalent to 2*dmax + min channel elevation
  
  #identify the right flood prone station
  rfpind = NULL
  rfp = NULL
  for (j in mindex:length(whys)) {
    
    if (abs(fpe-whys[j]) < tol) {
      rfpind = j
      rfp = exes[j]
      break
    }
    
    if (j == length(whys)) {
      rfpind = j
      rfp= exes[j]
    }
    
  }
  
  #identify the left flood prone station
  lfpind = NULL
  lfp = NULL
  for (j in mindex:1) {
    
    if (abs(fpe-whys[j]) < tol) {
      lfpind = j
      lfp = exes[j]
      break
    }
    
    if (j == 1) {
      lfpind = j
      lfp = exes[j]
    }
    
  }
  
  fpaw = rfp - lfp
  
  # find the entrenchment ratio
  enratio = fpaw/width
  
  if(!isRiff) {
    
    fpaw = NA
    enratio = NA
    
  }
  # find the bank height ratio
  bhratio = 1 # theoretically should always be one since we take the low bank height to be the given "bankfull"
  
  if (!is.na(lowBankHeight)) {
    numer = dmax+(lowBankHeight-bkf)
    denom = dmax
    bhratio = numer/denom
  }
  
  # variable names:
  # bkf width = width
  # fpa width = fpaw
  # bkf xs area = area
  # bkf mean depth = dmean
  # bkf max depth = dmax
  # W/D ratio = wdratio
  # entrenchment ratio = enratio
  # bank height ratio = bhratio
  

  
  statvec = c(width,fpaw,area,dmean,dmax,wdratio,enratio,bhratio,reachdata[currentsheet])
  istats[,currentsheet] = statvec
  names(istats)[currentsheet] = xsname


}

write.xlsx(x = istats, file = outname, sheetName = 'Individual XS Stats', row.names = TRUE, showNA = FALSE, append = FALSE)



##### now let's make the aggregate stats - min mean med max SD n


working = as.matrix(istats)
riffles = which(typevec == 'r' | typevec == 'R')

working = working[,riffles]
reachdata = reachdata[riffles]

stattype = c('Min','Mean','Med','Max','SD','n')
blankagdata = matrix(NA,nrow(working)-1,length(stattype)) # blank matrix to be filled
colnames(blankagdata) = stattype
rownames(blankagdata) = statnames[1:(length(statnames)-1)]

uniqueid = unique(reachdata)
lengther = length(uniqueid) # number of reaches

books = NULL # will contain statistical data for each reach


for (i in 1:lengther) {
  
  books[[i]] = blankagdata
  
  notthisreach = which(reachdata != uniqueid[i])
  if (length(notthisreach) != 0) {
    thisdata = working[,-notthisreach]
  }
  else { # if there is just 1 reach
    thisdata = working
  }
  
  for (j in 1:nrow(blankagdata)) {
    
    if (is.matrix(thisdata)) {
      notna = which(!is.na(thisdata[j,]))
      
      min = min(thisdata[j,notna])
      mean = mean(thisdata[j,notna])
      med = median(thisdata[j,notna])
      max = max(thisdata[j,notna])
      dev = sd(thisdata[j,notna])
      n = length(notna)
      
      if (n <= 2) {
        med = NA
      }
    } else {
      
      number = thisdata[j]
      min = number
      mean = number
      med = NA
      max = number
      dev = 0
      if (is.na(number)) {
        n = 0
      } else{
        n = 1
      }
      
    }
    
    #if (n <= 3) { # remove SD for statistics with less than three observations
    #  dev = NA
    #}
    
    stats = c(min,mean,med,max,dev,n)
    bstats = which(stats == Inf | stats == -Inf) # remove statistics equal to +/- Inf
    stats[bstats] = NA
    
    
    
    books[[i]][j,] = stats
    
  }
  
}

names(books) = paste('Reach', uniqueid)

for (i in 1:length(books)) {

  write.xlsx(x = books[[i]], file = outname, sheetName = names(books)[i], row.names = TRUE, showNA = FALSE, append = TRUE)

}