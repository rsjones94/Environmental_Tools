#rm(list=ls())



# if you don't have a local copy of package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)


########### USER INPUT

# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes

#mywd = 'M:/2014/16146867_NGP Monitoring/272-Ellerbe NGP - MY05/Stream' # where your file is located
#prosource = 'cleandatpro' # name of the file with the data (MUST end in .xlsx)
#proout = ''


########### END USER INPUT












### check out the input data
mywd = normalizePath(mywd)
setwd(mywd)
prosource = paste(prosource,'.xlsx',sep='')
sheets = getSheets(loadWorkbook(prosource))
numsheets = length(sheets)

outname = paste(proout,'.xlsx', sep = '')

rownames = c('SC/SA/G/C/B/BE', 'd16/d35/d50/d84/d95', 'ri/ru/po/gl/str', 'Channel Length', 'Drainage Area', 'Rosgen', 'Sinuosity', 'Water Surface Slope', 'Bankfull Slope', '% Eroding Banks', 'Reach Number')
idata = as.data.frame(matrix(NA,length(rownames),length(numsheets)))
rownames(idata) = rownames

########## INDIVIDUAL DATA

for (currentsheet in 1:numsheets) {
  
  data = read.xlsx(prosource, sheetIndex = currentsheet, header = TRUE)
  
  
  sta = data[,10] # station
  ele = data[,6] # thalweg elevation
  name = data[,5] # description
  ws = data[,7] # water surface
  bkf = data[,8] # bankfull
  proname = toString(data[1,13]) # assign name of the cross section
  whichreach = data[1,14]
  
  
  waters = which(!is.na(ws)) # indices of ws shots
  banks = which(!is.na(bkf)) # indices of bkf shots
  
  ## let's get the slope of the water surface
  
  if (length(waters) >= 2) {
    
    firstwater = ws[waters[1]]
    lastwater = ws[waters[length(waters)]]
    distance = sta[waters[length(waters)]] - sta[waters[1]]
    
    waterslope = (firstwater-lastwater)/distance # in ft/ft
    
    if (waterslope < 0) {
      waterslope = 0
    }
    
  } else {
    
    waterslope = NA
    
  }
  
  ## let's get the bankfull slope
  
  if (length(banks) >= 2) {
    
    firstbank = bkf[banks[1]]
    lastbank = bkf[banks[length(banks)]]
    distance = sta[banks[length(banks)]] - sta[banks[1]]
    
    bankslope = (firstbank-lastbank)/distance # in ft/ft
    
  } else {
    
    bankslope = NA
    
  }
  
  #### morphology breakdown
  
  sta = data[,10] # station
  ele = data[,6] # thalweg elevation
  morph = data[,12] # description
  proname = toString(data[1,13]) # assign name of the cross section
  reach = data[1,14]
  
  totaldist = sta[length(na.omit(sta))]
  
  bri = grep('bri', morph, value = FALSE, ignore.case = TRUE)
  eri = grep('eri', morph, value = FALSE, ignore.case = TRUE)
  ri = rbind(bri,eri)
  
  bru = grep('bru', morph, value = FALSE, ignore.case = TRUE)
  eru = grep('eru', morph, value = FALSE, ignore.case = TRUE)
  ru = rbind(bru,eru)
  
  bpo = grep('bpo', morph, value = FALSE, ignore.case = TRUE)
  epo = grep('epo', morph, value = FALSE, ignore.case = TRUE)
  po = rbind(bpo,epo)
  
  bgl = grep('bgl', morph, value = FALSE, ignore.case = TRUE)
  egl = grep('egl', morph, value = FALSE, ignore.case = TRUE)
  gl = rbind(bgl,egl)
  
  bst = grep('bst', morph, value = FALSE, ignore.case = TRUE)
  est = grep('est', morph, value = FALSE, ignore.case = TRUE)
  st = rbind(bst,est)
  
  inclusions = grep('IN', morph, value = FALSE, ignore.case = TRUE)
  beginin = inclusions[c(TRUE,FALSE)]
  endin = inclusions[c(FALSE,TRUE)]
  incl = rbind(beginin,endin)
  
  morphlist = NULL
  morphlist[[1]] = ri
  morphlist[[2]] = ru
  morphlist[[3]] = po
  morphlist[[4]] = gl
  morphlist[[5]] = st
  names(morphlist) = c('Riffles', 'Runs', 'Pools', 'Glides', 'Structures')
  
  
  morphdist = NULL
  for (i in 1:length(morphlist)) {
    
    distanceind = matrix(morphlist[[i]], nrow = 1)
    distances = sta[distanceind]
    distancemat = matrix(distances, nrow=2, byrow = FALSE)
    
    diffs = diff(distancemat)
    morphdist[i] = sum(diffs)
    
  }
  
  inclind = matrix(incl, nrow = 1)
  incldistances = sta[inclind]
  inclmat = matrix(incldistances, nrow=2, byrow = FALSE)
  incldiffs = diff(inclmat)
  
  mrnm = c('riffle', 'run', 'pool', 'glide', 'structure')
  whichone = NULL # can be used after the triple loop to confirm the inclusions were removed from the correct morphs
  
  if (!is.na(incl[1]) & !is.na(incl[2])) {
    for (i in 1:length(incldiffs)) { # for each reach identified as an inclusion
      
      for (j in 1:(length(morphlist)-1)) { # check against each list of reach pairs, excluding structures
        
        currentmorph = morphlist[[j]]
      
        for (k in 1:ncol(currentmorph)) { # to see if the inclusion is within that morphology
          
          if (incl[1,i] >= currentmorph[1,k] & incl[2,i] <= currentmorph[2,k]) { # if it is
            
            morphdist[j] = morphdist[j] - incldiffs[i] # subtract the inclusion length from that morphology
            whichone[i] = mrnm[j]
            
          }
          
        } 
        
      }
      
    }
  }
  # note that the above inclusion algorithm does allow for coterminous beginnings/endings, but does not allow for anything
  # but structures to be included
  
  
  
  
  breakdown = (morphdist/totaldist)*100
  rbd = round(breakdown, 1)
  
  bdown = paste(rbd[1], rbd[2], rbd[3], rbd[4], rbd[5], sep = ' / ')
  
  
  
  
  sedtype = NA
  dX = NA
  channellength = NA
  drainage = NA
  rosgen = NA
  sinu = NA
  # ws slope
  # bkf slope
  percerode = NA
  
  idata[,currentsheet] = c(sedtype,dX,bdown,channellength,drainage,rosgen,sinu,waterslope,bankslope,percerode,whichreach)
  names(idata)[currentsheet] = proname


}

write.xlsx(x = idata, file = outname, sheetName = 'Individual Stats', row.names = TRUE, showNA = FALSE, append = FALSE)











#### aggregate data

agdata = NULL
reachvec = NULL
for (currentsheet in 1:numsheets) {
  
  data = read.xlsx(prosource, sheetIndex = currentsheet, header = TRUE)
  
  morph = data[,12] # description
  proname = data[1,13]
  reach = data[1,14]
  reachvec[currentsheet] = reach
  
  
  bri = grep('bri', morph, value = FALSE, ignore.case = TRUE)
  eri = grep('eri', morph, value = FALSE, ignore.case = TRUE)
  
  bpo = grep('bpo', morph, value = FALSE, ignore.case = TRUE)
  epo = grep('epo', morph, value = FALSE, ignore.case = TRUE)
  
  ele = data[,6]
  sta = data[,10]
  ws = data[,7] # water surface
  
  filled = which(!is.na(ws)) # indices of ws shots
  
  # linear approximation of the water surface
  ws = approx(sta[filled], ws[filled], xout = sta, rule = 1)$y
  
  
  riffdat = matrix(NA,length(bri),2) # row will be each riffle, c1 is length and c2 is slope
  for (i in 1:length(bri)) {
    
    rlength = sta[eri[i]] - sta[bri[i]]
    slope = (ws[bri[i]] - ws[eri[i]]) / rlength
    
    if (slope < 0) { # water can't flow upstream... usually
      
      slope = 0
      
    }
    
    riffdat[i,] = c(rlength,slope)
    
  }
  
  ## now find each pool and calculate its length, max depth and spacing
  
  pooldat = matrix(NA,length(bpo),2) # row will be each pool, c1 is length and c2 is max depth
  for (i in 1:length(bpo)) {
    
    length = sta[epo[i]] - sta[bpo[i]]
    
    ## find closest water surface shot forward of the bpo
    fdist = 0
    while(TRUE) {
      
      
      if (!is.na(ws[bpo[i]+fdist])) { # if the ws data is not NA
        
        fel = ws[bpo[i]+fdist]
        break
        
      } else if (bpo[i]+fdist == length(ws)) { # if we reach the end of the data and haven't found anything
        
        fdist = NA
        break
        
      } else { # we haven't reached the end but still haven't found anything
        
        fdist = fdist+1
        
      }
      
    }
    
    
    # find the deepest point in the pool
    mdepth = 0
    
    thalvec = bpo[i]:epo[i]
    depths = ws[thalvec] - ele[thalvec]
    mdepth = max(depths)
    
    
    pooldat[i,] = c(length,mdepth)
    
  }
  
  ## now find the pool spacing and add it to pooldat (will have length 1 short than the number of pools)
  
  plength = NULL
  for (i in 2:length(bpo)) {
    
    plength[i-1] = sta[bpo[i]] - sta[bpo[i-1]]
    
  }
  
  if (length(plength) == nrow(pooldat)-1) {
    pooldat = cbind(pooldat, c(plength,NA))
  } else {
    pooldat = cbind(pooldat, plength)
  }
  
  
  
  
  
  ## now calculate the statistics
  ## yes this can be done with a loop that's like 10 lines long but I wrote this a while ago
  # variable codes:
  # Ri = Riffle, Po = Pool
  # L = length, Sl = Slope, D = Max Depth, Sp = Spacing
  
  # riffle length
  RiLmin = min(riffdat[,1], na.rm = TRUE)
  RiLmean = mean(riffdat[,1], na.rm = TRUE)
  RiLmedian = median(riffdat[,1], na.rm = TRUE)
  RiLmax = max(riffdat[,1], na.rm = TRUE)
  RiLsd = sd(riffdat[,1], na.rm = TRUE)
  RiLn = length(riffdat[,1][!is.na(riffdat[,1])])
  
  RiL = c(RiLmin, RiLmean, RiLmedian, RiLmax, RiLsd, RiLn)
  
  # riffle slope
  RiSlmin = min(riffdat[,2], na.rm = TRUE)
  RiSlmean = mean(riffdat[,2], na.rm = TRUE)
  RiSlmedian = median(riffdat[,2], na.rm = TRUE)
  RiSlmax = max(riffdat[,2], na.rm = TRUE)
  RiSlsd = sd(riffdat[,2], na.rm = TRUE)
  RiSln = length(riffdat[,2][!is.na(riffdat[,2])])
  
  RiSl = c(RiSlmin, RiSlmean, RiSlmedian, RiSlmax, RiSlsd, RiSln)
  
  # pool length
  PoLmin = min(pooldat[,1], na.rm = TRUE)
  PoLmean = mean(pooldat[,1], na.rm = TRUE)
  PoLmedian = median(pooldat[,1], na.rm = TRUE)
  PoLmax = max(pooldat[,1], na.rm = TRUE)
  PoLsd = sd(pooldat[,1], na.rm = TRUE)
  PoLn = length(pooldat[,1][!is.na(pooldat[,1])])
  
  PoL = c(PoLmin, PoLmean, PoLmedian, PoLmax, PoLsd, PoLn)
  
  #pool max depth
  PoDmin = min(pooldat[,2], na.rm = TRUE)
  PoDmean = mean(pooldat[,2], na.rm = TRUE)
  PoDmedian = median(pooldat[,2], na.rm = TRUE)
  PoDmax = max(pooldat[,2], na.rm = TRUE)
  PoDsd = sd(pooldat[,2], na.rm = TRUE)
  PoDn = length(pooldat[,2][!is.na(pooldat[,2])])
  
  PoD = c(PoDmin, PoDmean, PoDmedian, PoDmax, PoDsd, PoDn)
  
  # pool spacing
  PoSpmin = min(pooldat[,3], na.rm = TRUE)
  PoSpmean = mean(pooldat[,3], na.rm = TRUE)
  PoSpmedian = median(pooldat[,3], na.rm = TRUE)
  PoSpmax = max(pooldat[,3], na.rm = TRUE)
  PoSpsd = sd(pooldat[,3], na.rm = TRUE)
  PoSpn = length(pooldat[,3][!is.na(pooldat[,3])])
  
  PoSp = c(PoSpmin, PoSpmean, PoSpmedian, PoSpmax, PoSpsd, PoSpn)
  
  
  
  ### make a matrix of the data
  
  finalstats = rbind(RiL, RiSl, PoL, PoD, PoSp)
  colnames(finalstats) = c('Min', 'Mean', 'Med', 'Max', 'SD', 'n')
  rownames(finalstats) = c('Riffle Length (ft)', 'Riffle Slope (ft/ft)', 'Pool Length (ft)', 'Pool Max Depth (ft)', 'Pool Spacing (ft)')
  
  lessthan2 = which(finalstats[,6] <= 2, arr.ind = TRUE)
  finalstats[lessthan2,3] = NA # removes medians when there are two or fewer observations
  
  lessthan3 = which(finalstats[,6] <= 3, arr.ind = TRUE)
  finalstats[lessthan3,5] = NA # removes SDs when there are three or fewer observations
  
  
  
  agdata[[currentsheet]] = finalstats
  
}

for (i in 1:length(agdata)) {
  
  write.xlsx(x = agdata[[i]], file = outname, sheetName = paste('Reach', reachvec[i]), row.names = TRUE, showNA = FALSE, append = TRUE)
  
}

