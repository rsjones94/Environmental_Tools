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

#mywd = 'C:/Users/Randall.Jones/Desktop/xs_mock' # where your file is located
#xssource = 'ladderxs' # name of the file with XS data

########### END USER INPUT
























### check out the input data

plot = TRUE # true to save graphs, false to not

#currentsheet = 1

xssource = paste(xssource,'.xlsx',sep='')

mywd = normalizePath(mywd)
setwd(mywd)
sheets = getSheets(loadWorkbook(xssource))
numsheets = length(sheets)


for (currentsheet in 1:numsheets) {
  
  
  data = read.xlsx(xssource, sheetIndex = currentsheet, header = TRUE, stringsAsFactors=FALSE)
  
  bkf = as.numeric(data[1,2])
  type = data[3,2]
  membership = toString(data[4,2])
  #membership = as.numeric(data[4,2])
  title = data[2,2]
  
  basin = data[5,2]
  watershed = data[6,2]
  drainageArea = data[7,2]
  prepDate = data[10,2]
  crew = data[11,2]
  lowBankHeight = as.numeric(data[12,2])
  chartTitle = data[13,2]
  
  infoString = c('River Basin', 'Watershed', 'XS ID', 'Drainage Area', 'Date', 'Field Crew')
  infoDat = c(basin, watershed, title, drainageArea, prepDate, crew)
  # if the user specifies DELETE for an info field, delete it from the grob table
  toDELETE = which(infoDat == 'DELETE')
  
  if(!length(toDELETE) == 0) { # if there is something to delete
    infoString = infoString[-toDELETE]
    infoDat = infoDat[-toDELETE]
  }
  
  infoDF = as.data.frame(infoDat)
  colnames(infoDF) = NULL
  rownames(infoDF) = infoString
  
  
  whereImages = data[9,2]
  code = data[8,2]
  
  if (plot){
    savename = paste(watershed,' ',title,'.pdf',sep='')
    pdf(file = savename, width = 11, height = 8.5) # 12.7 x 7.7
  }
  
  cutData = data[,-(1:2)]
  numYears = ncol(cutData)/4
  colSeq = (0:(numYears-1))*4
  
  monYears = as.matrix(cutData[1,colSeq+1])
  elevations = as.matrix(cutData[,colSeq+2])
  descriptions = as.matrix(cutData[,colSeq+3])
  stations = as.matrix(cutData[,colSeq+4])
  
  # 
  # # assemble a dataframe to feed into a qplot
  # 
  # yearRep = rep(monYears, nrow(stations))
  # yearMat = t(matrix(yearRep, nrow = numYears))
  # yearVec = matrix(yearMat, ncol = 1)
  # 
  # stationVec = matrix(stations, ncol = 1)
  # 
  # descriptionVec = matrix(descriptions, ncol = 1)
  # 
  # elevationVec = matrix(elevations, ncol = 1)
  # 
  # df = data.frame(stationVec, elevationVec, descriptionVec, yearVec)
  # df = na.omit(df)
  # 
  # # plot the data
  # 
  # qplot(stationVec, elevationVec, data = df, colour = yearVec, geom = 'line')
  # 
  
  lowest = min(elevations, na.rm = TRUE)
  highest = max(elevations, na.rm = TRUE)
  leftest = min(stations, na.rm = TRUE)
  rightest = max(stations, na.rm = TRUE)
  
  
  #layWidths = c(3,4.5,7)
  #layHeights = c(3,3,6)
  #layout(matrix(c(3,2,5,3,4,5,3,1,1), nrow = 3, ncol = 3, byrow = TRUE), widths = layWidths, heights = layHeights)
  
  # layWidths = c(3,3.7,7)
  # layHeights = c(3,3,6)
  # layout(matrix(c(3,2,5,3,4,5,3,1,1), nrow = 3, ncol = 3, byrow = TRUE), widths = layWidths, heights = layHeights)
  
  #layWidths = c(3,3.7,6.5,0.5)
  #layHeights = c(3,3,5.5,0.5)
  #layout(matrix(c(3,2,5,5,3,4,5,5,3,1,1,0,3,0,0,0), nrow = 4, ncol = 4, byrow = TRUE), widths = layWidths, heights = layHeights)
  
  layWidths = c(3,3.7,6.84,0.16)
  layHeights = c(3,3,4.5,0.5)
  layout(matrix(c(3,2,5,5,3,4,5,5,3,1,1,0,3,0,0,0), nrow = 4, ncol = 4, byrow = TRUE), widths = layWidths, heights = layHeights)
  
  if (type == 'r'){
    bank = bkf
    echanmin = min(na.omit(elevations[,ncol(elevations)])) # elevation of channel minimum
    dmax = bank-echanmin
    fpe = bank + dmax # flood prone elevation, equivalent to 2*dmax + min channel elevation
    highest = max(fpe,highest)
  } else {
    fpe = NA
  }
  
  plottitle = chartTitle
  plot(NULL, xlim = c(leftest,rightest), ylim = c(lowest,highest),
       xlab = '', ylab = '', main = plottitle,
       yaxp  = c(floor(lowest), ceiling(highest), ceiling(highest)-floor(lowest)))
  title(xlab="Station (ft)", mgp=c(2.25,1,0))
  title(ylab="Elevation (ft)", mgp=c(2.25,1,0))
  
  grid(nx = NA, ny = NULL, col = "lightgray", lty = 1,
       lwd = 1, equilogs = TRUE)
  
  bkfCol = 'dodgerblue4'
  abline(h = bkf, col = bkfCol, lty = 2, lwd = 2)
  
  fpaCol = 'red1'
  abline(h = fpe, col = fpaCol, lty = 2, lwd = 2)
  
  saveWidths = NULL
  saveCols= NULL
  
  for (i in 1:numYears) {
    
    stationing = na.omit(stations[,i])
    elevationing = na.omit(elevations[,i])
    
    int = 0.01
    bin = 0 # 0 turns smoothing off
    interp = approx(stationing, elevationing, xout=seq(from = min(stationing), to = max(stationing),by=int))
    model = ksmooth(interp$x, interp$y, bandwidth=bin, x.points=interp$x)
    
    whys = model$y
    exes = model$x
    
    if (i == numYears) {
      thick = 3
      #points(stationing, elevationing, col = i)
    } else {
      thick = 0.75
    }
    
    #thiscol = i+1
    thiscol = i # NOTE : make a color vector later
    lines(exes, whys, col = thiscol, lwd = thick)
    
    saveWidths[i] = thick
    saveCols[i] = thiscol
    
  }
  
  ### legend
  if (type == 'r') {
    legText = c(monYears, 'Bankfull','Flood Prone Area')
    legSyms = c(seq(from = 1, to = 1, length = numYears), 2, 2)
    legWidths = c(saveWidths, 2, 2)
    legCols = c(saveCols,bkfCol,fpaCol)
    
  } else {
    legText = c(monYears, 'Bankfull')
    legSyms = c(seq(from = 1, to = 1, length = numYears), 2)
    legWidths = c(saveWidths, 2)
    legCols = c(saveCols,bkfCol)
  }
  
  legendX = rightest-(rightest-leftest)*0.25
  legendY = lowest+abs(highest-lowest)*0.25
  legend(legendX, legendY, legText, lty = legSyms, lwd = legWidths, col = legCols, bg = 'white')
  
  ######## XS stats
  
  bkfExceeded = FALSE
  fpaExceeded = FALSE
  
  desc = na.omit(descriptions[,i])
  
  wherelb = grep('ltob', desc)
  whererb = grep('rtob', desc)
  
  bank = bkf
  
  tol = int*10
  
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
        bkfExceeded = TRUE
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
        bkfExceeded = TRUE
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
  
  fpe = bkf+dmax
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
      fpaExceeded = TRUE
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
      fpaExceeded = TRUE
    }
    
  }
  
  fpaw = rfp - lfp
  
  
  # find the entrenchment ratio
  enratio = fpaw/width
  
  if(!isRiff) {
    
    fpaw = NA
    enratio = NA
    fpe = NA
    wdratio = NA
    
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
  
  show_bankheight = TRUE
  show_bankarea = TRUE
  show_bankwidth = TRUE
  show_floodplainelevation = TRUE
  show_floodplainwidth = TRUE
  show_maxdepth = TRUE
  show_meandepth = TRUE
  show_wdratio = TRUE
  show_entratio = TRUE
  show_bhratio = TRUE
  
  showVec = c(show_bankheight,
              show_bankarea,
              show_bankwidth,
              show_floodplainelevation,
              show_floodplainwidth,
              show_maxdepth,
              show_meandepth,
              show_wdratio,
              show_entratio,
              show_bhratio)
  
  showLogical = which(showVec == FALSE)
  
  
  statNames = c('Bankfull Elevation (ft)', 'Bankfull XS Area (sq ft)', 'Bankfull Width (ft)', 'Flood Prone Area Elevation (ft)', 'Flood Prone Width (ft)', 'Max Depth at Bankfull (ft)', 'Mean Depth at Bankfull (ft)', 'W / D Ratio', 'Entrenchment Ratio', 'Bank Height Ratio')
  statVec = round(c(bank,area,width,fpe,fpaw,dmax,dmean,wdratio,enratio,bhratio),1)
  statVec[5] = round(statVec[5]) # floodprone width only needs to go to the ones places
  statVec[which(is.na(statVec))] = '-'
  if (isRiff & fpaExceeded) {
    # if the floodprone height exceeds bank heights 
    statVec[5] = paste('>', statVec[5]) # we only know the minimum floodprone width

    statVec[9] = paste('>', statVec[9]) # we only know the minimum entrenchment ratio
    
  }
  
  # add trailing zero
  for (i in 1:length(statVec)) {
    alphy = toString(statVec[i])
    if (statVec[i] != '-' & !grepl('\\.', alphy) & i != 5) {
      statVec[i] = paste(alphy,'.0',sep='')
    }
  }
  
  if (isRiff & fpaExceeded) {
    # if the floodprone height exceeds bank heights
    statVec[5] = paste(statVec[5], '  ') # we only know the minimum floodprone width

    statVec[9] = paste(statVec[9], '  ') # we only know the minimum entrenchment ratio
    # just adding two spaces on the end to keep the actual number centered in the table
    # we have to do this after we add trailing zeroes (if applicable), otherwise the zeroes would be added after the spaces
  }
  
  if(!length(showLogical) == 0) { # if there's anything to remove
    statNames = statNames[-showLogical]
    statVec = statVec[-showLogical]
  }
  
  statDF = as.data.frame(statVec)
  
  
  
  
  
  
  ### making data frames and grobs
  
  trailStation = round(stationing,1)
  trailElevation = round(elevationing,2)
  
  for (i in 1:length(trailStation)) {
    staAlph = toString(trailStation[i])
    eleAlph = toString(trailElevation[i])
    
    if (!grepl('\\.[0-9]', staAlph)) {
      trailStation[i] = paste(staAlph,'.0', sep='')
    }
    
    if (!grepl('\\.[0-9][0-9]', eleAlph)) {
      
      if (grepl('\\.[0-9]', eleAlph)) {
        trailElevation[i] = paste(eleAlph,'0', sep='')
      } else {
        trailElevation[i] = paste(eleAlph,'.00', sep='')
      }
      
    }
    
  }
  
  shotDF = as.data.frame(cbind(trailStation,trailElevation))
  rownames(shotDF) = NULL
  colnames(shotDF) = c('Station (ft)', 'Elevation (ft)')
  
  bigtheme <- gridExtra::ttheme_default(
    core = list(fg_params=list(cex = .65)),
    colhead = list(fg_params=list(cex = .65, fontface = 4L)),
    rowhead = list(fg_params=list(cex = .6, fontface= 1L)))
  
  littletheme <- gridExtra::ttheme_default(
    core = list(fg_params=list(cex = .55)),
    colhead = list(fg_params=list(cex = .55)),
    rowhead = list(fg_params=list(cex = .55)))
  
  # specialtheme <- gridExtra::ttheme_default(
  #   core = list(fg_params=list(cex = .6, fontface = 'bold')),
  #   colhead = list(fg_params=list(cex = .6)),
  #   rowhead = list(fg_params=list(cex = .6, fontface=1L)))
  
  specialtheme <- gridExtra::ttheme_default(
    core = list(fg_params=list(cex = .65)),
    colhead = list(fg_params=list(cex = .65)),
    rowhead = list(fg_params=list(cex = .65, fontface=1L)))
  
  # grobs
  infoGrob = tableGrob(infoDF, theme = specialtheme) 
  shotGrob = tableGrob(shotDF, theme = littletheme, rows = NULL)
  dataGrob = tableGrob(statDF, theme = bigtheme, rows = statNames, cols = c('Summary Data'))
  
  # infoGrob$widths = unit(rep(.85, ncol(infoGrob)), 'npc')
  # infoGrob$heights = unit(rep(2.5/nrow(infoGrob), nrow(infoGrob)), 'npc')
  # 
  # shotGrob$widths = unit(rep(.4, ncol(shotGrob)), 'npc')
  # shotGrob$heights = unit(rep(.95/nrow(shotGrob), nrow(shotGrob)), 'npc')
  # 
  # dataGrob$widths = unit(rep(.4, ncol(dataGrob)), 'npc')
  # dataGrob$heights = unit(rep(1/nrow(dataGrob), nrow(dataGrob)), 'npc')
  
  infoGrob$widths = unit(rep(.8, ncol(infoGrob)), 'npc') # for XL, .9. For 8.5x11, .8
  infoGrob$heights = unit(rep(1.2/nrow(infoGrob), nrow(infoGrob)), 'npc')
  
  shotGrob$widths = unit(rep(.5, ncol(shotGrob)), 'npc') # .35 for XL, else .4
  shotGrob$heights = unit(rep(1/nrow(shotGrob), nrow(shotGrob)), 'npc') #1/nrow(shotGrob)
  
  dataGrob$widths = unit(rep(.7, ncol(dataGrob)), 'npc')# for XL, .6, else .65
  dataGrob$heights = unit(rep(2.1/nrow(dataGrob), nrow(dataGrob)), 'npc')
  
  
  
  # second base plot 
  frame()
  # Grid regions of current base plot (ie from frame)
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)
  
  grid.draw(infoGrob)
  popViewport(3)
  
  
  # third
  frame()
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)  
  grid.draw(shotGrob)
  popViewport(3)
  
  # fourth
  frame()
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)  
  grid.draw(dataGrob)
  popViewport(3)
  
  # fifth (image)
  #plot(1:10)
  
  #whereImages = "C:/Users/Randall.Jones/Desktop/xs_mock/Photos"
  tempwd = whereImages
  #setwd(tempwd)
  
  # function that takes a 3 channel image and adds a black border n layers thick
  addBorder = function(image, n) {
    
    workpic = image
    
    dim2 = dim(workpic)[2]
    
    addon = rep(0,n)
    addvert = matrix(rep(addon,dim2),ncol = dim2)
    
    r = rbind(addvert, workpic[,,1], addvert)
    g = rbind(addvert, workpic[,,2], addvert)
    b = rbind(addvert, workpic[,,3], addvert)
  
    dim1 = dim(r)[1]
    
    addhori = t(matrix(rep(addon,dim1),ncol = dim1))
    
    r = cbind(addhori, r, addhori)
    g = cbind(addhori, g, addhori)
    b = cbind(addhori, b, addhori)
    
    newimage = array(rep(0, dim(r)[1]*dim(r)[2]*3), dim=c(dim(r)[1], dim(r)[2], 3))
    newimage[,,1] = r
    newimage[,,2] = g
    newimage[,,3] = b
    
    return(newimage)
    
  }
  
  #plot(NULL,xlim=c(0,1), ylim=c(0,1))
  #rasterImage(newimage,0,0,1,1)
  
  potentials = list.files(tempwd)
  myregex = paste(code,'[^0-9]',sep='')
  finder = grep(myregex,potentials)
  
  img_name = potentials[finder]
  img_name = paste(tempwd,'\\',img_name,sep='')
  the_pic = readJPEG(img_name, native = FALSE)
  
  borderthickness = round(max(dim(the_pic))*.005)
  
  the_pic = addBorder(the_pic, n = borderthickness)
  rasta = grid::rasterGrob(the_pic) # height = 2
  
  frame()
  vps <- baseViewports()
  pushViewport(vps$inner, vps$figure, vps$plot)  
  grid.draw(rasta)
  popViewport(3)
  
  setwd(mywd)
  
  if (plot) {
    dev.off()
  }
  
}
