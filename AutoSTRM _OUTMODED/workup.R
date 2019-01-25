rm(list=ls())

# This series of scripts was last checked to work with R v3.3.0, 9/27/16

########### USER INPUT


mywd = 'X:\\2017\\161702811 TSMP Monitoring\\161702811E_NFMC_MY4\\Field\\Survey\\2017' # filepath to folder where your data is



workXS = T # set these to true to make XS, profile and particle graphs respectively
workPro = T
workPart = F


xssource ='NFMC_Stitch_Clean_xs' # where is the XS data coming from? (.xlsx files only)
xsout = 'NFMC_Stitch_Clean_xs_stats' # output filename for XSs

prosource ='NFMC_Stitch_Clean_pro' # where is the profile data coming from? (.xlsx files only)
proout = 'NFMC_Stitch_Clean_pro_stats' # output filename for profiles

partsource ='raw_pebbles' # where is the particle data coming from? (.xlsx files only)
partout = 'workpebble' # output filename for particles

########### END USER INPUT




















############## call the workup scripts



script.dir <- dirname(sys.frame(1)$ofile)

if (workXS) {
  
  xsworkuploc = paste(script.dir,'/Hidden/xsworkup.R',sep='')
  source(xsworkuploc)
  
}

if (workPro) {
  
  proworkuploc = paste(script.dir,'/Hidden/proworkup.R',sep='')
  source(proworkuploc)
  
}

if (workPart) {
  
  partworkuploc = paste(script.dir,'/Hidden/pebbleworkup.R',sep='')
  source(partworkuploc)
  
}