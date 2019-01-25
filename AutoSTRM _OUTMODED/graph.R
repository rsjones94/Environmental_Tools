rm(list=ls())

# This series of scripts was last checked to work with R v3.3.0, 9/27/16

########### USER INPUT

mywd = "X:\\2017\\161702811 TSMP Monitoring\\161702811B_Chilogatee_MY3\\Field\\Pebbles" # filepath to folder where your data is

# set these to true to make XS, profile and particle graphs respectively
workXS = F
workPro = F
workPart = T



xssource ='FT_pillow_database_xs' # where is the XS data coming from? (.xlsx files only)

prosource ='FT_pillow_database_pro' # where is the profile data coming from? (.xlsx files only)

partsource ='chil_pebble_database' # where is the particle data coming from? (.xlsx files only)

########### END USER INPUT





















############## call the graphing scripts


script.dir <- dirname(sys.frame(1)$ofile) # the directory that this script is in (only works when run as source)

if (workXS) {
  
  xsgrapherloc = normalizePath(paste(script.dir,'/Hidden/xsgrapher.R',sep=''))
  source(xsgrapherloc)
  
}

if (workPro) {
  
  prographerloc = normalizePath(paste(script.dir,'/Hidden/profilegrapher.R',sep=''))
  source(prographerloc)
  
}

if (workPart) {
  
  partgrapherloc = normalizePath(paste(script.dir,'/Hidden/pebblegrapher.R',sep=''))
  source(partgrapherloc)
  
}


