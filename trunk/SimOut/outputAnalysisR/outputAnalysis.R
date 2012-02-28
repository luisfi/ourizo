################################################################################
##
## Output Analysis Script 
## Takes outputs from model and analizes them. 
## Incoming outputs are assumed to have commas for decimal separators and 
## semicolons for field separators (as in spanish locale .csv).
## If your outputs have another locale change the configutarion of sep and
## dec arguments at code lines 56, 60 and 64.
##
## Created: 22-jan-2012
## Last modified: 29-jan-2012
## Version: 0.1
## Author/s: InsNay (inesnaya@gmail.com)  and AnaParma
##
################################################################################

# Set working directory (CHANGE IF REQUIRED)
setwd("/home/inesnaya/Laboratorio/Tesis/Modelos/ourizo/SimOut/outputAnalysisR")
pathToModelOutputFiles <- "modeloutputs/"
pathToOutputAnalysisFiles <- "analysisoutputs/"
## In pathToOutputAnalysisFiles: Objects are stored as CSV in forder csv.
pathToGrafs <- "analysisoutputs/graf/"

# Load libraries
library(ggplot2) #Graphics Library
library(grid)    # Grid Layouts: Many plot in one page.

################################################################################
# LOAD AUXILIARY FUNCTIONS #####################################################
################################################################################

source("auxiliaryFunctions.R")
# changePreValue((objectName, functionPart, subpart="", assignment)
# createPreObjects(objectName,  assignment)
# createLayout(n)
# vpLayout(x, y)

################################################################################
# LOAD DATA ARCHIVES # CREATE AUXILIARY DATASETS ###############################
################################################################################
# Reads each output file of name structure 
## S[0-9]+.M([0-9])+.P([0-9])+.A([0-9])+.U([0-9])+.(I[0-9]).csv and names it 
## with the corresponding file name for identification:

#First check filenames with conventional naming structure for the output files:
filenames <- dir(pathToModelOutputFiles, 
                 "S[0-9]+.M([0-9])+.P([0-9])+.A([0-9])+.U([0-9])+.(I[0-9])+.csv")
# Create list of object names for the outputs:
objectnames <- strsplit(filenames, ".csv")

#Prueba

  #Read main output
  assign(objectnames[[i]],
                       read.csv(paste(pathToModelOutputFiles,filenames[i],sep=''),
                       sep=';', dec=',', na.string='', header=T) )

# Read each output file and assign it to its name:
for (i in 1:length(filenames)) 
{
  #Read main output
   assign(objectnames[[i]],
                       read.csv(paste(pathToModelOutputFiles,filenames[i],sep=''),
                       sep=';', dec=',', na.string='', header=T))
  #Read totals 
  assign(paste(objectnames[[i]], ".totals", sep=""),             
        read.csv(paste(pathToModelOutputFiles,objectnames[i], '.totals.csv',sep=''),
                       sep=';', dec=',', na.string='', header=T)) 
  #Read totals by region
  assign(paste(objectnames[[i]], ".region", sep=""),             
        read.csv(paste(pathToModelOutputFiles,objectnames[i], '.region.csv',sep=''),
                       sep=';', dec=',', na.string='', header=T))
}

# Each model main output file has structure:
# > names(eval(as.name(objectnames[[1]])))
#  [1] "Monte"             "Region"            "Area"             
#  [4] "Year"              "HR"                "Catch"            
#  [7] "Effort"            "Bvulnerable"       "Bmature"          
# [10] "Btotal"            "Larvae"            "Settlers"         
# [13] "Recruits"          "Density"           "Depletion_Bvul"   
# [16] "Depletion_Bmature"    

################################################################################
# PRELIMINARY EXPLORATORY ANALISYS AND GRAPHS ##################################
################################################################################

# SUMMARY OUTPUTS OF DATA ######################################################
for (i in 1:length(filenames)) 
{
  # CODE FOR EACH OUTPUT FILE GOES HERE ########################################
  ## To get each element of the summaries: nameObject$nameVariable[,dimensions]
  ### Dimensions: [1]Min./[2]1st Qu./[3]Median/[4]Mean/[5]3rd Qu./[6] Max.
  tempname <- paste(objectnames[[i]],".meanvalues", sep="") 
  # Mean and dispersion values of all data
  assign(paste(tempname, ".total", sep=""), 
          aggregate(eval(as.name(paste(objectnames[[i]], ".totals", sep="")))[,-c(1:2)],
               by=list(dummy=rep(1, times=dim(eval(as.name(paste(objectnames[[i]], ".totals", sep=""))))[1])) 
            ,summary))
  ## To get each element nameObject$nameVariable[,dimensions]
  ### Dimensions: [1]Min./[2]1st Qu./[3]Median/[4]Mean/[5]3rd Qu./[6] Max.
 
  # Mean and dispersion values of all data by year
  assign(tempname,
          aggregate(eval(as.name(paste(objectnames[[i]], ".totals", sep="")))[,-(1:2)],
               by=list(Year=eval(as.name(paste(objectnames[[i]], ".totals", sep="")))$Year) 
            ,summary))
  # Mean and dispersion values by Nreplicates of the MONTECARLO simulations.
  assign(paste(tempname, ".monte", sep=""),
      aggregate(eval(as.name(paste(objectnames[[i]], ".totals", sep="")))[,-(1:2)],
               by=list(Monte=eval(as.name(paste(objectnames[[i]], ".totals", sep="")))$Monte) 
            ,summary))
  # Mean and dispersion values by Nreplicates of the MONTECARLO simulations AND REGION (If Region>1).
  if(length(unique(eval(as.name(objectnames[[i]]))$Region))>1)
  {
    assign(paste(tempname, ".region", sep=""),
        aggregate(eval(as.name(paste(objectnames[[i]], ".region", sep="")))[,-(1:3)],
               by=list(Monte=eval(as.name(paste(objectnames[[i]], ".region", sep="")))$Monte, 
                       Region=eval(as.name(paste(objectnames[[i]], ".region", sep="")))$Region)
                    ,summary))
  }
  # Mean and dispersion values by Nreplicates of the MONTECARLO simulations
  ## by AREA.
  assign(paste(tempname, ".area", sep=""),
      aggregate(eval(as.name(objectnames[[i]]))[,-(1:4)],
               by=list(Monte=eval(as.name(objectnames[[i]]))$Monte,
                       Area=eval(as.name(objectnames[[i]]))$Area) 
                  ,summary)))
  
  #Write mean and dispersion values to corresponding output files:
  for (j in ls(pattern=paste(tempname,".*", sep="")))
  {
    write.csv(eval(as.name(j)) 
              ,file=paste(pathToOutputAnalysisFiles,"csv/", j,".csv", sep=""))
  }
  # CODE FOR EACH OUTPUT FILE ENDS HERE ########################################
}

# GENERAL PLOTS ################################################################

## General output plots Indicator vs. Year 
### Totals output (all areas collapsed)
### General model outputs, for each montecarlo simulation and general tendencies
### Grey shaded area indicates 95% C.I. for the general tendency assuming
### normality.
pdf(paste(pathToGrafs, objectnames[[i]], ".totals.grafs.pdf", sep=""))
  for (k in names(eval(as.name(paste(objectnames[[i]], ".totals", sep=""))))[-c(1,2)])
  {
    print(qplot(Year, eval(as.name(k)), ylab=k ,data=eval(as.name(paste(objectnames[[i]], ".totals", sep=""))),
        group =Monte, geom="line",color= as.factor(Monte))+ theme_bw() + geom_smooth(aes(group=1), colour="grey30"))
     }
dev.off()

## Summary plots of all data
pdf(paste(pathToGrafs, objectnames[[i]], ".total.grafs.pdf", sep=""))
#  grid.newpage()
#  pushViewport(viewport(layout = grid.layout(2, 2)))

for (k in names(eval(as.name(paste(tempname, ".totals", sep=""))))[-c(1)])
{
  print(qplot(Year, eval(as.name(k))[,3], ylab=k ,data=eval(as.name(paste(tempname, ".totals", sep=""))),
        group =Monte, color= as.factor(Monte))+ theme_bw() + geom_smooth(aes(group=1), colour="grey30"))
}
dev.off()
 print(qplot(Year, eval(as.name(k)), ylab=k ,data=eval(as.name(paste(objectnames[[i]], ".totals", sep=""))),
        group =Monte, geom="line", color= as.factor(Monte))+ theme_bw() + geom_smooth(aes(group=1, y=eval(as.name(paste(tempname, "$", k,sep="")))[,3], ymin=eval(as.name(paste(tempname, "$", k,sep="")))[,2], ymax=eval(as.name(paste(tempname, "$", k,sep="")))[,5]), colour="grey30", , se=F))

abc <- adply(matrix(rnorm(100), ncol = 5), 2, quantile, c(0, .25, .5, .75, 1))
b <- ggplot(eval(as.name(paste(tempname, ".total", sep=""))), aes(x = names(eval(as.name(paste(tempname, ".totals", sep=""))))[-c(1)], ymin = `0%`, lower = `25%`, middle = `50%`, upper = `75%`, ymax = `100%`))

