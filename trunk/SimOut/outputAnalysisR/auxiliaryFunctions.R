################################################################################
# AUXILIARY FUNCTIONS LIBRARY###################################################
################################################################################

# FUNCTIONS FOR OBJECT CREATION AND MANIPULATION ###############################

# helps CREATES NEW OBJECTS: 
##To actually create the objectc you have to call the function within an eval()
## Both function arguments are strings.
createPreObject <- function(objectName,  assignment) 
   {
    return(parse(text=paste(objectName, "<-", assignment, sep="")))
    }
# helps TO CHANGE VALUES OR PARTS OF OBJETCS: 
##To actually change the VALUE you have to call the function within an eval()
## Both function arguments are strings.
changePreValue <- function(objectName, functionPart, subpart="", assignment) 
   {
    envelopingFuntion <- paste(as.character(functionPart),"(", objectName, ")",subpart,sep="")
    return(parse(text=paste(envelopingFuntion, "<-", assignment, sep="")))
    }

# GRID FUNCTIONS FOR PLOT LAYOUTS ##############################################
## Depend on package grid.

# Configures individual plot positions inside of the current viewport. 
vpLayout <- function(x, y)
  {
    viewport(layout.pos.row = x, layout.pos.col = y)
  }
# Push Viewport with layout dimensioned according to the number of elements to plot together
createLayout <- function(n)
  {  
    if (n == 1){
      nrow <- 1
      ncol <- 1
    } 
    if (n == 2){
      nrow <- 1
      ncol <- 2
    } 
    if (n==3){
      nrow <- 1
      ncol <- 3
    } 
    if (n==4){
      nrow <- 2
      ncol <- 2
    } 
    if (n==5 | n==6){
      nrow <- 2
      ncol <- 3
    } 
    if (n==7 & n==8){
      nrow <- 2
      ncol <- 4
    } 
    if (n==9){
      nrow <- 3
      ncol <- 3
    } 
    if (10>=n & n>=12){
      nrow <- 4
      ncol <- 3
    } 
    if (13>=n & n>=16){
      nrow <- 4
      ncol <- 4
    } 
    if (n<=17){
      nrow <- 5
      ncol <- 5
    }
      
    return (pushViewport(viewport(layout = grid.layout(nrow, ncol))))
  }
