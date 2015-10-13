# Setup
install.packages(c("RDCOMClient","XML","rvest"))  
library(RDCOMClient)   
library(XML)   
library(rvest)

# Start up PowerPoint 
pp <- COMCreate("Powerpoint.Application")

# Make the application visible
pp[["Visible"]] = 1

# Add a new presentation
presentation <- pp[["Presentations"]]$Add()

# The presentation is empty.  Add a slide to it. We'll get to the enumerated constants 
# used by Microsoft. For now, we'll use the value that results in a blank slide.
# slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
slide1 <- presentation[["Slides"]]$Add(1,12)

# Add shapes and apply animation
# Source the enumerated constants used by Microsoft for various parameters
# The file is available on github.  Save it to the working directory 
# https://github.com/asifsalam/PowerPoint_from_R/blob/master/mso.txt
source("mso.txt")
shp1 <- slide1[["Shapes"]]$AddShape(ms$msoShape12pointStar,20,20,100,100)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectFadedSwivel,
                                                 trigger=ms$msoAnimTriggerAfterPrevious)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectPathBounceRight,
                                                 trigger=ms$msoAnimTriggerAfterPrevious)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectSpin,
                                                 trigger=ms$msoAnimTriggerAfterPrevious)

shp1$PickupAnimation()
shp2 <- slide1[["Shapes"]]$AddShape(ms$msoShapeHexagon,100,20,100,100)
shp2$ApplyAnimation()

shp3 <- slide1[["Shapes"]]$AddShape(ms$msoShapeCloud,180,20,100,200)
shp3$ApplyAnimation()

# Add text to the shapes.  While this works, R files a complaint
shp1[["TextFrame"]][["TextRange"]][["Text"]] <- "Shp1"

# This way seems to function better
shp1_tr <- shp1[["TextFrame"]][["TextRange"]]
shp1_tr[["Text"]] <- "ONE"


# Set some shape attributes.  
# The `Fill` property is used for the colors, and the `Line`Line property for the border.

shp1_color <- shp1[["Fill"]]
shp1_color[["ForeColor"]][["RGB"]] <- (0+170*256+170*256^2)
# That's how the RGB value is calculated: r +  g*256 + b*256*256 

# Remove the line
shp1_line <- shp1[["Line"]]
shp1_line[["Visible"]] <- 0

# Create function for the rgb calculation
pp_rgb <- function(r,g,b) {
    return(r + g*256 + b*256^2)
}

shp2_tr <- shp2[["TextFrame"]][["TextRange"]]
shp2_tr[["Text"]] <- "TWO"
shp2_color <- shp2[["Fill"]]
shp2_color[["ForeColor"]][["RGB"]] <- pp_rgb(170,170,0)
shp2_line <- shp2[["Line"]]
shp2_line[["Visible"]] <- 0

shp3_tr <- shp3[["TextFrame"]][["TextRange"]]
shp3_tr[["Text"]] <- "THREE"
shp3_color <- shp3[["Fill"]]
shp3_color[["ForeColor"]][["RGB"]] <- pp_rgb(170,0,170)
shp3_line <- shp3[["Line"]]
shp3_line[["Visible"]] <- 0

# Finally, save the file in the working directory
presentation$SaveAs(paste0(getwd(),"/PowerPoint_R_Part_1.pptx"))
