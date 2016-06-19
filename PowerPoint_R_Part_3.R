library(RDCOMClient)
library(plyr)
library(dplyr)
library(stringr)
library(rvest)

# Assuming all the files are stored in the
# download and read in the data files

download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_films.csv",
              destfile = "eastwood_films.csv")
download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_box_office.csv",
              destfile = "box_office.csv")
films <- read.table("eastwood_films.tsv",header=TRUE,sep="\t", stringsAsFactors=FALSE)
box_office <- read.table("eastwood_box_office.tsv",header=TRUE,sep="\t", stringsAsFactors=FALSE)
source("mso.txt")
actor_name <- "Clint Eastwood"
img_dir <- "img"

# If you haven't downloaded the images, already -
# Loop through the films and download the poster image into the "img" subdirectory.
# If the poster is not found, flag the file name with 0.
if (!file.exists("img")) dir.create("img")
for (i in 1:nrow(films)) {
    img_node <- html(films$url[i]) %>%
        html_nodes(xpath='//td[@id="img_primary"]//img')
    if (length(img_node)==0) {
        films$img_file[i] <- "img/img00.jpg"
        cat(i," : img file NOT FOUND: ",films$img_file[i],"\n")
    }
    else {
        img_link <- html_attr(img_node,"src")
        cat(i," :",films$img_file[i]," : ",img_link,"\n")
        download.file(img_link,films$img_file[i],method="internal",mode="wb")
    }
}

# Check which of the files were not found and download them manually
films$title[which(films$img_file=="img/img00.jpg")]

# These images don't exist.  Download appropriate images manually, and rename
# Films, Star in the Dust, The First Traveling Sales Lady, Dumbo Pilot
films[55,"img_file"] <- "img/img55.jpg"
films[54,"img_file"] <- "img/img54.jpg"
films[52,"img_file"] <- "img/img52.jpg"

########## Creating the PowerPoint Slide ###################33
# Create the PowerPoint slide
pp <- COMCreate("Powerpoint.Application")
pp[["Visible"]] = 1
presentation <- pp[["Presentations"]]$Add()
#slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutTitleOnly)
slide_width <- presentation[["PageSetup"]]$SlideWidth()
slide_height <- presentation[["PageSetup"]]$SlideHeight()

# Function that sets the colors
pp_rgb <- function(r,g,b) {
    return( r + g*256 + b*256^2)
}

# Set some slide attributes
slide_color <- slide1[["ColorScheme"]]$Colors(ms$ppBackground)
slide_color[["RGB"]] <- pp_rgb(0,0,0)

slide_title <- slide1[["Shapes"]][["Title"]]
slide_title_color <- slide1[["ColorScheme"]]$Colors(ms$ppTitle)
slide_title_color[["RGB"]] <- pp_rgb(200,200,200)

# Add a title
# AutoSize: https://msdn.microsoft.com/EN-US/library/office/ff745311(v=office.15).aspx
slide_title_frame <- slide_title[["TextFrame"]]
#slide_title_frame[["AutoSize"]] <- ms$ppAutoSizeNone
slide_title_frame[["AutoSize"]] <- ms$ppAutoSizeShapeToFitText

slide_title[["Top"]] <- 0
slide_title[["Left"]] <- 0
title_text <- slide_title[["TextFrame"]][["TextRange"]]
title_text[["Text"]] <- paste("Filmography: ",actor_name,sep="")
title_font <- title_text[["Font"]]

title_font[["Color"]][["RGB"]] <- pp_rgb(102,255,220)
title_font[["Size"]] <- 36
title_font[["Name"]] <- "Calibri"

# Add some decorative elements
diameter <- 100
# Add a gratuitous line
#line1 <- slide1[["Shapes"]]$AddLine(0,diameter/2,slide_width-diameter+2,diameter/2)
line1 <- slide1[["Shapes"]]$AddLine(0,diameter/2,slide_width,diameter/2)
line1_attr <- line1[["Line"]]
line1_attr[["Weight"]] <- 1
line1_attr[["ForeColor"]][["RGB"]] <- pp_rgb(102,255,220)

# Add a gratuitous circle
circle1 <- slide1[["Shapes"]]$AddShape(ms$msoShapeOval,slide_width-diameter,0,diameter,diameter)
circle1[["Top"]] <- 0
circle1[["Left"]] <- slide_width - diameter
circle1[["Width"]] <- diameter
circle1[["Height"]] <- diameter
circle1_color <- circle1[["Fill"]]
circle1_color[["ForeColor"]][["RGB"]] <- pp_rgb(102,255,220)
total_earnings <- format(sum(as.numeric(box_office$adjusted_gross))/1000000000,digits=3)
circle_text <- circle1[["TextFrame"]][["TextRange"]]
circle_text[["Text"]] <- paste(total_earnings,"BUSD","")
circle_font <- circle_text[["Font"]]
circle_font[["Name"]] <- "Calibri"
circle_font[["Size"]] <- 24

earnings_text <- slide1[["Shapes"]]$AddTextbox(ms$msoTextOrientationHorizontal,
                                               slide_width-4*diameter,diameter/2-25,diameter*3+5,20)
earnings_range <- earnings_text[["TextFrame"]][["TextRange"]]
earnings_range[["Text"]] <- "Total Box Office Earnings"
earnings_font <- earnings_range[["Font"]]
#earnings_font[["Color"]] <- pp_rgb(255,255,255)
earnings_font[["Color"]] <- pp_rgb(102,255,220)
earnings_font[["Size"]] <- 20
# When you are returning an object, you need to create a variable, and then set the properties
# This doesn't work
# earnings_range[["ParagraphFormat"]][["Alignment"]] <- ms$ppAlignRight
earnings_para <- earnings_range[["ParagraphFormat"]]
earnings_para[["Alignment"]] <- ms$ppAlignRight

# Set up some parameters for placing the shapes on the slide
# There are 60 movie images that need to be placed, so 20 columns by 3 rows
num_cols <- 20
num_rows <- ceiling(nrow(films)/num_cols)

# Based on the number of rows and columns, calculate the image height and width
image_width=floor(slide_width/num_cols)
image_height=floor(slide_height/(num_rows+3))
image_offset <- 3

# We need this to manipulate the images after they have been populated
images <- list()
image <- NULL

# Let's see how this will work on the slide, by arranging boxes as placeholders for the poster images
for (i in 1:nrow(films)) {
    x = image_width * ((i-1) %% num_cols)
    y = image_height*3 + image_height * ((i-1) %/% num_cols)
    r = slide1[["Shapes"]]$AddShape(
        ms$msoShapeRectangle,
        x, y,
        image_width, image_height)

}

# Neat! We can now place images and shapes quite precisely on the slide.
# Let's see how the images look on the slide
# We start by deleting the shapes

remove_shapes <- function(shape_name="Rectangle") {
    shp_todelete <- list()
    n_shapes <- slide1$Shapes()$Count()
    j=0
    for (i in 1:n_shapes) {
        shp <- slide1$Shapes(i)
        #print(paste0("index",i," - Shape: ",shp[["Name"]]))
        rect <- grepl(shape_name,shp[["Name"]])
        if (rect) {
            j <- j +1
            shp_todelete[[j]] <- slide1$Shapes(i)
            #shp$Delete()
            print(paste0("Shape : ",shp[["Name"]]," deleted..."))
        }
    }

    for (i in 1:j) {
        shp_todelete[[i]]$Delete()
    }

}

remove_shapes("Rectangle")

# Let's see how this works with poster images

for (i in 1:nrow(films)) {

    x = 0 + image_width * ((i-1) %% num_cols)
    y = image_height*image_offset + image_height * ((i-1) %/% num_cols)
    image_file <- gsub("/","\\\\",paste(getwd(),"/",films$img_file[i],sep=""))
    images[[as.character(i)]] <- slide1[["Shapes"]]$AddPicture(image_file,TRUE,FALSE,x+1,y+1,image_width-2,image_height-2)
    image <- images[[as.character(i)]]

    line <- image[["Line"]]
    line[["Style"]] <- ms$msoLineSingle
    line[["Weight"]] <- 2
    line[["ForeColor"]][["RGB"]] <- pp_rgb(255,255,255)
}

# The slide is still static.  Let's add animation.  We will include the ability to sort the
# poster images: 1-Sort by release date, and 2-Sort by film title
# We'll create two buttons for the sort, and then animate the posters based on a click event.

# We will remove the images and add them again, inlcuding the animation this time.
remove_shapes("Pict")

# Let's add the buttons.  These are standard rectangle shapes
# Add sort buttons

button_alpha <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350,150,150,40)
btn <- button_alpha[["TextFrame"]][["TextRange"]]
btn[["Text"]] <- "Alphabetical"
btn_font <- btn[["Font"]]
btn_font[["Size"]] <- 18
btn_font[["Color"]] <- pp_rgb(102,255,220)

btn_fill <- button_alpha[["Fill"]]
btn_fill[["Visible"]] <-0
btn_rgb <- btn_fill[["ForeColor"]][["RGB"]]

btn_line <- button_alpha[["Line"]]
btn_line[["ForeColor"]][["RGB"]] <- pp_rgb(102,255,220)

button_date <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350+180,150,150,40)
btn <- button_date[["TextFrame"]][["TextRange"]]
btn[["Text"]] <- "Release Year"
btn_font <- btn[["Font"]]
btn_font[["Size"]] <- 18
btn_font[["Color"]] <- pp_rgb(102,255,220)

btn_fill <- button_date[["Fill"]]
btn_fill[["Visible"]] <-0
btn_rgb <- btn_fill[["ForeColor"]][["RGB"]]

btn_line <- button_date[["Line"]]
btn_line[["ForeColor"]][["RGB"]] <- pp_rgb(102,255,220)

# As far as I can tell, the way animation seems to work in PowerPoint is as follows:
# 1 - An animation effect (fade, swivel and so on) is applied to a specific object
# 2 - The animation is added to a timeline, which specifies the sequence in which the effect will be executed
# 3 - The animation can include how it is triggered, the duration, and some effect specific behaviours (such as the path)
# 4 - An external trigger can also be specified

# Since we want to sort in two ways, we need to add two timeline sequences

seq_alpha = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_date = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()

# We can create a function that will apply animation to a shape, in this case the poster image
# The goal is to move the image from one point to another
# This function takes an timeline (sequence), the poster image that will be animated,
# the button that will trigger the animation, the path along which the image will move
# and the duration, and applies the animation and parameters to the target poster image

animate_image <- function(seq,image,trigger,path,duration=1.5) {
    effect <- seq$AddEffect(Shape=image,effectID=ms$msoAnimEffectPathDown,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    ani <- effect[["Behaviors"]]$Add(ms$msoAnimTypeMotion)
    aniMotionEffect <- ani[["MotionEffect"]]
    aniMotionEffect[["Path"]] <- path
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- trigger
    effectTiming[["Duration"]] <- duration
}

# Let's add some animation to the poster images
# The tricky bit here is the getting the path right
for (i in 1:nrow(films)) {

    # Populate the slide with the poster images
    x = 0 + image_width * ((i-1) %% num_cols)
    y = image_height*3 + image_height * ((i-1) %/% num_cols)
    img_file <- gsub("/","\\\\",paste(getwd(),"/",films$img_file[i],sep=""))
    images[[as.character(i)]] <- slide1[["Shapes"]]$AddPicture(img_file,TRUE,FALSE,x+1,y+1,image_width-1,image_height-1)
    image <- images[[as.character(i)]]

    # Some formatting
    line <- image[["Line"]]
    line[["Style"]] <- ms$msoLineSingle
    line[["Weight"]] <- 2
    line[["ForeColor"]][["RGB"]] <- pp_rgb(205,255,243)

    # Add the url to the IMDB page, and tool tip - the film title, character name and the release year
    link <- image$ActionSettings(ms$ppMouseClick)[["Hyperlink"]]
    link[["Address"]] <- films$url[i]
    link[["ScreenTip"]] <- paste0(films$title[i],"\nCharacter: ",films$character_name[i],"\nRelease Year: ",films$year[i])

    # Animate on button click, so the posters resort based on title or release year
    index <- which(films$title[order(films$title)]==films$title[i]) - 1
    l1 <- format((0 + image_width * (index %% num_cols) - x)/slide_width,digits=3)
    l2 <- format((image_height*3 + image_height * (index %/% num_cols) - y)/slide_height,digits=3)
    path <- paste0("M0,0 L",l1,",",l2)
    animate_image(seq_alpha,image,button_alpha,path)
    path <- paste0("M",l1,",",l2," L0,0")
    animate_image(seq_date,image,button_date,path)

}