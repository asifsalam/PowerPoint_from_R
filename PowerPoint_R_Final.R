# Created by: Asif Salam
# email: asif.salam@outlook.com

library(RDCOMClient)
library(plyr)
library(dplyr)
library(stringr)
library(rvest)

# download and read in the data files
download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_films.tsv",
              destfile = "films.tsv")
download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_box_office.tsv",
              destfile = "eastwood_box_office.tsv")

films <- read.table("eastwood_films.csv",header=TRUE, stringsAsFactors=FALSE)
box_office <- read.table("eastwood_box_office.csv",header=TRUE, stringsAsFactors=FALSE)
source("mso.txt")
actor_name <- "Clint Eastwood"

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
#films[55,"img_file"] <- "img/img55.jpg"
#films[54,"img_file"] <- "img/img54.jpg"
#films[52,"img_file"] <- "img/img52.jpg"

########## Creating the PowerPoint Slide ###################
# Create the PowerPoint slide
pp <- COMCreate("Powerpoint.Application")
pp[["Visible"]] = 1
presentation <- pp[["Presentations"]]$Add()
#slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutTitleOnly)
slide_width <- presentation[["PageSetup"]]$SlideWidth()
slide_height <- presentation[["PageSetup"]]$SlideHeight()

# Set some slide attributes
slide_color <- slide1[["ColorScheme"]]$Colors(ms$ppBackground)
slide_color[["RGB"]] <- pp_rgb(0,0,0)

img_file <- gsub("/","\\\\",paste(getwd(),"/","img/clint_background_1.png",sep=""))
bg_image <- slide1[["Shapes"]]$AddPicture(img_file,TRUE,FALSE,0,0,slide_width,slide_height)
bg_rect <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,0, 0,slide_width, slide_height)
bg_rect_fill <- bg_rect[["Fill"]]
bg_rect_fill[["ForeColor"]][["RGB"]] <- pp_rgb(102, 25, 13)
bg_rect_fill[["Transparency"]] <- 0.1
bg_rect_line <- bg_rect[["Line"]]
bg_rect_line[["ForeColor"]][["RGB"]] <- pp_rgb(102,25,13)

slide_title <- slide1[["Shapes"]][["Title"]]
slide_title_color <- slide1[["ColorScheme"]]$Colors(ms$ppTitle)
slide_title_color[["RGB"]] <- pp_rgb(243,211,129)
slide_title$ZOrder(ms$msoBringToFront)

# Add a title
# AutoSize: https://msdn.microsoft.com/EN-US/library/office/ff745311(v=office.15).aspx
slide_title_frame <- slide_title[["TextFrame"]]
slide_title_frame[["AutoSize"]] <- ms$ppAutoSizeShapeToFitText
slide_title[["Top"]] <- 0
slide_title[["Left"]] <- 0
title_text <- slide_title[["TextFrame"]][["TextRange"]]
title_text[["Text"]] <- paste("Filmography: ",actor_name,sep="")
title_font <- title_text[["Font"]]
title_font[["Color"]][["RGB"]] <- pp_rgb(243,211,129)
title_font[["Size"]] <- 36
title_font[["Name"]] <- "Calibri"

# Add some decorative elements
diameter <- 85
# Add a gratuitous line
line1 <- slide1[["Shapes"]]$AddLine(0,diameter/2,slide_width,diameter/2)
line1_attr <- line1[["Line"]]
line1_attr[["Weight"]] <- 1
line1_attr[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

# Add a gratuitous circle, showing total box office earnings
circle1 <- slide1[["Shapes"]]$AddShape(ms$msoShapeOval,slide_width-diameter,0,diameter,diameter)
circle1[["Top"]] <- 0
circle1[["Left"]] <- slide_width - diameter
circle1[["Width"]] <- diameter
circle1[["Height"]] <- diameter
circle1_color <- circle1[["Fill"]]
circle1_color[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
total_earnings <- format(sum(as.numeric(box_office$adjusted_gross))/1000000000,digits=3)
circle_frame <- circle1[["TextFrame"]]
circle_frame[["MarginTop"]] <- 0
circle_frame[["MarginLeft"]] <- 0
circle_frame[["MarginRight"]] <- 0
circle_frame[["MarginBottom"]] <- 0
circle_text <- circle1[["TextFrame"]][["TextRange"]]
circle_text[["Text"]] <- paste(total_earnings,"BUSD","")
circle_font <- circle_text[["Font"]]
circle_font[["Name"]] <- "Calibri"
circle_font[["Size"]] <- 24
circle_font[["Color"]][["RGB"]] <- pp_rgb(102,25,13)
circle_font[["Bold"]] <- 1
circle_line <- circle1[["Line"]]
circle_line[["Weight"]] <- 2
circle_line[["ForeColor"]][["RGB"]] <- pp_rgb(102,25,13)

earnings_text <- slide1[["Shapes"]]$AddTextbox(ms$msoTextOrientationHorizontal,
                                               slide_width-4*diameter,diameter/2-25,diameter*3+1,20)
earnings_range <- earnings_text[["TextFrame"]][["TextRange"]]
earnings_range[["Text"]] <- "Total Box Office Earnings"
earnings_font <- earnings_range[["Font"]]
earnings_font[["Color"]] <- pp_rgb(243,211,129)
earnings_font[["Size"]] <- 20
# When you are returning an object, you need to create a variable, and then set the properties
# This doesn't work
# earnings_range[["ParagraphFormat"]][["Alignment"]] <- ms$ppAlignRight
earnings_para <- earnings_range[["ParagraphFormat"]]
earnings_para[["Alignment"]] <- ms$ppAlignRight

# Set up parameters for placing the shapes on the slide
# There are 60 movie images that need to be placed, so 20 columns by 3 rows
num_cols <- 20
num_rows <- ceiling(nrow(films)/num_cols)

# Based on the number of rows and columns, calculate the image height and width
image_width=floor((slide_width)/num_cols)
image_height=floor(slide_height/(num_rows+3))
image_offset <- 3

# We need this to manipulate the images after they have been populated
images <- list()
image <- NULL

# As far as I can tell, the way animation seems to works in PowerPoint is as follows:
# 1 - An animation effect (fade, swivel and so on) is applied to a specific object
# 2 - The animation is added to a timeline, which specifies the sequence in which the effect will be executed
# 3 - The animation can include how it is triggered, the duration, and some effect specific behaviours (such as the path)
# 4 - An external trigger can also be specified
# We can create a function that will apply animation to a shape, in this case the poster image
# The goal is to move the image from one point to another
# This function takes an timeline (sequence), the poster image that will be animated,
# the button that will trigger the animation, the path along which the image will move
# and the duration, and applies the animation and parameters to the target poster image
# (See the section at the bottom for function - animate_shape())


# Add some explantory text
sort_text <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 400,10,100,40)
sort_range <- sort_text[["TextFrame"]][["TextRange"]]
sort_para <- sort_range[["ParagraphFormat"]]
sort_para[["Alignment"]] <- ms$ppAlignLeft
sort_range[["Text"]] <- "Sort posters by: "
sort_font <- sort_range[["Font"]]
sort_font[["Size"]] <- 14
sort_font[["Color"]] <- pp_rgb(233,174,27)
sort_text[["Width"]] <- 110
sort_text[["Height"]] <- 15
sort_text[["Top"]] <- image_height*image_offset - 20 - 3
#sort_text[["Left"]] <- slide_width - 205
sort_text[["Left"]] <- 0
sort_fill <- sort_text[["Fill"]]
sort_fill[["Visible"]] <- 0
sort_line <- sort_text[["Line"]]
sort_line[["Visible"]] <- 0

# Create buttons that will animate the poster images
button_alpha <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 300,10,100,40)
bta_fill <- button_alpha[["Fill"]]
bta_fill[["Visible"]] <-0
bta_rgb <- bta_fill[["ForeColor"]][["RGB"]]

bta_line <- button_alpha[["Line"]]
bta_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

bta <- button_alpha[["TextFrame"]][["TextRange"]]
bta[["Text"]] <- "Title"
bta_font <- bta[["Font"]]
bta_font[["Size"]] <- 14
bta_font[["Color"]] <- pp_rgb(243,211,129)

button_alpha[["Width"]] <- 90
button_alpha[["Height"]] <- 15
button_alpha[["Top"]] <- image_height*image_offset - 20 - 3
button_alpha[["Left"]] <- sort_text[["Width"]] - 2

# Add a sequence to the slide timeline: https://msdn.microsoft.com/en-us/library/office/ff746823.aspx
seq_main <- slide1[["TimeLine"]][["MainSequence"]]

animation_start(seq_main,slide_title,ms$msoAnimEffectDescend,ms$msoAnimTriggerWithPrevious,
              0, -20,0,0,1,0)
animation_start(seq_main,line1,ms$msoAnimEffectFly,ms$msoAnimTriggerAfterPrevious,
              100, diameter/slide_height,0,diameter/slide_height,1,0)
animation_start(seq_main,circle1,ms$msoAnimEffectFly,ms$msoAnimTriggerWithPrevious,
              -100, diameter/slide_height,0,diameter/slide_height,1,0)
animation_start(seq_main,earnings_text,ms$msoAnimEffectFly,ms$msoAnimTriggerWithPrevious,
              0, 100,(slide_width-4*diameter)/slide_width,diameter/slide_height,1,0)
animation_start(seq_main,sort_text,ms$msoAnimEffectWipe,ms$msoAnimTriggerAfterPrevious,
              0, 0,0,0,1,0)
animation_start(seq_main,button_alpha,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
              0, 0,0,0,1,0)

button_date <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350+160,10,150,40)
btd <- button_date[["TextFrame"]][["TextRange"]]
btd[["Text"]] <- "Release Year"

btd_fill <- button_date[["Fill"]]
btd_fill[["Visible"]] <-0

btd_line <- button_date[["Line"]]
btd_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

btd_font <- btd[["Font"]]
btd_font[["Size"]] <- 14
btd_font[["Color"]] <- pp_rgb(243,211,129)

button_date[["Width"]] <- 90
button_date[["Height"]] <- 15
button_date[["Top"]] <- image_height*image_offset - 20 - 3
#button_date[["Left"]] <- slide_width - button_date[["Width"]] - 10
button_date[["Left"]] <- sort_text[["Width"]] - 2

# https://msdn.microsoft.com/EN-US/library/office/ff745511.aspx

seq_alpha = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_date = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_identify = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_earnings = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_year = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()

seq_alpha2 = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_date2 = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_earnings2 = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_year2 = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()

toggle_button(seq_alpha2,button_alpha,button_date,1)
toggle_button(seq_date2,button_date,button_alpha,1)

for (i in 1:nrow(films)) {
    x = 1 + image_width * ((i-1) %% num_cols)
    y = image_height*image_offset + image_height * ((i-1) %/% num_cols)
    img_file <- gsub("/","\\\\",paste(getwd(),"/",films$img_file[i],sep=""))
    images[[as.character(i)]] <- slide1[["Shapes"]]$AddPicture(img_file,TRUE,FALSE,x,y,image_width-2,image_height-2)
    image <- images[[as.character(i)]]

    line <- image[["Line"]]
    line[["Style"]] <- ms$msoLineSingle
    line[["Weight"]] <- 1
    line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

    #glow <- image[["Glow"]]
    #glow[["Radius"]] <- 3
    #glow[["Transparency"]] <- 0.7
    #glow[["Color"]] <- pp_rgb(200,200,200)

    link <- image$ActionSettings(ms$ppMouseClick)[["Hyperlink"]]
    link[["Address"]] <- films$url[i]
    link[["ScreenTip"]] <- paste0(films$title[i],"\nCharacter: ",films$character_name[i],"\nRelease Year: ",films$year[i])

    index <- which(films$title[order(films$title)]==films$title[i]) - 1
    l1 <- format((0 + image_width * (index %% num_cols) - x)/slide_width,digits=3)
    l2 <- format((image_height*image_offset + image_height * (index %/% num_cols) - y)/slide_height,digits=3)
    path <- paste0("M0,0 L",l1,",",l2)
    animate_image(seq_alpha,image,button_alpha,path,2.0)
    path <- paste0("M",l1,",",l2," L0,0")
    animate_image(seq_date,image,button_date,path,2.0)
    trigger_seq <- ms$msoAnimTriggerWithPrevious
    if (i == 1) trigger_seq <- ms$msoAnimTriggerAfterPrevious

    animation_start(seq_main,image,ms$msoAnimEffectDissolve,trigger_seq,
                  0, 0,0,0,0.5,0.1*i)
}

chart_top <- 70
max_value <- max(box_office$adjusted_gross)
max_height <- slide_height/3.6
margin_left <- 45
margin_right <- 5
max_bar_h <- 130
bar_gap <- 0
num_bars <- nrow(box_office)

bar_w <- trunc((slide_width - (margin_left + margin_right))/num_bars - bar_gap)

chart_text <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,margin_left,50,200,20)
chart_text_range <- chart_text[["TextFrame"]][["TextRange"]]
chart_para <- chart_text_range[["ParagraphFormat"]]
chart_para[["Alignment"]] <- ms$ppAlignCenter
chart_text_range[["Text"]] <- "Box Office Earnings per Movie (MUSD)"
chart_text[["Width"]] <- 400
chart_text_fill <- chart_text[["Fill"]]
chart_text_fill[["Visible"]] <- 0
chart_text_line <- chart_text[["Line"]]
chart_text_line[["Visible"]] <- 0
chart_font <- chart_text_range[["Font"]]
chart_font[["Color"]] <- pp_rgb(243,211,129)
chart_text[["Left"]] <- slide_width/2 - 200

sort_text2<- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 2*80-100-5,
                                            image_height*image_offset - 20 - 10,100,15)
sort_range <- sort_text2[["TextFrame"]][["TextRange"]]
sort_para <- sort_range[["ParagraphFormat"]]
sort_para[["Alignment"]] <- ms$ppAlignRight

sort_range[["Text"]] <- "Sort chart by: "
sort_fill <- sort_text2[["Fill"]]
sort_fill[["Visible"]] <- 0
sort_line <- sort_text2[["Line"]]
sort_line[["Visible"]] <- 0
sort_font <- sort_range[["Font"]]
sort_font[["Color"]] <- pp_rgb(233,174,27)
sort_font[["Size"]] <- 14
sort_text2[["Width"]] <- 100
sort_text2[["Left"]] <- 665
sort_text2[["Top"]] <- 55

btn_release <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350+160,10,150,40)
btd <- btn_release[["TextFrame"]][["TextRange"]]
btd[["Text"]] <- "Release Year"

btd_fill <- btn_release[["Fill"]]
btd_fill[["Visible"]] <-0

btd_line <- btn_release[["Line"]]
btd_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

btd_font <- btd[["Font"]]
btd_font[["Size"]] <- 14
btd_font[["Color"]] <- pp_rgb(243,211,129)

btn_release[["Width"]] <- 90
btn_release[["Height"]] <- 15
btn_release[["Top"]] <- 55
btn_release[["Left"]] <- sort_text2[["Left"]] + sort_text2[["Width"]]

animation_start(seq_main,chart_text,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
              0, 0,0,0,1,4.5)
animation_start(seq_main,btn_release,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
              0, 0,0,0,1,4.5)
animation_start(seq_main,sort_text2,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
              0, 0,0,0,1,4.5)

btn_earnings <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350+160,10,150,40)
btd <- btn_earnings[["TextFrame"]][["TextRange"]]
btd[["Text"]] <- "Earnings"

btd_fill <- btn_earnings[["Fill"]]
btd_fill[["Visible"]] <-0

btd_line <- btn_earnings[["Line"]]
btd_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

btd_font <- btd[["Font"]]
btd_font[["Size"]] <- 14
btd_font[["Color"]] <- pp_rgb(243,211,129)

btn_earnings[["Width"]] <- 90
btn_earnings[["Height"]] <- 15
btn_earnings[["Top"]] <- 55
btn_earnings[["Left"]] <- sort_text2[["Left"]] + sort_text2[["Width"]]

toggle_button(seq_year2,btn_release,btn_earnings,1)
toggle_button(seq_earnings2,btn_earnings,btn_release,1)

for (i in 1:nrow(box_office)) {
    image_num <- which(box_office$key[i]==films$key)
    image <- images[[as.character(image_num)]]

    x <- margin_left + (i-1)*(bar_w + bar_gap)
    bar_height <- scale_bar_height(max_value,max_height,box_office$adjusted_gross[i])
    y <- chart_top + max_height - bar_height

    bar <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,x+1,y,bar_w-2,bar_height-1)
    bar_fill <- bar[["Fill"]]
    #bar_fill[["ForeColor"]][["RGB"]] <- pp_rgb(0,138,105)
    #bar_fill[["ForeColor"]][["RGB"]] <- pp_rgb(233,174,27)
    bar_fill[["ForeColor"]][["RGB"]] <- pp_rgb(217,161,21)

    index <- which(box_office$index[order(box_office$index)]==box_office$index[i]) - 1
    l1 <- format((margin_left + bar_w * index - x)/slide_width,digits=3)
    l2 <- format((0)/slide_height,digits=3)

    line <- bar[["Line"]]
    line[["Style"]] <- ms$msoLineSingle
    line[["Weight"]] <- 2
    line[["ForeColor"]][["RGB"]] <- pp_rgb(118, 50, 39)
    line[["Transparency"]] <- 0
    line[["Visible"]] <- 0

    bar_textframe <- bar[["TextFrame"]]
    bar_textframe[["TextRange"]][["Text"]] <- ""
    bar_textframe[["TextRange"]][["Text"]] <- box_office$title[i]
    bar_textframe[["Orientation"]] <- ms$msoTextOrientationUpward
    bar_textframe[["WordWrap"]] <- ms$msoFalse
    bar_font <- bar_textframe[["TextRange"]][["Font"]]
    bar_font[["Size"]] <- 12
    bar_para <- bar_textframe[["TextRange"]][["ParagraphFormat"]]
    bar_para[["Alignment"]] <- ms$msoAlignLeft
    bar_font[["Color"]][["RGB"]] <- pp_rgb(247,224,167)

    year <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,x,y+bar_height+3,bar_w,20)
    year_textframe <- year[["TextFrame"]]
    year_textframe[["TextRange"]][["Text"]] <- box_office$year[i]
    year_textframe[["Orientation"]] <- ms$msoTextOrientationUpward
    year_textframe[["WordWrap"]] <- ms$msoFalse
    year_font <- year_textframe[["TextRange"]][["Font"]]
    year_font[["Size"]] <- 9
    year_font[["Color"]][["RGB"]] <- pp_rgb(247,224,167)
    year_para <- year_textframe[["TextRange"]][["ParagraphFormat"]]
    year_para[["Alignment"]] <- ms$msoAlignLeft
    year_fill <- year[["Fill"]]
    year_fill[["Visible"]] <- 0
    year_fill[["ForeColor"]][["RGB"]] <- pp_rgb(0,0,0)
    year_line <- year[["Line"]]
    year_line[["Visible"]] <- 0

    click_box <- slide1[["Shapes"]]$AddShape(1,x,chart_top,bar_w,max_height)
    click_box_fill <- click_box[["Fill"]]
    click_box_fill[["Visible"]] <- 1
    click_box_fill[["ForeColor"]][["RGB"]] <- pp_rgb(247,224,167)
    click_box_fill[["Transparency"]] <- 0.99

    line <- click_box[["Line"]]
    #line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
    line[["Visible"]] <- 0

    path1 <- paste0("M0,0 L",l1,",",l2)
    animate_image(seq_earnings,bar,btn_release,path1)
    animate_image(seq_earnings,year,btn_release,path1)
    animate_image(seq_earnings,click_box,btn_release,path1)

    path2 <- paste0("M",l1,",",l2," L0,0")
    animate_image(seq_year,bar,btn_earnings,path2)
    animate_image(seq_year,year,btn_earnings,path2)
    animate_image(seq_year,click_box,btn_earnings,path2)

    animate_bar(seq_identify,image,click_box,2)
    trigger_seq <- ms$msoAnimTriggerWithPrevious
    animation_start(seq_main,bar,ms$msoAnimEffectEaseIn,trigger_seq,
                   0, 0,0,0,1,5+0.1*i)
    animation_start(seq_main,year,ms$msoAnimEffectEaseIn,trigger_seq,
                   0, 0,0,0,1,5+0.1*i)
    animation_start(seq_main,click_box,ms$msoAnimEffectEaseIn,trigger_seq,
                   0, 0,0,0,1,5+0.1*i)
}

grid_pos1 <- scale_bar_height(max_value,max_height,round(max_value,-8)/3)
x <- margin_left - 5
y <- chart_top + max_height - grid_pos1
line1 <- slide1[["Shapes"]]$AddLine(x,y,margin_left + (nrow(box_office))*(bar_w + bar_gap),y)
line1_attr <- line1[["Line"]]
line1_attr[["Weight"]] <- 1
line1_attr[["ForeColor"]][["RGB"]] <- pp_rgb(247,224,167)
line1_attr[["Transparency"]] <- 0.6
label1 <- slide1[["Shapes"]]$AddTextBox(ms$msoTextOrientationHorizontal,0,y,margin_left,20)

label1_fill <- label1[["Fill"]]
label1_fill[["Visible"]] <- 0
label1[["Top"]] <- y - label1[["Height"]]/2
label1_frame <- label1[["TextFrame"]]
label1_frame[["TextRange"]] <- round(max_value,-8)/3000000
label1_font <- label1[["TextFrame"]][["TextRange"]][["Font"]]
label1_font[["Color"]][["RGB"]] <- pp_rgb(243,211,129)
label1_font[["Size"]] <- 14
label1_para <- label1[["TextFrame"]][["TextRange"]][["ParagraphFormat"]]
label1_para[["Alignment"]] <- ms$ppAlignRight

grid_pos2 <- scale_bar_height(max_value,max_height,round(max_value,-8)*2/3)
y <- chart_top + max_height - grid_pos2
line2 <- slide1[["Shapes"]]$AddLine(x,y,margin_left + (nrow(box_office))*(bar_w + bar_gap),y)
line2_attr <- line2[["Line"]]
line2_attr[["Weight"]] <- 1
line2_attr[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
line2_attr[["Transparency"]] <- 0.6

label2 <- slide1[["Shapes"]]$AddTextBox(ms$msoTextOrientationHorizontal,0,y,margin_left,20)
label2_fill <- label1[["Fill"]]
label2_fill[["Visible"]] <- 0
label2[["Top"]] <- y - label2[["Height"]]/2
label2_frame <- label2[["TextFrame"]]
label2_frame[["TextRange"]] <- round(max_value,-8)*2/3000000
label2_font <- label2_frame[["TextRange"]][["Font"]]
label2_font[["Size"]] <- 14
label2_font[["Color"]][["RGB"]] <- pp_rgb(243,211,129)
label2_para <- label2_frame[["TextRange"]][["ParagraphFormat"]]
label2_para[["Alignment"]] <- ms$ppAlignRight

animation_start(seq_main,line1,ms$msoAnimEffectEaseIn,trigger_seq,
              0, 0,0,0,1,6+0.1*i)
animation_start(seq_main,line2,ms$msoAnimEffectEaseIn,trigger_seq,
              0, 0,0,0,1,6+0.1*i)
animation_start(seq_main,label1,ms$msoAnimEffectEaseIn,trigger_seq,
                0, 0,0,0,1,6+0.1*i)
animation_start(seq_main,label2,ms$msoAnimEffectEaseIn,trigger_seq,
                0, 0,0,0,1,6+0.1*i)

# ========================= Functions ===================== #

toggle_button <- function(seq,button1,button2,duration=1.5) {
    effect <- seq$AddEffect(Shape=button1,effectID=ms$msoAnimEffectDissolve,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    effect[["Exit"]] <- 1
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerOnShapeClick
    effectTiming[["TriggerShape"]] <- button1
    effectTiming[["Duration"]] <- duration
    effect <- seq$AddEffect(Shape=button2,effectID=ms$msoAnimEffectDissolve,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    effect[["Exit"]] <- 0
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- button1
    effectTiming[["Duration"]] <- duration
}

animation_start <- function(seq,shape,effectID,trigger,from_x,from_y,to_x,to_y,duration,delay_time) {

    effect <- seq$AddEffect(Shape=shape,EffectID=effectID,Trigger=trigger)
    ani <- effect[["Behaviors"]]$Add(ms$msoAnimTypeMotion)
    # MotionEffect Object: https://msdn.microsoft.com/EN-US/library/office/ff745317(v=office.15).aspx
    aniMotionEffect <- ani[["MotionEffect"]]
    # https://msdn.microsoft.com/EN-US/library/office/ff745317.aspx
    aniMotionEffect[["FromX"]] <- from_x
    aniMotionEffect[["ToX"]] <- to_x
    aniMotionEffect[["FromY"]] <- from_y
    aniMotionEffect[["ToY"]] <- to_y
    effectTiming <- effect[["Timing"]]
    effectTiming[["Duration"]] <- duration
    effectTiming[["TriggerDelayTime"]] <- delay_time
}

animate_image <- function(seq,image,trigger,path,duration=2.0) {
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

animate_bar <- function(seq,image,trigger,duration=1.5,delay_time=0) {
    effect <- seq$AddEffect(Shape=image,effectID=ms$msoAnimEffectSpin,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- trigger
    effectTiming[["Duration"]] <- duration
    effectTiming[["TriggerDelayTime"]] <- delay_time

}

pp_rgb <- function(r,g,b) {
    return( r + g*256 + b*256^2)
}

scale_bar_height <- function(max_value,max_height,value) {
    bar_height <- value*max_height/max_value
}

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