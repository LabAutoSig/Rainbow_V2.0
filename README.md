# An expandable voice user interface as lab assistant based on an improved version of Google’s speech recognition

![image](https://github.com/user-attachments/assets/5b4cd0ab-f053-4479-b048-d053eabe3fcc)

## Introduction
Welcome to Rainbow – your open source voice user interface (VUI) for scientific lab. 
Rainbow is AutoIt-based, meaning we provide here mainly AutoIt scripts. You can 
view Rainbows skills and further explanations in [**DOI: 10.1038/s41598-023-46185-x**](https://doi.org/10.1038/s41598-023-46185-x).
This file gives step-by-step instructions of how to configure Rainbow to work on your 
Windows 10 PC. Please adapt script Rainbow_2.au3 (Mainscript), 
Rainbow_Timer.au3 and SMP_Script.au3.

**Important!**
When you finished the adjustment of all AutoIt.au3 scripts, please convert 
them into executables: *right-click on “Rainbow_2”, “Rainbow_Timer” and “SMP_script” >> Compile Script*

## Installation
1. Please install “AutoIt full installation” [AutoIt Download](https://www.autoitscript.com/site/autoit/downloads/)
2. Download and unzip Rainbow Zip Folder

The folder consists of several Word, Excel and text files and a folder called “program” 
where you can find all AutoIt scripts.

## Adaption
1. Please start Google Translate via **Google Chrome** and allow voice recording rights if necessary.
2. Open “Rainbow_2” via *right-click >> edit script*

To adapt Rainbow to your environment, you need to make adjustments in certain 
places. These places are marked with "!" in the rainbow script. Follow the note next to 
the “!”. Please adjust !1 - !62.
For adjustments of pixel numbers like mouse clicks or window sizes open the window 
tool *"AutoIt v3 Window Info" >> options >> Freeze*

## Adjusting Pixel coordinates requires the same conditions as expected running rainbow.
a. **Adjust mouse clicks: Do not move the corresponding window.** Mouse 
clicks consist of the x and y position of the mouse cursor. Open the Mouse tab 
on AutoIt v3 window info, open the interface with the desired button, position 
the cursor on the button (without clicking) and read the position of the mouse 
cursor under "Position" at AutoIt v3 Window Info. The two pixel coordinates are 
transferred exactly the same in the AutoIt script *Mouseclick("left",x,y)*.

b. **Adjust Window size and position:** Window Information consists of the position x,y and the size width and height. Open the window tab on AutoIt v3 window info. Try to adjust the window size near to 360, 615 (you can view the current state at “size”). Position the window of google translate in google chrome in the lower right corner of the desktop. Insert the Position (x,y) and the size (widh, height) into the AutoIt script. Window name is given in “Titel”. *WinMove(“WindowName”,x,y,widh,height)*. 
![image](https://github.com/user-attachments/assets/594e1860-1955-4c8a-bd9a-53b79a52a8c2)

c. **Adjust PixCheckColor:** Note for !19 – Please enter a misspelled word in GTS to simulate an error. Please position the mouse cursor on the blue star (see circled in red in the figure) and insert its x,y position in the AutoIt command PixelCheckColor on line !19. *PixelCheckColor(x,y)* 
![image](https://github.com/user-attachments/assets/97f87821-7a7b-44f1-be44-ee6b3c8c5c78)

d. **Adjust PixelCheckSum:** This AutoIt command counts all pixel values in a specified area. The area is specified as a rectangle with the x,y coordinates of the upper left corner and the lower right corner. Please specify the coordinates as shown in the figure (red dots) in lines !21 and !23. *PixelCheckColor(x,y,x,y)* 
![image](https://github.com/user-attachments/assets/93cf980b-df74-4db8-b8eb-a4c6feb2b9cb)

Note for !35 and !36 – To determine the pixel sum of a specific area, please open 
and save a new AutoIt file. Write:
*WinActivate(“Titel of window line !34”)
Sleep(1000)
MsgBox(0,””,PixelCheckSum(x,y,x,y)*
For the x,y,x,y coordinates insert the pixel coordinates of the red dots in the following 
figure.
![image](https://github.com/user-attachments/assets/6b9f7afc-412a-43ec-b72c-748784ace8d0)
## Run the script
Start the script by pressing key F5 >> Go back to rainbow script and insert the 
displayed pixel sum (Message Box) in line !36.

*If $sum = xxxxxxx then*
