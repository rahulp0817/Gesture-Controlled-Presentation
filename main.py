import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import os
import pygetwindow as gw
import numpy as np
import aspose.slides as slides
import aspose.pydrawing as drawing
import os


# Open the presentation and slide show window

current_dir = os.path.dirname(os.path.abspath(__file__))
pptx_file = os.path.join(current_dir, "rvmb.pptx")

Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open(pptx_file)
print(Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters

width, height = 900, 720
gestureThreshold = 300


# Camera Setup

cap = cv2.VideoCapture(0) # Webcam
cap.set(3, width)
cap.set(4, height) 


# Hand Detector

detectorHand = HandDetector(detectionCon = 0.8, maxHands = 1)

# Variables

imgList = []
delay = 30 # delay for button press or move of slide
buttonPressed = False
counter = 0
drawMode = False
imgNumber = 20
delayCounter = 0
annotations = [[]] # whenever we want to draw on index finger
annotationNumber = -1
annotationStart = False
drawing_points = []  


while True:
    success, img = cap.read() # Get image frame and Find the hand and its landmarks
    
    hands, img = detectorHand.findHands(img)  # with draw

    if hands and not buttonPressed:  # If hand is detected
        hand = hands[0]
        cx, cy = hand["center"]

        lmList = hand["lmList"]  # List of 21 Landmark points

        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up

        if cy <= gestureThreshold:  #if hand is at the height of the face
            
            # gesture 1 - left

            if fingers == [1, 1, 1, 1, 1]:
                print("Left")
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Next()
                    imgNumber -= 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

            # gesture 2 - right
                    
            if fingers == [1, 0, 0, 0, 0]:
                print("Right")
                buttonPressed = True
                if imgNumber > 0 :
                    Presentation.SlideShowWindow.View.Previous()
                    imgNumber += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

            ## gesture 3 - pointer display
                    
            if fingers == [0, 1, 1, 0, 0]:  # Pointer gesture
                indexFingerTip = lmList[8]  # Index Finger Tip position
                cv2.circle(img, (indexFingerTip[0], indexFingerTip[1]), 15, (0, 0, 255), cv2.FILLED)
              
            ## gesture 4 - draw
            
            if fingers == [0, 1, 0, 0]:  # Drawing Gesture
                indexFingerTip = lmList[8]  # Index Finger Tip position
                drawing_points.append(indexFingerTip)
                for i in range(1, len(drawing_points)):
                    cv2.line(img, (drawing_points[i-1][0], drawing_points[i-1][1]), (drawing_points[i][0], drawing_points[i][1]), (255, 0, 0), 15)
            
            # gesture 5 - erase
               
 
    else:
        annotationStart = False
 
    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False
 
    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(imgCurrent, annotation[j - 1], annotation[j], (0, 0, 200), 12)
 
    cv2.imshow("Image", img)

    ## Bring the webcam to the top of presentation

    window = gw.getWindowsWithTitle('Image')[0]
    window.activate()
    
    key = cv2.waitKey(1)
    if key == ord('q'):
        break

