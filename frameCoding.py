#!python
#coding: ascii
import os
import cv2
import numpy as np
import openpyxl
from openpyxl import Workbook
import pandas as pd

#INDEX FOR FRAME #
frameIndex = 0
# INDEX FOR ROW # (file starts @ 1, so inputs start @ 2)
rowIndex = 2

# Create a VideoCapture object and read from input file
# If the input is the camera, pass 0 instead of the video file name
# CHANGE VIDEO TITLE WHEN NEEDED
videoFile = 'exampleVideo.mp4'
cap = cv2.VideoCapture(videoFile)

# #print total # of frames
# totalFrames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
# print(totalFrames)

# CREATE EXCEL SHEET
excelTitle = videoFile + '.xlsx'

excelFile = Workbook()
excelFile.save(excelTitle)
#grab active sheet
sheet = excelFile.active
#add column titles
sheet["A1"] = "Frame #"
sheet["B1"] = "Time"
sheet["C1"] = "Action Class"

# Check if camera or file opened successfully
if (cap.isOpened()== False):
  print("Error opening video stream or file")

# Read until video is completed
while(cap.isOpened()):
  # Capture frame-by-frame
  ret, frame = cap.read()
  if ret == True:

    # Display the resulting frame
    cv2.imshow('Frame',frame)


    # waitKey(0) keeps frame on screen indefinitely until keypress
    key = cv2.waitKey(0)
    # TO TEST WHICH KEY IS WHAT ASCII CODE
    # print('You pressed %d (0x%x), LSB: %d (%s)' % (key, key, key % 256,
    #     repr(chr(key%256)) if key%256 < 128 else '?'))

    #get timestamp
    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
    print(frameTime)
    # ***BUG*** LAST FEW FRAMES, THE _MSEC PROPERTY OUTPUTS 0??
    # PERSONAL FIX:
    # totalFrames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    # print(totalFrames)


    if key == 0x31: # KEYPRESS 1 = NEUTRAL
        print("KEY 1 = NEUTRAL")

        # #grab active sheet
        # sheet = excelFile.active

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 1

    elif key == 0x32: # KEYPRESS 2 = BELLY TOUCH
        print("KEY 2 = BELLY TOUCH")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 2

    elif key == 0x33: # KEYPRESS 3 = SINGING/HUMMING
        print("KEY 3 = SINGING/HUMMING")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 3

    elif key == 0x34: # KEYPRESS 4 = DANCING
        print("KEY 4 = DANCING")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 4

    elif key == 0x35: # KEYPRESS 5 = TALK TO BABY
        print("KEY 5 = TALK TO BABY")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 5

    elif key == 0x36: # KEYPRESS 6 = TALK TO MUSICIAN
        print("KEY 6 = TALK TO MUSICIAN")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 6

    elif key == 0x37: # KEYPRESS 7 = CRYING
        print("KEY 7 = CRYING")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
    #    frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = 7
    else:
        print("INVALID")

        #add frame # to 1st column
        sheet['A' + str(rowIndex)] = frameIndex

        #add time code to 2nd column
        #get timestamp
#        frameTime = cap.get(cv2.CAP_PROP_POS_MSEC)
        sheet['B' + str(rowIndex)] = frameTime

        #add action code to 3rd column
        sheet['C' + str(rowIndex)] = "INVALID"
    # Press Q on keyboard to  exit
    #if cv2.waitKey(25) & 0xFF == ord('q'):
    #  break

    #increment frameIndex and rowIndex
    frameIndex += 1
    rowIndex += 1

  # Break the loop
  else:
    break

#SAVE EXCEL FILE
excelFile.save(excelTitle)

# When everything done, release the video capture object
cap.release()

# Closes all the frames
cv2.destroyAllWindows()
