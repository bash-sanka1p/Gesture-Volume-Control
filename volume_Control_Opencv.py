# -*- coding: utf-8 -*-
"""
Created on Thu Apr  7 23:58:18 2022

@author: Sankalp Singh Bais
"""

import cv2
import mediapipe as mp
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import numpy as np

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

def main():
    # print(volume.GetVolumeRange())
    #volume range this pc and pycaw -> (0,100)=(-65.25,0)

    cap = cv2.VideoCapture(0) # capture device
    cv2.startWindowThread()

    mpHands = mp.solutions.hands
    hands = mpHands.Hands()
    mpDraw = mp.solutions.drawing_utils


    while True:
        try :
            success, fimg = cap.read()
            img = cv2.flip(fimg,1)  #flipping image horizontally
            imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            results =hands.process(imgRGB)
            #print(results.multi_hand_landmarks)

            if results.multi_hand_landmarks:
                for handLms in results.multi_hand_landmarks:
                    lmList=[]
                    for id , lm in enumerate(handLms.landmark):

                        #print(id , lm)
                        h , w , c = img.shape
                        cx , cy = int(lm.x*w) ,int(lm.y*h)
                        #print(id,cx,cy)
                        lmList.append([id, cx, cy])
                        #uncomment below line if want to visulize the hand connections
                        # mpDraw.draw_landmarks(img , handLms , mpHands.HAND_CONNECTIONS)

                        #print(results.multi_hand_landmarks)

                if lmList:
                    #coordinates of the fore finger and thumb and encircling it
                    x1, y1, = lmList[4][1], lmList[4][2]
                    x2, y2, = lmList[8][1], lmList[8][2]
                    cv2.circle(img , (x1 ,y1) ,15 , (134, 138, 204) ,cv2.FILLED)
                    cv2.circle(img , (x2 ,y2) ,15 , (134, 138, 204) ,cv2.FILLED)
                    #line joining above
                    cv2.line(img, (x1,y1) , (x2,y2) , (37, 157, 140) ,3)

                    z1 = (x1+x2)//2
                    z2 = (y1+y2)//2

                    length = math.hypot(x2-x1, y2-y1)
                    print(length)

                    if length<30:
                        cv2.circle(img,(z1,z2), 10, (0,0,0), cv2.FILLED)
                    else:
                        cv2.circle(img,(z1,z2), 10, (255,255,255), cv2.FILLED)


                # length between 30 to 200
                volRange = volume.GetVolumeRange()
                minVol = volRange[0]
                maxVol = volRange[1]
                vol = np.interp(length, [30, 200], [minVol, maxVol])
                volBar = np.interp(length, [30,200], [300,100])
                volPer = np.interp(length, [30,200], [0,100])
                print(volRange, minVol, maxVol, vol)

                volume.SetMasterVolumeLevel(vol, None)

                # cv2.rectangle(img, (50, 100), (85, 300), (40, 123, 104))
                cv2.rectangle(img, (50,int(volBar)), (85, 300),  (40, 123, 104), cv2.FILLED)
                cv2.putText(img ,'MUTE' if volPer==0 else f'{int(volPer)}%' ,
                            (15,350),cv2.FONT_HERSHEY_PLAIN,3 ,(0,255,0) ,3)


            cv2.imshow("Image", img)
            cv2.waitKey(1)
            if cv2.getWindowProperty('Image',cv2.WND_PROP_VISIBLE)<1:
                print("!!!----TERMINATING-PROGRAM----!!!")
                cap.release()
                cv2.destroyAllWindows()
                break

        except KeyboardInterrupt:
            print("!!!----TERMINATING-PROGRAM----!!!")
            cap.release()
            cv2.destroyAllWindows()
            break



if __name__ == '__main__':
    main()
else:
    print("Unknown Error Occurred...")
