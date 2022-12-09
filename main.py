import cv2
import time
import HandTrackingModule as htm
import math
import numpy as np
import pyautogui
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pygetwindow as gw

pyautogui.FAILSAFE = False

pTime = 0
cTime = 0
wCamera, hCamera = 640, 480
wScreen, hScreen = pyautogui.size()
cap = cv2.VideoCapture(0)
cap.set(3, wCamera)
cap.set(4, hCamera)
detector = htm.HandDetector(detectionCon=0.1, maxHands=1, modelComplexity=1)
px, py = -1, -1
print("screen h", hScreen, wScreen)

device = AudioUtilities.GetSpeakers()
interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
volume.SetMasterVolumeLevel(0, None)
minVol = volRange[0]
maxVol = volRange[1]


frameVal = 100


def modeToFollow():
    fingers = detector.fingersUp()
    print(fingers)
    if fingers == [0, 1, 0, 0, 0]:
        return "MOVE"
    if fingers == [1, 1, 0, 0, 0]:
        return "LEFT_CLICK"
    if fingers == [1, 1, 1, 0, 0]:
        return "VOLUME_CONTROL"
    if fingers == [0, 1, 1, 0, 0]:
        return "RIGHT_CLICK"
    if fingers == [0, 0, 0, 0, 0]:
        return "CLOSE_APP"
    return None


def volumeControl(lmList):
    x22 = lmList[12][1]
    xVol = np.interp(x22, (frameVal, wCamera - frameVal), (minVol, maxVol))
    print("XVOL", xVol)
    volume.SetMasterVolumeLevel(xVol, None)


# Stores the previous mode used as it is necessary for the Closing Appications
prevmode = None
counter = 0
prevX, prevY = -1, -1


def getDist(l1, l2):
    return math.sqrt((l1[0] - l2[0]) * (l1[0] - l2[0]) + (l1[1] - l2[1]) * (l1[1] - l2[1]))


detect = False

while True:
    success, img = cap.read()
    img = detector.findHands(img, draw=True)
    lmList, bbox = detector.findPosition(img, draw=False)

    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime

    cv2.rectangle(img, (frameVal, frameVal), (wCamera - frameVal, hCamera - frameVal),
                  (255, 0, 255), 2)
    if len(lmList) > 0:
        if not detect:
            pyautogui.sleep(1)
            detect = True
        # print(lmList[4], lmList[8])
        # print(gw.getActiveWindow().title)
        # subprocess.call("TASKKILL /F /IM firefox.exe", shell=True)

        mode = modeToFollow()

        print(mode)
        # x -> 260-360
        # y -> 400-460
        if mode == "MOVE":
            x2, y2 = lmList[8][1], lmList[8][2]  # index finger
            xT = np.interp(x2, (frameVal, wCamera - frameVal), (0, wScreen))
            yT = np.interp(y2, (frameVal, hCamera - frameVal), (0, hScreen))
            print(xT, yT, getDist([prevX, prevY], [xT, yT]))
            if (prevX == -1 and prevY == -1) or getDist([prevX, prevY], [xT, yT]) > 15:
                # print(yT)
                pyautogui.moveTo(wScreen - xT, yT, _pause=False)
            prevX, prevY = xT, yT
        elif mode == "LEFT_CLICK":
            pyautogui.leftClick()
        elif mode == "RIGHT_CLICK":
            pyautogui.rightClick()
        elif mode == "VOLUME_CONTROL":
            if prevmode == "VOLUME_CONTROL":    
                counter += 1
            else:
                counter = 1
            if counter > fps:
                volumeControl(lmList)
        elif mode == "CLOSE_APP":
            if prevmode == "CLOSE_APP":
                print("Nope!!!")
                continue
            if gw.getActiveWindow():
                gw.getActiveWindow().close()
        prevmode = mode

    else:
        detect = False

    cv2.putText(img, str(int(fps)), (10, 70), cv2.FONT_HERSHEY_PLAIN, 3,
                (255, 0, 255), 3)

    cv2.imshow("Motion", img)
    cv2.waitKey(1)