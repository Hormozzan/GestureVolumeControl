from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from math import hypot
from numpy import interp
import cv2
import time
import HandTrackingModule as htm

### Initializing webCam object
cap = cv2.VideoCapture(0)

### Instantiating from handDetector class with some detection confidence (detectionCon)
### Change detectionCon to change model's strictness of hand detection
detector = htm.handDetector(detectionCon=0.8)

### Identifying audio interface
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

### Finding volume range
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]

### Initializing some variables
volBar = 400
volPer = 0
pTime = 0

while True:
    ### Reading from the webCam object
    success, img = cap.read()

    ### Detecting landmarks of hand
    img = detector.findHands(img)

    ### Getting coordinates of the landmarks
    lmList = detector.findPosition(img, draw=False)

    if len(lmList) != 0:
        # print(lmList[4], lmList[8])

        ### Extracting coordinates of the thumb, the index finger, and the middle point
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

        ### Drawing a circle on each of these points, drawing a line between the thumb and the index finger
        cv2.circle(img, (x1, y1), 10, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 10, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (cx, cy), 10, (255, 0, 255), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)

        ### Finding the length of this line
        length = hypot(x2 - x1, y2 - y1)
        # print(length)

        ### Mapping the length of this line onto the volume range,
        ### onto the location of volume bar, and onto volume percentage
        ### 20 is the length of the connecting line, when the thumb and the index finger are close to each other
        ### 180 is the length of the connecting line, when the thumb and the index finger are apart from each other
        ### 20 and 180 vary with respect to the distance of hand from the webCam
        vol = interp(length, [20, 180], [minVol, maxVol])
        volBar = interp(length, [20, 180], [400, 150])
        volPer = interp(length, [20, 180], [0, 100])
        # print(int(length), vol)

        ### Setting the mapped volume to the system volume level
        volume.SetMasterVolumeLevel(vol, None)

        ### Drawing a green circle on the point which the thumb and the index finger meet
        if length < 20:
            cv2.circle(img, (cx, cy), 10, (0, 255, 0), cv2.FILLED)

    ### Drawing volume bar, writing volume percentage
    cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
    cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
    cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 2)

    ### Computing and writing frame-per-second
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (20, 50), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 2)

    ### Showing the window
    cv2.imshow("Gesture Volume Control", img)
    cv2.waitKey(1)
