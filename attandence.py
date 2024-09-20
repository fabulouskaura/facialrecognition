from time import time
import time
import cv2
import numpy as np
import face_recognition
import os
from datetime import datetime
from win32com.client import dispatch

def speak(str1):
    speak = dispatch(("SAPI.spvoice"))
    speak.speak(str1)


path = 'C:/Users/Fabulous_Kaura/Downloads/OpencvProject/facialRecognition/images'
images = []
personName = []
myList = os.listdir(path)
print(myList)

for cu_img in myList:
    current_Img = cv2.imread(f'{path}/{cu_img}')
    images.append(current_Img)
    personName.append(os.path.splitext(cu_img)[0])
print(personName)

def faceEncodings(images):
    encodeList = []
    for img in images:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encode = face_recognition.face_encodings(img)[0]
        encodeList.append(encode)
    return encodeList


encodeListKnown = faceEncodings(images)
print("All Encodings are Completed!!!")
print(encodeListKnown)

def attendance(name):
   with open('Attendance.csv', 'r+') as f:
       myDataList = f.readlines()
       nameList = []
       for line in myDataList:
           entry = line.split(',')
           nameList.append(entry[0])

       if name not in nameList:
            time_now = datetime.now()
            tStr = time_now.strftime('%H:%M:%S')
            dStr = time_now.strftime('%D:%M:%Y')
            f.writelines(f'\n{name}, {tStr}, {dStr}')

cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    faces = cv2.resize(frame, (0,0), None, 0.25, 0.25)
    faces = cv2.cvtColor(faces, cv2.COLOR_BGR2RGB)

    facesCurrentFrame = face_recognition.face_locations(faces)
    encodesFacesCurrentFrame = face_recognition.face_encodings(faces, facesCurrentFrame)


    for encodeFace, faceLoc in zip(encodesFacesCurrentFrame, facesCurrentFrame):
        matches = face_recognition.compare_faces(encodeListKnown, encodeFace)
        facesDist = face_recognition.face_distance(encodeListKnown, encodeFace)



        matchIndex = np.argmin(facesDist)  # index of face with min distance
        matchDis = facesDist[matchIndex]  # distance of face matched

        dist_threshold = 0.45

        if matchDis <= dist_threshold:
            matches[matchIndex]
            name = personName[matchIndex].upper()
            speak('name')
            time.sleep(5)
        else:
            name = "Unknown"

        y1,x2,y2,x1 = faceLoc
        y1, x2, y2, x1 = y1*4,x2*4,y2*4,x1*4
        cv2.rectangle(frame, (x1,y1),(x2,y2), (0,255,0), 2)
        cv2.rectangle(frame, (x1, y2-35),(x2,y2),(0,255,0), cv2.FILLED)
        cv2.putText(frame, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (0,0,255), 2)
        attendance(name)
        

    cv2.imshow("Camera", frame)
    if cv2.waitKey(10) & 0xFF == ord('q'):
        break
cap.release()
cv2.destroyAllWindows()


