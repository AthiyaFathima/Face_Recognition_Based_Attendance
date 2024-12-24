from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1) 

video = cv2.VideoCapture(0)
facesdetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as f:
    LABELS= pickle.load(f)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES,LABELS)

imgBackground = cv2.imread("background.png")
if imgBackground is None:
    print("Error loading background image.")
else:
    print("Background image loaded successfully.")

COL_NAMES = ['NAME', 'TIME']

# Variable to track if attendance was taken
attendance_taken = False

while True:
    ret,frame = video.read()
    if ret:
        imgBackground_resized = cv2.resize(imgBackground, (frame.shape[1], frame.shape[0]), interpolation=cv2.INTER_CUBIC)
    if not ret:
        break

    # Blur the background image
    blurred_background = cv2.GaussianBlur(imgBackground_resized, (21, 21), 0)

    gray = cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
    faces = facesdetect.detectMultiScale(gray,1.3, 5)

    # Resize the background to match the frame size
    imgBackground_resized = cv2.resize(imgBackground, (frame.shape[1], frame.shape[0]), interpolation=cv2.INTER_CUBIC)

    for (x,y,w,h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img,(50,50)).flatten().reshape(1,-1)
        output  = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H-%M-%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")

        # Get the sharp face region from the original frame
        face_region = frame[y:y+h, x:x+w]

        # Place the sharp face region back on the blurred background
        blurred_background[y:y+h, x:x+w] = face_region

        # Draw rectangles around the face in the frame
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)

        # --- Drawing the red box for the name on the frame ---

        # Draw a solid red rectangle as the background for the name
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (0, 0, 255), -1)  # Solid red rectangle

        # Write the name in white on top of the red box
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)  # White text

        # --- Drawing the red box for the name on the blurred background ---

        # Draw a solid red rectangle behind the name on the blurred background
        cv2.rectangle(blurred_background, (x, y - 40), (x + w, y), (0, 0, 255), -1)  # Solid red rectangle

        # Write the name in white on top of the red box in the blurred background
        cv2.putText(blurred_background, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)  # White text

    # cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)

        attendance = [str(output[0]), str(timestamp)]

    # Display "Press 'o' to Take Attendance" message
    text = "Press 'o' to Mark an Attendance"
    text_size = cv2.getTextSize(text, cv2.FONT_HERSHEY_COMPLEX, 1, 2)[0]

    # Calculate the position for the rectangle
    text_x = 50
    text_y = 50

    # Draw a solid red rectangle behind the text
    cv2.rectangle(blurred_background, (text_x, text_y - text_size[1] - 10), (text_x + text_size[0], text_y + 10), (0, 0, 255), -1)  # Solid red rectangle

    # Write the text in white on top of the red rectangle
    cv2.putText(blurred_background, text, (text_x, text_y), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)  # White text

    # Step 5: Display the final frame with a blurred background and sharp face(s)
    cv2.imshow("Frame", blurred_background)
    
    # Convert the background image to color if necessary
    if len(imgBackground.shape) != 3 or imgBackground.shape[2] != 3:
        imgBackground = cv2.cvtColor(imgBackground, cv2.COLOR_GRAY2BGR)
    k = cv2.waitKey(1)
    if k == ord('o'):
        speak("Attendance Taken...")
        attendance_taken = True
        attendance_message_start_time = time.time()  # Record the time when the message appears
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
            
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()
    if k == ord('q'):
        break
video.release()
cv2.destroyAllWindows()