📸 Smart Attendance System using OpenCV & Excel

This project is a real-time smart attendance system built using Python and OpenCV that detects human faces through a webcam and records attendance automatically in an Excel file with date and time stamps.

The system uses Haar Cascade face detection to identify faces from a live video stream. When a face is detected, the user is prompted to enter a name once, and the attendance is marked only once per day to prevent duplicate entries. Attendance data is stored persistently in an Excel sheet using openpyxl.

🔑 Key Features

Real-time face detection using OpenCV

Webcam-based live video processing

Automatic attendance logging in Excel (.xlsx)

Date-wise and time-stamped records

Prevents duplicate attendance for the same person on the same day

Live FPS counter and timestamp overlay

User-friendly exit control (Q key)

🛠 Technologies Used

Python

OpenCV (cv2)

Haar Cascade Classifier

OpenPyXL (Excel handling)

Webcam / Computer Vision

📂 Output

attendance.xlsx file containing:

Name

Date

Time of attendance

🎯 Use Cases

Classroom attendance

Office employee tracking

Small-scale biometric attendance systems

Computer vision learning projects

This project demonstrates practical usage of computer vision, file handling, and real-time video processing, making it suitable for academic submissions, portfolios, and beginner-to-intermediate AI projects.
