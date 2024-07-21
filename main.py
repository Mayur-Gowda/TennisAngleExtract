"""The code asks for an input whether to pre-process images or not. If processed, accuracy improves significantly
but takes a bit of time to complete. If not done, the accuracy is not good. It asks for another input to specify
which handed the player is."""

import os
import cv2
import TennisAnalysis as Ta
import mediapipe as mp
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='google.protobuf.symbol_database')


input_folder = "Tennis Dataset/Serve Dataset/Swiatek-R"
output_folder = f"CroppedImages/{input_folder[15:]}"


side = input("The player is which handed?").lower()

if input("Would you like to pre-process the images? Y/N: \n").lower() == 'y':
    roi = Ta.ROI(input_folder)
    output_folder = roi.output_folder


mp_pose = mp.solutions.pose
pose = mp_pose.Pose(static_image_mode=True, enable_segmentation=False, min_detection_confidence=0.8, model_complexity=2)
mp_drawing = mp.solutions.drawing_utils

image_files = sorted(os.listdir(output_folder))

angFunc = Ta.Angle(filename=f"{input_folder[15:]}", side=side)

for image_file in image_files:
    # Check if the file is an image (assuming all files in the folder are images)
    posList = []
    if image_file.endswith(('.jpg', '.jpeg', '.png')):
        # Read the image
        image_path = os.path.join(output_folder, image_file)
        img = cv2.imread(image_path)

        # Process the image
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

        results = pose.process(img_rgb)

        # Apply background subtraction
        backSub = cv2.createBackgroundSubtractorMOG2()
        fg_mask = backSub.apply(img_rgb)

        # Find contours and crop the image to the largest contour
        contours, _ = cv2.findContours(fg_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if contours:
            largest_contour = max(contours, key=cv2.contourArea)
            x, y, w, h = cv2.boundingRect(largest_contour)
            cropped_image = img_rgb[y:y + h, x:x + w]

            # Process the cropped image with MediaPipe Pose
            results = pose.process(cropped_image)

            # Draw the pose annotation on the original image
            if results.pose_landmarks:
                # Convert landmarks to original image coordinates
                for lm in results.pose_landmarks.landmark:
                    cx, cy, cz = int(lm.x * w) + x, int(lm.y * h) + y, int(lm.z * w)
                    posList.extend([cx, cy, cz])
                    # cv2.circle()

        try:
            angFunc.calculate(posList)
            angFunc.save_values(image_file)
            # angFunc.values()
            # angFunc.saveToExcel(image_file)
            # plt = angFunc.createPlot()
            # plt.show()

        except IndexError:
            continue

        # Display the processed image (optional)
        mp_drawing.draw_landmarks(img_rgb, results.pose_landmarks, mp_pose.POSE_CONNECTIONS)
        imS = cv2.resize(img_rgb, (0, 0), fx=0.4, fy=0.4)

        cv2.imshow("Image", imS)
        key = cv2.waitKey(1)
        # if key == ord('q'):
        #     continue
        # elif key == ord('s'):
        #     angFunc.save_values(image_file)

angFunc.save_files()

# Release resources and close windows
cv2.destroyAllWindows()
