import cv2
import numpy as np
import matplotlib.pyplot as plt
import os
import subprocess


class ROI:
    def __init__(self, input_folder):
        self.input_folder = input_folder
        self.renameFiles()
        self.output_folder = os.path.join(os.path.dirname(__file__), '../CroppedImages', f"Cropped{input_folder[15:]}")
        self.createDir()
        self.classes, self.net = self.loadYOLO()
        self.processImages()

    def createDir(self):
        """Create the output directory if it does not exist"""
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder, exist_ok=True)

    def renameFiles(self):
        """Renames the files in a folder to be incremental by using command script"""
        script = """
            $files = Get-ChildItem -Path '{0}'

            $i = 1
            foreach ($file in $files) {{
                $extension = $file.Extension
                $newName = "{{0:D3}}{{1}}" -f $i, $extension  # Format number with leading zeros
                Rename-Item -Path $file.FullName -NewName $newName
                $i++
            }}
            """.format(self.input_folder)

        # Save the PowerShell script to a temporary file
        with open('rename_files.ps1', 'w') as file:
            file.write(script)

        # Path to the PowerShell script file
        script_path = 'rename_files.ps1'

        try:
            # Run the PowerShell script using subprocess
            subprocess.run(["powershell", "-File", script_path], check=True)
        finally:
            # Delete the PowerShell script file
            if os.path.exists(script_path):
                os.remove(script_path)

        print("Files have been renamed successfully.")

    @staticmethod
    def loadYOLO():
        """Load the YOLO files"""
        # Load classes
        with open(f"{os.path.dirname(__file__)}/YOLOFiles/coco.names", 'r') as f:
            classes = [line.strip() for line in f.readlines()]

        # Load YOLO model
        net = cv2.dnn.readNet(f"{os.path.dirname(__file__)}/YOLOFiles/yolov3.weights",
                              f"{os.path.dirname(__file__)}/YOLOFiles/yolov3.cfg")
        return classes, net

    def process_image(self, input_image_path, output_image_path):
        image = plt.imread(input_image_path)
        # Convert image to the right format
        if image.dtype == np.float32:
            image = (image * 255).astype(np.uint8)

        # Check if the image has an alpha channel (4 channels)
        if image.shape[2] == 4:
            image = cv2.cvtColor(image, cv2.COLOR_RGBA2RGB)

        self.net.setInput(cv2.dnn.blobFromImage(image, 0.00392, (416, 416), (0, 0, 0), swapRB=True, crop=False))
        layer_names = self.net.getLayerNames()
        output_layers = [layer_names[i - 1] for i in self.net.getUnconnectedOutLayers()]
        outs = self.net.forward(output_layers)

        # Analyze detections
        class_ids = []
        confidences = []
        boxes = []
        Width = image.shape[1]
        Height = image.shape[0]
        for out in outs:
            for detection in out:
                scores = detection[5:]
                class_id = np.argmax(scores)
                confidence = scores[class_id]
                if confidence > 0.5:  # Increased confidence threshold
                    center_x = int(detection[0] * Width)
                    center_y = int(detection[1] * Height)
                    w = int(detection[2] * Width)
                    h = int(detection[3] * Height)
                    x = center_x - w // 2  # Ensure integer indices
                    y = center_y - h // 2  # Ensure integer indices
                    class_ids.append(class_id)
                    confidences.append(float(confidence))
                    boxes.append([x, y, w, h])

        indices = cv2.dnn.NMSBoxes(boxes, confidences, 0.5, 0.4)  # Adjusted NMS threshold

        # Select the best bounding box for the player based on size, aspect ratio, and proximity to center
        best_box = None
        largest_area = 0
        center_x, center_y = Width // 2, Height // 2

        if len(indices) > 0:
            for i in indices.flatten():
                if class_ids[i] == 0:  # Check if the detected class is 'person'
                    box = boxes[i]
                    x, y, w, h = box
                    area = w * h
                    aspect_ratio = h / w

                    # Filtering conditions
                    if aspect_ratio > 1.2 and 0.5 * Width * Height > area > 0.02 * Width * Height:
                        box_center_x = x + w // 2
                        box_center_y = y + h // 2
                        distance_to_center = np.sqrt((box_center_x - center_x) ** 2 + (box_center_y - center_y) ** 2)
                        if area > largest_area and distance_to_center < max(Width, Height) / 2:
                            largest_area = area
                            best_box = box

        # Crop and display the best bounding box
        if best_box is not None:
            x, y, w, h = best_box
            x, y = max(0, x), max(0, y)  # Ensure x and y are not negative

            # Increase width and height by 50% from all sides
            w_increase = int(w * 0.35)
            h_increase = int(h * 0.35)

            # Adjust the coordinates to maintain the center position
            x = max(0, x - w_increase // 2)
            y = max(0, y - h_increase // 2)
            w += w_increase
            h += h_increase

            # Ensure that the new bounding box remains within the boundaries of the original image
            w = min(w, Width - x)
            h = min(h, Height - y)

            person_image = image[y:y + h, x:x + w]

            # # Display the image with the bounding box
            # fig, ax = plt.subplots(1)
            # ax.imshow(image)
            # rect = plt.Rectangle((x, y), w, h, linewidth=2, edgecolor='r', facecolor='none')
            # ax.add_patch(rect)
            # plt.axis('off')  # Hide axis
            # plt.show()

            # Save the cropped image
            if person_image.shape[0] > 0 and person_image.shape[1] > 0:  # Ensure the cropped image is valid
                cv2.imwrite(output_image_path, cv2.cvtColor(person_image, cv2.COLOR_RGB2BGR))

    def processImages(self):
        """Process all the images in a provided folder"""
        print("Processing Images")
        for filename in os.listdir(self.input_folder):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                input_image_path = os.path.join(self.input_folder, filename)
                output_image_path = os.path.join(self.output_folder, filename)
                self.process_image(input_image_path, output_image_path)

        print("Processing complete.")
