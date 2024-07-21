import numpy as np
import math
import matplotlib.pyplot as plt
from TennisAnalysis import excel, orgData

'''Note: In this z-axis is the FrontView (Blue), 
the x-axis is the SideView (Red) and the y-axis is the TopView (Green)'''


class Angle:
    """Calculates angles in 3 dimensions for the specified joints."""

    def __init__(self, **kwargs):
        self.angles = [[11, 13, 15], [13, 11, 23], [11, 24, 23], [23, 24, 26], [24, 26, 28],
                       [12, 14, 16], [14, 12, 24], [12, 23, 24], [24, 23, 25], [23, 25, 27],
                       ]
        self.filename = kwargs['filename']
        self.side = kwargs['side'].lower()
        self.df = orgData.CreateDF(self.side, self.filename)
        self.norm = self.x_flat = self.y_flat = self.z_flat = None
        self.xFlatAngle = self.yFlatAngle = self.zFlatAngle = self.normAngle = None

    @staticmethod
    def calculate_angle(point1, point2, point3):  # internal Usage
        """Function to calculate angle between three points"""

        vector1 = np.array(point1) - np.array(point2)  # Calculate vectors
        vector2 = np.array(point3) - np.array(point2)  # ||
        magnitude1 = np.linalg.norm(vector1)  # Calculate magnitudes
        magnitude2 = np.linalg.norm(vector2)  # ||
        v1_norm = np.divide(vector1, magnitude1)  # Normalized vectors
        v2_norm = np.divide(vector2, magnitude2)  # ||
        dot_product = np.dot(v1_norm, v2_norm)  # Calculate dot product
        angle_rad = math.acos(dot_product)  # Calculate angle in radians
        angle_deg = math.degrees(angle_rad)  # Convert angle to degrees
        return angle_deg

    @staticmethod
    def maxMagnitude(coordList):  # internal Usage
        """Returns the max value from a list coordinates"""

        return max(max(coordList[i:i + 3]) for i in range(0, len(coordList), 3))

    def calculate(self, coordList):
        self.norm, self.x_flat, self.y_flat, self.z_flat = [], [], [], []
        self.xFlatAngle, self.yFlatAngle, self.zFlatAngle, self.normAngle = {}, {}, {}, {}
        self.flatValues(coordList)
        self.angleValues()

    def flatValues(self, coordList):  # internal Usage
        """Gives 4 lists where in 3 of them one of the coords is kept constant"""

        # Parse coordinates from lines
        for i in range(0, len(coordList), 3):
            x, y, z = coordList[i:i + 3]
            c = 1.2 * self.maxMagnitude(coordList)
            self.norm.append([x / 10, y / 10, z / 10])
            self.x_flat.append([c / 10, y / 10, z / 10])
            self.y_flat.append([x / 10, c / 10, z / 10])
            self.z_flat.append([x / 10, y / 10, c / 10])

    def drawLines(self, ax):  # internal Usage
        """Used for visualization in the plots by connecting the joints"""

        colorList = ["midnightblue", "darkgreen", "darkred", "olive", "teal",
                     "deepskyblue", "springgreen", "coral", "gold", "steelblue"]

        for points, color in zip(self.angles, colorList):
            for i in range(len(points) - 1):
                ax.plot([self.z_flat[points[i]][0], self.z_flat[points[i+1]][0]],
                        [self.z_flat[points[i]][1], self.z_flat[points[i+1]][1]],
                        [self.z_flat[points[i]][2], self.z_flat[points[i+1]][2]], color=color)

    def createPlot(self):
        """Returns a plot by placing the lists in 3D space"""

        fig = plt.figure()
        ax = fig.add_subplot(111, projection='3d')
        ax.scatter([point[0] for point in self.x_flat],
                   [point[1] for point in self.x_flat],
                   [point[2] for point in self.x_flat], c='r', marker='o')
        ax.scatter([point[0] for point in self.y_flat],
                   [point[1] for point in self.y_flat],
                   [point[2] for point in self.y_flat], c='g', marker='o')
        ax.scatter([point[0] for point in self.z_flat],
                   [point[1] for point in self.z_flat],
                   [point[2] for point in self.z_flat], c='b', marker='o')
        ax.scatter([point[0] for point in self.norm],
                   [point[1] for point in self.norm],
                   [point[2] for point in self.norm], c='k', marker='o')

        # ax.invert_xaxis()
        # Set labels and title
        ax.set_xlabel('X Label')
        ax.set_ylabel('Y Label')
        ax.set_zlabel('Z Label')
        ax.set_title('3D Scatter Plot')

        self.drawLines(ax)

        return plt

    def angleValues(self):  # internal Usage
        """Stores the angles in a dictionary"""

        joints = ["LeftElbow", "LeftShoulder", "UpRightHip", "DownLeftHip", "LeftKnee",
                  "RightElbow", "RightShoulder", "UpLeftHip", "DownRightHip", "RightKnee"]

        for angle, joint in zip(self.angles, joints):
            self.xFlatAngle[joint] = round(self.calculate_angle(self.x_flat[angle[0]],
                                                                self.x_flat[angle[1]],
                                                                self.x_flat[angle[2]]), 3)
            self.yFlatAngle[joint] = round(self.calculate_angle(self.y_flat[angle[0]],
                                                                self.y_flat[angle[1]],
                                                                self.y_flat[angle[2]]), 3)
            self.zFlatAngle[joint] = round(self.calculate_angle(self.z_flat[angle[0]],
                                                                self.z_flat[angle[1]],
                                                                self.z_flat[angle[2]]), 3)
            self.normAngle[joint] = round(self.calculate_angle(self.norm[angle[0]],
                                                               self.norm[angle[1]],
                                                               self.norm[angle[2]]), 3)

    def values(self):
        """Prints the angles"""

        print("\n(Blue)FrontView Angles:")
        [print(f"{key}: {value}") for key, value in self.zFlatAngle.items()]
        print("\n(Green)TopView Angles:")
        [print(f"{key}: {value}") for key, value in self.yFlatAngle.items()]
        print("\n(Red)SideView Angles:")
        [print(f"{key}: {value}") for key, value in self.xFlatAngle.items()]
        print("\n(Black)NormAngles:")
        [print(f"{key}: {value}") for key, value in self.normAngle.items()]

    def save_values(self, pic):
        """Uses the excel.py file to save the angle inside the Excel sheet."""

        s1Vals = []
        s2Vals = []

        for subList in [self.zFlatAngle.values(), self.yFlatAngle.values(),
                        self.xFlatAngle.values(), self.normAngle.values()]:
            s1Vals.extend(list(subList)[:5])
            s2Vals.extend(list(subList)[5:10])
        if self.side == "left":
            self.df.add_values(pic, [s1Vals, s2Vals])
        else:
            self.df.add_values(pic, [s2Vals, s1Vals])

    def save_files(self):
        wb = excel.ExcelSave(self.filename, self.side)
        wb.addValues(0, self.df.FOF)
        wb.addValues(1, self.df.FOS)
        wb.workbook.close()
        self.df.save_as_csv()
