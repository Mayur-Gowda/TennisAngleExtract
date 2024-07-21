import os
import pandas as pd
from TennisAnalysis import compute


class CreateDF:
    def __init__(self, side, filename):
        self.views = ["FrontView", "TopView", "SideView", "NormView"]
        self.labels = ["Elbow", "Shoulder", "UpHip", "DownHip", "Knee"]
        self.values = {"left": "right", "up": "down"}
        self.values = {**self.values, **{v: k for k, v in self.values.items()}}
        self.folder = self.create_folder(filename)
        self.filename = filename
        self.force = side
        self.stability = self.values[side]
        self.FOF, self.FOS = self.createDf()

    @staticmethod
    def create_folder(foldername):
        folder_path = os.path.join(os.path.dirname(__file__), '../AnalyzedAngles/CSVFiles', foldername)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        return folder_path

    def createDf(self):
        index = pd.MultiIndex.from_product([self.views, self.labels])
        forceDataFrame = pd.DataFrame(columns=index)
        forceFrame = self.renameColumns(forceDataFrame, self.force)

        stabilityDataFrame = pd.DataFrame(columns=index)
        stabilityFrame = self.renameColumns(stabilityDataFrame, self.stability)
        return forceFrame, stabilityFrame

    def renameColumns(self, df, side):
        new_columns = []
        for col_index, (view, label) in enumerate(df.columns):
            if col_index % 5 > 1:
                new_label = f"{self.values[side]}{label}"
            else:
                new_label = f"{side}{label}"

            new_columns.append((view, new_label))

        df.columns = pd.MultiIndex.from_tuples(new_columns)
        return df

    def add_values(self, index_name, lt):
        self.FOF.loc[index_name] = lt[0]
        self.FOS.loc[index_name] = lt[1]

    def save_as_csv(self):
        force_file_path = os.path.join(self.folder, f"forceFrame.csv")
        stability_file_path = os.path.join(self.folder, f"stabilityFrame.csv")
        self.FOF.to_csv(force_file_path)
        self.FOS.to_csv(stability_file_path)
        compute.Compute("force", self.filename, self.force, force_file_path)
        compute.Compute("stable", self.filename, self.stability, stability_file_path)


