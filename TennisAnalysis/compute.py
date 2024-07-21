import os

import numpy as np
import pandas as pd
from .excel import sheetSetup


class Compute:
    def __init__(self, frame, filename, side, csv_file):
        self.frame = frame
        self.filename = filename
        self.side = side
        self.csv_file = csv_file
        self.compute_csv()
        self.save_to_excel()

    @staticmethod
    def add_columns_to_the_right(df, original_col, rounded_col_name, deviation_col_name, rounded_values,
                                 deviation_values):
        col_index = df.columns.get_loc(original_col)
        df.insert(col_index + 1, rounded_col_name, rounded_values)
        df.insert(col_index + 2, deviation_col_name, deviation_values)
        return df

    def compute_csv(self):
        df = pd.read_csv(self.csv_file, header=[0, 1], index_col=0)
        new_df = pd.DataFrame(index=df.index, columns=df.columns)

        for col in df.columns:
            # Get the original column values
            new_df[col] = df[col]
            rounded_column = round(df[col] / 10)
            deviation_column = rounded_column - np.mean(rounded_column)

            var_step = np.sqrt(np.sum(np.power(deviation_column, 2))/(len(deviation_column) - 1))
            deviation_column = round(rounded_column - np.mean(rounded_column), 2)

            rounded_col_name = (col[0], f'rounded_{col[1]}')
            deviation_col_name = (col[0], f'{col[1]}_deviation')

            new_df = self.add_columns_to_the_right(new_df, col, rounded_col_name, deviation_col_name,
                                                   rounded_column,
                                                   deviation_column)

        new_csv_file = self.csv_file.replace('.csv', '_processed.csv')
        new_df.to_csv(new_csv_file)
        self.csv_file = new_csv_file

    def save_to_excel(self):
        df = pd.read_csv(self.csv_file, index_col=0, header=[0, 1])
        filename = os.path.join(
                        os.path.dirname(__file__), '../AnalyzedAngles/ExcelSheets', f"{self.filename}.xlsx")
        sheetSetup(self.frame, filename, self.side, df)
