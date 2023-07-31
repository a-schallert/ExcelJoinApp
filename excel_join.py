import pandas as pd

class ExcelJoin():

    def __init__(self, file_1, file_2):
        self.file_1 = pd.read_excel(file_1)
        self.file_2 = pd.read_excel(file_2)

    def __call__(self, on, how=['inner', 'left', 'right', 'outer']):
        join = pd.merge(self.file_1, self.file_2, on=on, how=how)
        return join

    def index(self):
        df1_index = self.file_1.columns.to_list()
        df2_index = self.file_2.columns.to_list()
        combined_index = list(set(df1_index + df2_index))
        return combined_index
 
    def intersection(self):
        df1_index = self.file_1.columns.to_list()
        df2_index = self.file_2.columns.to_list()
        intersected_list = list(set(df1_index).intersection(df2_index))
        return intersected_list
