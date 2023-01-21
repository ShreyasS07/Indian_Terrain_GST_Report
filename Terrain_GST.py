import os
import openpyxl
import pandas as pd
from tkinter import *
from tkinter import filedialog

def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    print("Selected folder path is : ", folder_path)
    # files = os.listdir(folder_path)

    # Get list of files in selected folder
    file_list = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    print('Files inside the selected folder is:')
    for file in file_list:
        print(file)

    # Create dataframes for each CSV and XLSX file
    for file in file_list:
        if file.endswith(".csv"):
            df = pd.read_csv(os.path.join(folder_path, file))
            df_name = file.replace(".csv", "") + "_df"
            globals()[df_name] = df
            print(df_name, "shape:", df.shape)
        elif file.endswith(".xlsx"):
            df = pd.read_excel(os.path.join(folder_path, file))
            df_name = file.replace(".xlsx", "") + "_df"
            globals()[df_name] = df
            print(df_name, "shape:", df.shape)

    # Combine all dataframes into a single dataframe
    final_df = pd.concat([v for k, v in globals().items() if k.endswith("_df")], ignore_index=True)

    print("final_df shape is :", final_df.shape)
    print("Column Names:", final_df.columns)
    print("\nData Count:")
    print(final_df.count())
    print("Final dataframe contains null values:", final_df.isnull().values.any())
    print("Final dataframe contains duplicate data:", final_df.duplicated().any())

    #Saving the final Dataframe into Excel file in the same folder

    final_df.to_excel(os.path.join(folder_path, "Final report.xlsx"), index=False)
    print("All Excel files extracted & saved to single excel file & stored in same folder")
    print("Final report saved at:", os.path.join(folder_path, "Final report.xlsx"))


window = Tk()
window.title(" Mindful Automation Pvt Ltd ")
window.geometry('600x150')
window.configure(background="white")
label_file_explorer = Label(window, text=" Indian Terrain GST Report Automation ",
                            width=100, height=3, fg="blue")
file = Button(window,
                  text=" Select the folder ", command=select_folder)
label_file_explorer.grid(column=1, row=1)
file.grid(column=1, row=2)
window.mainloop()
