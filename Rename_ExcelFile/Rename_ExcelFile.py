import os

path = os.getcwd() # current working directory

for filename in os.listdir(path):
    if filename.startswith("PM_IG27_"):
        new_filename = filename.replace("PM_IG27_", "NewPMDataQueryResult_")
        new_filename = new_filename.replace("_01.csv", ".csv")
        new_filename = new_filename.replace("NewPMDataQueryResult_15_", "NewPMDataQueryResult_")
        os.rename(os.path.join(path, filename), os.path.join(path, new_filename))
