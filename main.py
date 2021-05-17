from pathlib import Path
import modules
import os

file_path = Path(r"c:\Users\ehom\Documents\IdeaProjects\Python\Projects\trackChangesQuote\sample\\")

for file in file_path.iterdir():
    tcq_folder = str(file.parent)+"\\"+file.name+"_tcq"
    os.mkdir(tcq_folder)
    acceptPath = modules.accept_all(str(file), tcq_folder)
    rejectPath = modules.reject_all(str(file), tcq_folder)
    modules.convert_to_txlf(tcq_folder, "en-us")
    modules.segment_and_pseudo(tcq_folder)
    tm_path = tcq_folder+"\\tm"
    os.mkdir(tm_path)
    modules.create_tm_and_update(tm_path, rejectPath)
    modules.analyze_accepted(tm_path, acceptPath, tcq_folder)


