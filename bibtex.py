import os
import zipfile

current_path = os.path.abspath(os.path.dirname(__file__))

# get the word file
link_file = current_path + "/" + input("Enter the name/path of the word document: ")

links = []

# format the file 
os.rename(link_file, link_file + ".zip")
with zipfile.ZipFile(link_file + ".zip", 'r') as zip_ref:
    print("extracting to " + current_path + "/word_file/")
    zip_ref.extractall(current_path + "/word_file/")

# get links from the text in the file
with open(current_path + "/word_file/word/document.xml", "r") as file:
    for line in file:
        if "http" in line:
            links.append(line)

print(links)

# make the zip back into word
os.rename(link_file + ".zip", link_file)
