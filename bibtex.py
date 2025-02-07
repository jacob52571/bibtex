import os
import zipfile
import xml.etree.ElementTree as ET
import sys
import urllib.request
import shutil

def handle_arxiv_links(link):
    link_data = link.split("/")
    paper_id = link_data[-2] if link_data[-1] == "" else link_data[-1]
    # if the links are from arxiv, check if they have a bibtex citation available already
    try:
        file_contents = urllib.request.urlopen("https://arxiv.org/bibtex/" + paper_id).read().decode("utf-8")
        return [True, file_contents]
    except:
        
        return [False, "https://arxiv.org/pdf/" + paper_id]


    # if the links are from arxiv, i want the pdf link only
    

if __name__ == "__main__":
    current_path = os.path.abspath(os.path.dirname(__file__))

    # get the word file
    link_file = current_path + "/" + input("Enter the name/path of the word document: ")

    links = []
    citations = []
    try:
        # format the file 
        os.rename(link_file, link_file + ".zip")
        with zipfile.ZipFile(link_file + ".zip", 'r') as zip_ref:
            print("extracting to " + current_path + "/word_file/")
            zip_ref.extractall(current_path + "/word_file/")

        # get links from the text in the file
        with open(current_path + "/word_file/word/_rels/document.xml.rels", "r") as file:
            tree = ET.parse(file)
            relationships = tree.getroot()
            for relationship in relationships:
                if "TargetMode" in relationship.attrib and relationship.attrib["TargetMode"] == "External" and relationship.attrib["Target"][0:4] == "http":
                    link = relationship.attrib["Target"]
                    # if the link is arxiv, get the bibtex automatically if possible, otherwise add the pdf link
                    if "arxiv" in link:
                        link = handle_arxiv_links(link)
                        if link[0]:
                            citations.append(link[1])
                        else:
                            links.append(link[1])
                    else:
                        links.append(link)
                    
        # print the result
        print(len(links), "links found!")

        # make the folder for the pdfs
        pdf_paths = current_path + "/pdfs/"
        if not os.path.exists(pdf_paths):
            os.makedirs(pdf_paths)

        # get the files from the links
        for i in range(0, len(links)):
            link = links[i]
            
            # try to get the file from urllib and using curl if that fails
            file_contents = ""
            try:
                file_contents = urllib.request.urlopen(link).read()
                with open(pdf_paths + i + ".pdf", "w") as file:
                    file.write(file_contents)
            except:
                print("urllib failed, using curl...")
                os.system("curl " + link + " --output " + pdf_paths + str(i) + ".pdf")
        
        print("All pdfs downloaded, starting processing now...")

        for file in os.walk(pdf_paths):
            # check that the file is actually a pdf
            with open(file, "r") as f:
                pass

        with open("citations.txt", "w") as f:
            for line in citations:
                f.write(line + "\n")
    except FileNotFoundError:
        # error handling if file name was wrong
        print("File was not found.")
        sys.exit(0)
    finally:
        # the doc will always get changed back into .docx if the script fails
        try:
            # make the zip back into word
            os.rename(link_file + ".zip", link_file)
            # remove the folder
            shutil.rmtree(current_path + "/word_file/")
        finally:
            pass
