import os
import zipfile
import xml.etree.ElementTree as ET
import sys
import urllib.request

def handle_arxiv_links(link):
    # if the links are from arxiv, i want the pdf link only
    # link format: https://arxiv.org/abs/1604.07316
    link_data = link.split("/")
    paper_id = link_data[-2] if link_data[-1] == "" else link_data[-1]
    return "https://arxiv.org/pdf/" + paper_id

if __name__ == "__main__":
    current_path = os.path.abspath(os.path.dirname(__file__))

    # get the word file
    link_file = current_path + "/" + input("Enter the name/path of the word document: ")

    links = []
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
                    links.append(relationship.attrib["Target"])

        # print the result
        print(len(links), "links found!")

        # make the folder for the pdfs
        pdf_paths = current_path + "/pdfs/"
        if not os.path.exists(pdf_paths):
            os.makedirs(pdf_paths)

        # get the files from the links
        for i in range(0, len(links)):
            link = links[i]

            # if the link is arxiv, ensure it is the pdf link
            if "arxiv" in link:
                link = handle_arxiv_links(link)
            
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

    except FileNotFoundError:
        # error handling if file name was wrong
        print("File was not found.")
        sys.exit(0)
    finally:
        # the doc will always get changed back into .docx if the script fails
        try:
            # make the zip back into word
            os.rename(link_file + ".zip", link_file)
        finally:
            pass
