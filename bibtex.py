import os
import zipfile
import xml.etree.ElementTree as ET
import sys
import urllib.request
import shutil
from pypdf import PdfReader

def handle_arxiv_links(paper_id):
    # if the links are from arxiv, check if they have a bibtex citation available already
    try:
        file_contents = urllib.request.urlopen("https://arxiv.org/bibtex/" + paper_id).read().decode("utf-8")
        return [True, file_contents]
    except:
        # if the citation doesn't work, then just return the pdf link and get citation the normal way
        return [False, "https://arxiv.org/pdf/" + paper_id]

def handle_iacr_links(link, file_name):
    if link[-1] == "/":
        link = link[:-1]
    try:
        # get the citation straight from the website by getting the html
        os.system("curl " + link + " --output " + file_name + " > /dev/null 2>&1")
        # get the citation from the html
        with open(file_name, "r") as f:
            contents = f.read()
            contents = contents.split("<pre id=\"bibtex\">")[1].split("</pre>")[0]
            # make sure the citation is right
            if contents.startswith("\n@"):
                return [True, contents]
            else:
                return [False, link + ".pdf"]
        # delete file
        os.remove(file_name)
    except:
        return [False, link + ".pdf"]

def generate_article_bibtex(title, author, journal, volume, number, pages, year, publisher):
    key = author.split(",")[0] + number
    cite = f"""@article{{{key},
  title={{{title}}},
  author={{{author}}},
  journal={{{journal}}},
  volume={{{volume}}},
  number={{{number}}},
  pages={{{pages}}},
  year={{{year}}},
  publisher={{{publisher}}},
}}"""
    print(cite)
    return cite

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
                        link_data = link.split("/")
                        paper_id = link_data[-2] if link_data[-1] == "" else link_data[-1]
                        link = handle_arxiv_links(paper_id)
                        if link[0]:
                            citations.append(link[1])
                            print("arxiv paper", paper_id, "found and cited")
                        else:
                            links.append(link[1])
                            print("arxiv paper", paper_id, "found, awaiting citation")
                    # same thing with iacr
                    elif "iacr" in link:
                        url = link.split("/")
                        file_name = ""
                        # checks if the url is https://eprint.iacr.org/xxxx/xxx/ or https://eprint.iacr.org/xxxx/xxx
                        if url[-1] == "":
                            file_name = url[-3] + "-" + url[-2]
                            link = handle_iacr_links(link, current_path + "/" + url[-3] + "-" + url[-2] + ".html")
                        else:
                            file_name = url[-2] + "-" + url[-1]
                            link = handle_iacr_links(link, current_path + "/" + url[-2] + "-" + url[-1] + ".html")
    
                        if link[0]:
                            citations.append(link[1])
                            print("iacr paper", file_name, "found and cited")
                        else:
                            links.append(link[1])
                            print("iacr paper", file_name, "found, awaiting citation")
                    else:
                        links.append(link)
                        print("paper from", link, "found, awaiting citation")
                    
        # print the result
        print(len(citations) + len(links), "total papers found.")
        print(len(citations), "already cited,", len(links), "to download")
        if len(links) > 0:
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

            for dirpath, dirnames, filenames in os.walk(pdf_paths):
                for file in filenames:
                    print("trying to read", pdf_paths + file)
                    # check that the file is actually a pdf
                    reader = PdfReader(pdf_paths + file)
                    page = reader.pages[0]
                    text = page.extract_text()
                    with open("out.txt", "w") as f:
                        f.write(text)

        with open("output.bib", "w") as f:
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
            #shutil.rmtree(current_path + "/pdfs/")
        except:
            pass
        finally:
            print("All done")
