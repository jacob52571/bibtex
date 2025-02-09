import os
import zipfile
import xml.etree.ElementTree as ET
import sys
import urllib.request
import shutil
from pypdf import PdfReader
import re

def handle_arxiv_links(paper_id, pdf_paths, auto_keywords=False):
    # if the links are from arxiv, check if they have a bibtex citation available already
    try:
        file_contents = urllib.request.urlopen("https://arxiv.org/bibtex/" + paper_id).read().decode("utf-8")
        # add doi to the citation
        doi = f"10.48550/arXiv.{paper_id}"
        try:
            # get page count and try to get keywords
            os.system("curl https://arxiv.org/pdf/" + paper_id + " --output " + pdf_paths + paper_id + ".pdf > /dev/null 2>&1")
            reader = PdfReader(pdf_paths + paper_id + ".pdf")
            # get page count
            page_count = len(reader.pages)
            keywords = ""
            if auto_keywords:
                # read the first page and see if there are any keywords
                page = reader.pages[0]
                text = page.extract_text()
                # format the pdf
                text = " ".join(text.split("\n"))
                # check if there are keywords
                if "keywords" in text.lower():
                    # find the start and end of the keywords
                    loc_of_terms = text.lower().find("keywords") + 8
                    end_of_terms = min(text.lower().find("introduction"), text.lower()[loc_of_terms:].find("1") + loc_of_terms, key=lambda v: v if v > 0 else float('inf'))
                    # gather all the keywords by finding the first letter (since the string may start with ": ")
                    citation_start = lambda s: next((i for i, c in enumerate(s) if c.isalpha()), None)
                    # match regex to find where the keywords are split
                    keywords_list = re.split(r'[^\w\s]+', text[loc_of_terms:end_of_terms])
                    # join them into a string
                    keywords = ";".join([x.strip() for x in keywords_list if len(x) > 2])
                elif "index terms" in text.lower():
                    # the start and end of the keywords
                    loc_of_terms = text.lower().find("index terms") + 11
                    end_of_terms = text.lower().find(". i.")
                    # find the first letter
                    citation_start = lambda s: next((i for i, c in enumerate(s) if c.isalpha()), None)
                    # get the keywords and format them
                    keywords = text[loc_of_terms:end_of_terms][citation_start(text[loc_of_terms:end_of_terms]):].replace("- ", "").replace(", ", ";")
        except Exception as e:
            # if there's an error, just add the doi
            print(f'Error: {e}')
            file_contents = file_contents.replace("}, \n}", f"}},\n      doi={{{doi}}},\n}}")
            return [True, file_contents]
        else:
            # if everything works, then delete the pdf and add the page count and keywords
            os.remove(pdf_paths + paper_id + ".pdf")
            file_contents = file_contents.replace("}, \n}", f"}},\n      doi={{{doi}}},\n      pages={{1-{page_count}}}{f",\n      keywords={{{keywords}}}" if len(keywords) > 0 else ""}\n}}")
            return [True, file_contents]
    except:
        # if the citation doesn't work, then just return the pdf link and get citation the normal way
        return [False, "https://arxiv.org/pdf/" + paper_id]

def handle_iacr_links(link, file_name):
    # gets rid of a / if there's one at the end of the url
    if link[-1] == "/":
        link = link[:-1]
    try:
        # get the citation straight from the website by getting the html
        os.system("curl " + link + " --output " + file_name + " > /dev/null 2>&1")
        # get the citation from the html
        contents = ""
        keywords = ""
        with open(file_name, "r") as f:
            page_data = f.read()
            contents = page_data.split("<pre id=\"bibtex\">\n")[1].split("</pre>")[0]
            # get the keywords
            keywords_list = page_data.split(" class=\"me-2 badge bg-secondary keyword\">")[1:]

            if keywords_list:
                keywords = ";".join([kw.split("</a>")[0] for kw in keywords_list])
        # delete file
        os.remove(file_name)
        # make sure the citation is right
        if contents.startswith("@"):
            return [True, contents.replace("}\n}", f"}},{f",\n      keywords={{{keywords}}}" if len(keywords) > 0 else ""}\n}}")]
        else:
            return [False, link + ".pdf"]
    except Exception as e:
        print(f'Error: {e}')
        # if it doesnt work then return pdf link
        return [False, link + ".pdf"]

if __name__ == "__main__":
    current_path = os.path.abspath(os.path.dirname(__file__))
    auto_cite_arxiv = None

    # get the word file
    link_file = current_path + "/" + input("Enter the name/path of the word document: ")

    links = []
    citations = []
    try:
        # format the file 
        os.rename(link_file, link_file + ".zip")
        with zipfile.ZipFile(link_file + ".zip", 'r') as zip_ref:
            zip_ref.extractall(current_path + "/word_file/")

        # make the folder for the pdfs
        pdf_paths = current_path + "/pdfs/"
        if not os.path.exists(pdf_paths):
            os.makedirs(pdf_paths)
        
        # get links from the text in the file
        with open(current_path + "/word_file/word/_rels/document.xml.rels", "r") as file:
            # parse the xml file that has all the hyperlinks
            tree = ET.parse(file)
            relationships = tree.getroot()
            for relationship in relationships:
                if "TargetMode" in relationship.attrib and relationship.attrib["TargetMode"] == "External" and relationship.attrib["Target"][0:4] == "http":
                    link = relationship.attrib["Target"]
                    # if the link is arxiv, get the bibtex automatically if possible, otherwise add the pdf link
                    if "arxiv" in link:
                        if auto_cite_arxiv == None:
                            auto_cite_arxiv = input("Do you want to automatically get keywords for arxiv papers? (y/n): ").lower() == "y"
                        link_data = link.split("/")
                        paper_id = link_data[-2] if link_data[-1] == "" else link_data[-1]
                        link = handle_arxiv_links(paper_id, pdf_paths, auto_keywords=auto_cite_arxiv)
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
                    # for general pdfs, add them to the list and cite them the normal way
                    else:
                        links.append(link)
                        print("paper from", link, "found, awaiting citation")
                    
        # print the result
        print(len(citations) + len(links), "total papers found.")
        print(len(citations), "already cited,", len(links), "to download")
        if len(links) > 0:
            print("The following papers will need manual citation:")
            for link in links:
                print(link)
            # this code doesn't work yet, but it's supposed to download the pdfs and extract the data
            #"""
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
            
            # reads all pdfs and extracts the data
            for dirpath, dirnames, filenames in os.walk(pdf_paths):
                for file in filenames:
                    print("trying to read", pdf_paths + file)
                    # check that the file is actually a pdf
                    reader = PdfReader(pdf_paths + file)
                    page = reader.pages[0]
                    text = page.extract_text()
                    with open("out.txt", "w") as f:
                        f.write(text)
            #"""

        # write the cites to the output file
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
            shutil.rmtree(current_path + "/pdfs/")
        except:
            pass
        finally:
            print("All done")
