import os
import codecs
from bs4 import BeautifulSoup


def parseFile(filepath):
    if os.path.exists(filepath):
        content = ""
        try:
            with open(filepath, 'r') as fp:
                    encoding = 'utf-16-le'
                    with codecs.open(filepath, 'r', encoding) as fp2:
                        soup = BeautifulSoup(fp2,'lxml')
                        content += (str(soup))      
        except Exception as ex:
            print('[ERROR]--',ex)

        content = content.replace("ISO-10646-UCS-2","utf-8")
        with open(filepath,"w",encoding="utf-8") as f:
            f.write(content)
    else:
        pass

    
if __name__ == "__main__":
    filepath = './result/4625.xml'
    parseFile(filepath)






