import os
from optparse import OptionParser
from change_file import parseFile

from threading import Thread
import logID






def show(list,path,id,filter):
    if filter == "":
        if id == "":
            pass
        else:
            filter = f" where EventID={id}"
    else:
        filter = f" where {filter}"

    os.system(f'logparser.exe "select {list} from {path}{filter}" -i:EVT -o:DATAGRID')





def save(list,path,id,filter):
    if filter == "":
        if id == "":
            pass
        else:
            filter = f" where EventID={id}"
    else:
        filter = f" where {filter}"

    os.system(f'logparser.exe "select {list} into ./result/{id}.xml from {path}{filter}" -i:EVT -o:xml')



def task(id,file_name):
    save(list=options.list , path=f"{options.path}/{file_name}" , id=id , filter="")



process = []

def default_id():
    global process
    for id in ["4624","4625","1102","4720","4726","4728","4732","4756"]:
        id = Thread(target=task,args=(id,"Security.evtx"))
        id.start()
        process.append(id)


    for id in ["104","6004","6005"]:
        id = Thread(target=task,args=(id,"System.evtx"))
        id.start()
        process.append(id)


    for id in ["1149"]:
        id = Thread(target=task,args=(id,"Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx"))
        id.start()
        process.append(id)

    for id in ["25"]:
        id = Thread(target=task,args=(id,"Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx"))
        id.start()
        process.append(id)


    for p in process:
        p.join()


    for id in ["4624","4625","1102","4720","4726","4728","4732","4756","104","6004","6005","1149","25"]:
        file = f'./result/{id}.xml'
        parseFile(file)




def logo():

    logAnalyse = r'''
    __                ___                __
   / /   ____  ____ _/   |  ____  ____ _/ /_  __________
  / /   / __ \/ __ `/ /| | / __ \/ __ `/ / / / / ___/ _ \
 / /___/ /_/ / /_/ / ___ |/ / / / /_/ / / /_/ (__  )  __/
/_____/\____/\__, /_/  |_/_/ /_/\__,_/_/\__, /____/\___/        --by 0V0h
            /____/                     /____/

    '''
    print(logAnalyse)
    print("\n日志文件默认存放在log目录下，输出默认存放在result目录\n\n")






if __name__ == "__main__":

    logo()

    parser = OptionParser()  


    parser.add_option("-d", default="n",dest="default",help='是否默认执行进程，请选择 "y" or "n"')
    parser.add_option("-l", default="*",dest="list",help="输出展示的列，默认使用*")
    parser.add_option("-p", default="./log/",dest="path",help="日志文件路径")
    parser.add_option("-i", default="",dest="id",help="日志id")
    parser.add_option("-f", default="",dest="filter",help="筛选条件")


    options, args = parser.parse_args()

    default = options.default
    list = options.list
    path = options.path
    id = options.id
    filter = options.filter


    if list == "*" and path == "./log/" and id == "" and filter == "":
        default = "y"
    
    if default == "y":

        default_id()

        if os.path.exists("./result/4624.xml"):
            logID.id_4624()
        if os.path.exists("./result/4625.xml"):
            logID.id_4625()
        if os.path.exists("./result/1102.xml"):
            logID.id_1102()
        if os.path.exists("./result/1149.xml"):
            logID.id_1149()
        if os.path.exists("./result/6005.xml"):
            logID.id_6005()
        if os.path.exists("./result/25.xml"):
            logID.id_25()
        if os.path.exists("./result/104.xml"):
            logID.id_104()
        if os.path.exists("./result/4720.xml"):
            logID.id_4720()


        # 获取目录下所有文件名
        file_names = os.listdir("./result/")

        # 遍历文件列表，逐个删除后缀为 .xml 的文件
        for file_name in file_names:
            # 判断文件名是否以 .xml 结尾
            if file_name.endswith('.xml'):
                file_path = os.path.join("./result/", file_name)
                try:
                    os.remove(file_path)
                except OSError as e:
                    print(f"Error deleting file: {file_path} - {e}")


        print("默认执行成功，输出请查看result文件夹")
    else:
        if path == "./log/":
            show(list=list,path="./log/Security.evtx",id=id,filter=filter)
        else:
            show(list=list,path=path,id=id,filter=filter)
