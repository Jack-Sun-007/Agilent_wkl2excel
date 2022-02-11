import os
import pandas as pd
from xml.dom.minidom import parse


def get_wkl_to_excel(file):
    dom_tree = parse(file)
    # 文档根元素
    root = dom_tree.documentElement
    sample_info = ['Name', 'SamplePosition', 'AcqMethod', 'DAMethod', 'DataFileName', 'SampleType', 'CalibLevelName', 'InjectionVolume', 'Description', 'SampleGroup', 'SampleInformation', 'Identifier', 'RackCode', 'RackPosition', 'PlateCode', 'PlatePosition', 'MethodExecutionType', 'BalanceType']
    table = dict()
    # 遍历所需的节点
    for node in sample_info:
        roots = root.getElementsByTagName(node)
        info = list()
        d = 0
        for i in range(len(roots)):
            root_text = roots[i].firstChild.data
            # 空值里有换行，进行替换
            if '\n' in root_text:
                root_text = None
            else:
                pass
            # DataFileName前2个值不在样品里
            if node == 'DataFileName':
                d += 1
                if d > 2:
                    info.append(root_text)
            # MethodExecutionType前1个值不在样品里
            elif node == 'MethodExecutionType':
                d += 1
                if d > 1:
                    info.append(root_text)
            # Description前1个值不在样品里
            elif node == 'Description':
                d += 1
                if d > 1:
                    info.append(root_text)
            # InjectionVolume -1表示进样体积同方法
            elif node == 'InjectionVolume':
                if root_text == '-1':
                    root_text = 'As Method'
                info.append(root_text)
            else:
                info.append(root_text)
        table[node] = info
    # 最后一列判断有没有待机script
    if not root.getElementsByTagName('ScriptInfo'):
        table['ScriptInfo'] = 'No Script'
    else:
        table['ScriptInfo'] = 'The Last Sample Have Script'
    df = pd.DataFrame.from_dict(table)
    df.to_excel(file[:-4] + '.xlsx', index=False)
    pass


def main():
    for files in os.listdir(os.getcwd()):
        if files.endswith('.wkl'):
            get_wkl_to_excel(files)


if __name__ == '__main__':
    main()
