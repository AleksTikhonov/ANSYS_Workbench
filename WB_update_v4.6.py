# encoding: utf-8
# 2019 R3
SetScriptVersion(Version="19.5.112")

import os
import shutil
import re
import logging
from sys import platform


logging.basicConfig(filename='python.log',filemode='w',level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%H:%M:%S')

wdpath = os.getcwd()

for f in os.listdir(wdpath):
    if f.endswith(".wbpz"):
        projectname = os.path.splitext(f)[0]
        break

if "win" in platform.lower():
    slash = '\\'
else:
    slash = '/'

projectfiledirpath = wdpath + slash + projectname + '_files'
projecinsidepath = wdpath + slash + projectname + '_files' + slash + 'dp0'
projectfilepath = wdpath + slash + projectname + '.wbpj'
projectarchivepath = wdpath + slash + projectname + '.wbpz'
csvpath = wdpath + slash + projectname + '.csv'
pattern = r"(P\d{1,3})\s-\s([a-zA-Z _.@0-9]+)\s?(\[.+?\])*"
pattern_2 = r"DP\s.*"
output_file_name = wdpath + slash + 'values' + slash + 'values.csv'
res = []
val = {}


def process_file(csvpath):
    with open(csvpath, "r") as f:
        for line in f:
            _res = re.findall(pattern, line)
            if len(_res) > 0:
                for _r in _res:
                    res.append(list(_r))
            _r = re.match(pattern_2, line)
            if _r is not None:
                _val = line.strip().split(',')
                _cnt = 0
                for v in _val[1:]:
                    res[_cnt].append(v)
                    _cnt += 1


def save_output():
    cnt = 1
  
    with open(output_file_name, 'w') as f:
        f.write("#,Name,Value,Target,Condition,Dimension,Description,Comments,\n")
        for e in res:
            a = str(cnt) + ',' + str(e[1].strip()) + ',' + str(e[3]) + ',' + ',' + ',' + str(e[2]) + ',' + ',' + ','
            f.write(a)
            f.write("\n")
            cnt += 1


def save_results():
    
    for root, dirs, files in os.walk(projectfiledirpath):
        for file in files:
            if file.endswith('.res'):
                shutil.copy(os.path.join(root, file), os.path.join(wdpath + slash + 'main_results' + slash + 'cfd', file))
                export_cfd_report()
            if file.endswith('.gtm') or file.endswith('.mshdb'):
                shutil.copy(os.path.join(root, file), os.path.join(wdpath + slash + 'main_results' + slash + 'cfd' + slash + 'mesh', file))


def export_report():
    ExportReport(FilePath=wdpath + slash + 'main_results' + slash + 'report.html')
    for root, dirs, files in os.walk(projecinsidepath):
        for file in files:
            if file.endswith('.png') or file.endswith('.descr'):
                shutil.copy(os.path.join(root, file), os.path.join(wdpath + slash + 'pictures', file))
    try:
        shutil.copy(wdpath + slash + 'main_results' + slash + 'report_images' + slash + 'ProjectSchematic.png', wdpath + slash + 'pictures' + slash + 'model_overview.png')
    except:
        pass


def export_cfd_report():
    for dir in os.listdir(projecinsidepath):
        if dir.startswith('CFX') or dir.startswith('Post'):
            dir = dir.replace('-', ' ')
            system1 = GetSystem(Name=dir)
            results1 = system1.GetContainer(ComponentName="Results")
            cfdResults1 = results1.GetCFDResults()
            for fl in os.listdir(wdpath):
                if fl.endswith(".cse"):
                    cfdResults1.LoadReport = "Custom"
                    cfdResults1.CustomReportTemplate = wdpath + slash + fl
                    break
                else:
                    cfdResults1.LoadReport = "GenericReport"
            cfdResults1.GenerateReport = True
            cfdResults1.ReportLocationDirectory = wdpath + slash + 'main_results' + slash + 'cfd'
            resultsComponent1 = system1.GetComponent(Name="Results")
            resultsComponent1.Update(AllDependencies=True)
    
    for root, dirs, files in os.walk(wdpath + slash + 'main_results' + slash + 'cfd'):
        for file in files:
            if file.endswith('.png'):
                shutil.copy(os.path.join(root, file), os.path.join(wdpath + slash + 'pictures', file))


def make_dirs():
    try:
        os.mkdir(wdpath + slash + 'values')
    except OSError:
        pass
    try:
        os.mkdir(wdpath + slash + 'main_results')
    except OSError:
        pass
    try:
        os.mkdir(wdpath + slash + 'pictures')
    except OSError:
        pass
    try:
        os.mkdir(wdpath + slash + 'main_results' + slash + 'cfd')
    except OSError:
        pass
    try:
        os.mkdir(wdpath + slash + 'main_results' + slash + 'cfd' + slash + 'mesh')
    except OSError:
        pass

def change_parameters():
    designPoint1 = Parameters.GetDesignPoint(Name="0")
    for inp in os.listdir(wdpath):
        if inp.startswith('input') and inp.endswith('.txt'):
            f = open(wdpath + slash + inp)
            for line in f.readlines():
                a = line.split()
                parameter1 = Parameters.GetParameter(Name=a[0])
                try:
                    designPoint1.SetParameterExpression(
                    Parameter=parameter1,
                    Expression=a[2] + " " + a[3])
                except:
                    designPoint1.SetParameterExpression(
                    Parameter=parameter1,
                    Expression=a[2])
            f.close()

Unarchive(
    ArchivePath=projectarchivepath,
    ProjectPath=projectfilepath,
    Overwrite=True)

Save(Overwrite=True)
change_parameters()
Update()
Parameters.ExportAllDesignPointsData(FilePath=csvpath)
make_dirs()
export_report()
export_cfd_report()
save_results()

Reset()

try:
    process_file(csvpath)
    save_output()
    os.remove(csvpath)
except:
    pass

os.remove(projectfilepath)
os.remove(projectarchivepath)
shutil.rmtree(projectfiledirpath)

logging.info("ALL DONE")
