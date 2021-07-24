import subprocess        # import文なので次以降の例では省略します
import csv


subprocess.run([r'C:/Program Files/WinMerge/WinMergeU.exe', \
    'target/Before', \
    'target/After', \
	'-minimize', \
	'-noninteractive', \
	'-noprefs', \
	'-cfg', 'Settings/DirViewExpandSubdirs=1', \
	'-cfg', 'ReportFiles/ReportType=0', \
	'-cfg', 'ReportFiles/IncludeFileCmpReport=1', \
	'-r', \
	'-u', \
	'-or', 'target/out.csv'])

with open('target/out.csv') as f:
    reader = csv.reader(f)
    for row in reader:
        print(row)

        