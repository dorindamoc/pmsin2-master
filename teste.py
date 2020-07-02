

'''

my_list = ['1','a', 'b', 'c', '1', 'd', 'e', 'f', '1', 'g', 'h']

sphs = [x for x,y in enumerate(my_list) if y == '1']
sphs.append(len(my_list))
plist = [my_list[x+1:y] for x, y in zip(sphs, sphs[1:])]




chk = False

if chk:
    print('stot')
else:
    print('E fals')






from classes import site

s = site('ROSPA0061')

print(s.bf())

'''
'''
import pandas as pd
from io import BytesIO
import requests


url1 = 'https://docs.google.com/spreadsheets/d/1awwKOi1lY8KTQeRPp3D7BS_mLB3oBXrb98VgyJyvWeY/edit#gid=1943342272'
url2 = 'https://docs.google.com/spreadsheets/d/1awwKOi1lY8KTQeRPp3D7BS_mLB3oBXrb98VgyJyvWeY/export?format=csv&gid=1943342272'

df=pd.read_csv(url2)

print(df.head())


r = requests.get(url1)
data = r.content

df2 = pd.read_csv(BytesIO(data))


print(df2.head())

'''



'''
from classes import site

tst = site('ROSPA0061')



dm = tst.master('dummy_data.xlsx')
di = tst.impacts('dummy_data.xlsx')
dms = tst.masuri('dummy_data.xlsx')
dd = tst.descrieri('dummy_data.xlsx')

#print(dm.head())
#print(dm.tail())

print(di.head())
print(di.tail())

#print(dms.head())
#print(dms.tail())

#print(dd.head())
#print(dd.tail())
'''


'''
xs = '5322.0'

xf = float(xs)

xi = int(float(xs))


print(xi)
'''

'''
import pandas as pd

letters = ['a','d','f','g','h']
nrs1 = [1,2,3,4,5,]
nrs2 = [x for x in range(5,10)]
nrs3 = [x for x in range(15,20)]

dfd = { 'ltrs':letters, 'n1':nrs1, 'n2':nrs2, 'n3':nrs3 }


df = pd.DataFrame(dfd)


c1 = df['n1'].isin([2,3,4])
c2 = df['n2'].isin([5,6,7])




tst = next(iter(df[c1&c2]['ltrs']),'')


print(tst)
'''

import sys
import time

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QDialog,
                             QProgressBar, QPushButton)

TIME_LIMIT = 100

class External(QThread):
    """
    Runs a counter thread.
    """
    countChanged = pyqtSignal(int)

    def run(self):
        count = 0
        while count < TIME_LIMIT:
            count +=1
            time.sleep(1)
            self.countChanged.emit(count)

class Actions(QDialog):
    """
    Simple dialog that consists of a Progress Bar and a Button.
    Clicking on the button results in the start of a timer and
    updates the progress bar.
    """
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Progress Bar')
        self.progress = QProgressBar(self)
        self.progress.setGeometry(0, 0, 300, 25)
        self.progress.setMaximum(100)
        self.button = QPushButton('Start', self)
        self.button.move(0, 30)
        self.show()

        self.button.clicked.connect(self.onButtonClick)

    def onButtonClick(self):
        self.calc = External()
        self.calc.countChanged.connect(self.onCountChanged)
        self.calc.start()

    def onCountChanged(self, value):
        self.progress.setValue(value)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Actions()
    sys.exit(app.exec_())