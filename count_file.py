# _*_ coding: utf-8 _*_
import os
import numpy as np
import matplotlib.pyplot as plt

path = r'G:/GetFile/163/'
count = 0
lists = []
for i in os.listdir(path):
    count = count + 1
    lists.append(i)
print(lists)
print('总共有%s名联系人' % (count))
email_count = []
for root,dirs,files in os.walk(path):

    for dir in dirs:
        new_path = os.path.join(path,dir)
        cnt = 0
        for j in os.listdir(new_path):
            cnt = cnt + 1
        email_count.append(cnt)
        print('%s总共发送了%s封附件' % (dir,cnt))
print(email_count)

n = len(email_count)
Y = email_count
fig, ax = plt.subplots()
index = np.arange(n)
bar_width = 0.35
opacity = 0.4
rects1 = plt.bar(index,Y,bar_width,alpha = opacity,color='b',label='email_count')

plt.xlabel('linkman')
plt.ylabel('numbers')
plt.title('number of attachments')
plt.xticks(index ,(lists))
plt.ylim(0,50)
plt.legend()

plt.tight_layout()
plt.show()









